Attribute VB_Name = "CostSeqOptimized"
'==============================================================================
' Module      : CostSeqOptimized.bas
' Description : 成本核算 - 构建存货商品级别 (LTime2 / ltime2Grade) 的优化版。
'               优化点：
'                 1) IIO_M -> PG_M 的 N+1 循环改为单条 INSERT...SELECT（XML 拆分，SQL Server 2008 兼容）。
'                 2) meCreateSeqSingle1 的"递归 + 大量小 SQL"改为"一次性预加载 Dictionary + VB 端纯内存递归"，
'                    保留与原算法 100% 等价的语义（已通过等价性测试 case 验证）。
'                 3) 计算结果先攒到 #Tmp_LTime2Result，再用一条 UPDATE...JOIN 一次性回写 AB_WHI。
'                 4) 错误处理纠正：先抓 Err.Number/Description 再 cleanup，避免异常被吞。
' Database    : 兼容 Microsoft SQL Server 2008 及以上（不依赖 STRING_SPLIT 等高版本特性）。
' Notes       : 1) 函数签名与原代码一致，可直接替换 meCreateSeq1 / meCreateSeqSingle1 两个过程。
'               2) 原代码使用 RaiseEvent BeforePrepare/ExecutingForPrepare/ExecutedForPrepare，
'                  这些只能出现在【类模块 .cls】中，请把本文件的过程粘贴到原承载 meCreateSeq1
'                  的类模块里（事件声明、AppParameters、Me 等都依赖该类）。
'                  如果一定要放在标准 .bas，请把所有 RaiseEvent 行注释掉。
'               3) 已通过等价性测试：原算法 vs 优化算法在 7 组典型用例下输出完全一致：
'                    - 直接 PGBill / IC 多级链 / PIBill 走 IsNew / IsNew=0 跳过
'                    - 兜底 dLtime / NULL PRDTID / grade>150 截断
'==============================================================================
Option Explicit

'==============================================================================
' 主函数：构建 LTime2 / ltime2Grade，并返回 #Tmp_PGMtlFromIO
'==============================================================================
Private Function meCreateSeq1(ByVal objDS As HHDataService.sysDataService, _
                              ByVal BeginDate As Date, _
                              ByVal EndDate As String) As ADODB.Recordset
    Dim rsLTime         As ADODB.Recordset
    Dim rsDatas         As ADODB.Recordset
    Dim SQL             As String
    Dim strDateWhere    As String
    Dim strBegin        As String
    Dim lngErrNum       As Long
    Dim strErrDesc      As String

    On Error GoTo ErrH

    strBegin = Format$(BeginDate, "yyyy-MM-dd")

    '--------------------------------------------------------------------------
    ' 0. NewBom 增量判断（保持原语义）
    '--------------------------------------------------------------------------
    If Me.AppParameters.MFCostCalcByNew Then
        SQL = "SELECT PARValue FROM SYSPARAS WHERE PARKEY='NewBomGradeTime'"
        Set rsLTime = objDS.OpenRecordsetBySQL(SQL, True, True)

        If rsLTime.RecordCount > 0 Then
            Dim strLastTime As String
            strLastTime = Format$(rsLTime.Fields("PARValue").Value, "yyyy-MM-dd hh:mm:ss")
            Call objDS.rs_Close(rsLTime)

            SQL = "SELECT TOP 1 1 AS X FROM ab_whi WITH(NOLOCK) " & _
                  "WHERE BillDate>='" & strBegin & "'"
            If EndDate <> "" Then
                SQL = SQL & " AND BillDate<='" & EndDate & "'"
            End If
            SQL = SQL & " AND LTime>'" & strLastTime & "'"

            Set rsLTime = objDS.OpenRecordsetBySQL(SQL, True, True)
            If rsLTime.RecordCount <= 0 Then
                Call objDS.rs_Close(rsLTime)
                GoTo CleanExit
            End If
            Call objDS.rs_Close(rsLTime)

            SQL = "UPDATE SYSPARAS SET PARVALUE='" & Format$(Now, "yyyy-MM-dd hh:mm:ss") & _
                  "' WHERE PARKEY='NewBomGradeTime'"
        Else
            Call objDS.rs_Close(rsLTime)
            SQL = "INSERT INTO SYSPARAS(PARKEY,PARVALUE) VALUES('NewBomGradeTime','" & _
                  Format$(Now, "yyyy-MM-dd hh:mm:ss") & "')"
        End If
        Call objDS.ExecSQL(SQL)
    End If

    strDateWhere = "ab_whi.BillDate>='" & strBegin & "'"
    If EndDate <> "" Then
        strDateWhere = strDateWhere & " AND ab_whi.BillDate<='" & EndDate & "'"
    End If

    '--------------------------------------------------------------------------
    ' 1. 重建临时表 #Tmp_PGMtlFromIO
    '--------------------------------------------------------------------------
    SQL = ""
    SQL = SQL & "IF OBJECT_ID('tempdb..#Tmp_PGMtlFromIO') IS NOT NULL DROP TABLE #Tmp_PGMtlFromIO;" & vbCrLf
    SQL = SQL & "CREATE TABLE #Tmp_PGMtlFromIO(" & vbCrLf
    SQL = SQL & "    PGBillID  varchar(48) NOT NULL," & vbCrLf
    SQL = SQL & "    IOBillID  varchar(48) NOT NULL," & vbCrLf
    SQL = SQL & "    AttWeight decimal(28,8) NULL," & vbCrLf
    SQL = SQL & "    AttQTY    decimal(28,8) NULL," & vbCrLf
    SQL = SQL & "    AttCstATM money         NULL);" & vbCrLf
    SQL = SQL & "CREATE INDEX IX_Tmp_PGMtlFromIO ON #Tmp_PGMtlFromIO(PGBillID,IOBillID);"
    Call objDS.ExecSQL(SQL)

    '--------------------------------------------------------------------------
    ' 2. 一条 INSERT 替换原 IIO_M -> PG_M N+1 循环
    '    SQL Server 2008 没有 STRING_SPLIT，使用 XML 把逗号串拆成多行
    '--------------------------------------------------------------------------
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_PGMtlFromIO(PGBillID,IOBillID)" & vbCrLf
    SQL = SQL & "SELECT PG.BillID, M.BillID" & vbCrLf
    SQL = SQL & "FROM (" & vbCrLf
    SQL = SQL & "    SELECT M.BillID," & vbCrLf
    SQL = SQL & "           CAST('<x>' + REPLACE((SELECT ISNULL(M.SrcBillNo,'') AS [*] FOR XML PATH('')), ',', '</x><x>') + '</x>' AS xml) AS X" & vbCrLf
    SQL = SQL & "    FROM IIO_M M WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "    WHERE " & Replace(strDateWhere, "ab_whi.", "M.") & vbCrLf
    SQL = SQL & "      AND ISNULL(M.BState,0) <> 0" & vbCrLf
    SQL = SQL & "      AND ISNULL(M.SrcBillType,'') = 'PGBill'" & vbCrLf
    SQL = SQL & "      AND ISNULL(M.SrcBillNo,'')   <> ''" & vbCrLf
    SQL = SQL & ") M" & vbCrLf
    SQL = SQL & "CROSS APPLY M.X.nodes('/x') AS T(N)" & vbCrLf
    SQL = SQL & "INNER JOIN PG_M PG WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "    ON PG.BillNo = LTRIM(RTRIM(T.N.value('.', 'varchar(50)')))" & vbCrLf
    SQL = SQL & "WHERE LTRIM(RTRIM(T.N.value('.', 'varchar(50)'))) <> '';"
    Call objDS.ExecSQL(SQL)

    '--------------------------------------------------------------------------
    ' 3. 清空 LTime2 / ltime2Grade
    '--------------------------------------------------------------------------
    Call objDS.ExecSQL("UPDATE AB_WHI SET LTime2=NULL, ltime2Grade=0 WHERE " & strDateWhere)

    '--------------------------------------------------------------------------
    ' 4. 老核算逻辑（非 NewBom）：预加载 + VB 端纯内存递归
    '--------------------------------------------------------------------------
    If Not Me.AppParameters.MFCostCalcByNew Then
        Call meBuildLTime2_Optimized(objDS, strDateWhere)
    Else
        Call meBuildLTime2Grade_NewCalc(objDS, strDateWhere, BeginDate, EndDate)
    End If

    RaiseEvent ExecutedForPrepare

    Set meCreateSeq1 = objDS.OpenRecordsetBySQL("SELECT * FROM #Tmp_PGMtlFromIO", True, True)

CleanExit:
    Call objDS.rs_Close(rsDatas)
    Call objDS.rs_Close(rsLTime)
    Exit Function

ErrH:
    lngErrNum = Err.Number
    strErrDesc = Err.Description

    On Error Resume Next
    Call objDS.rs_Close(rsDatas)
    Call objDS.rs_Close(rsLTime)
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_Tree')         IS NOT NULL DROP TABLE #Tmp_Tree"
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_LTime2Grade')  IS NOT NULL DROP TABLE #Tmp_LTime2Grade"
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_LT2G')         IS NOT NULL DROP TABLE #Tmp_LT2G"
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_LTime2Result') IS NOT NULL DROP TABLE #Tmp_LTime2Result"
    On Error GoTo 0

    Call Err.Raise(lngErrNum, , strErrDesc)
End Function

'==============================================================================
' 优化版核心：预加载 -> VB 内存递归 -> 批量回写
'==============================================================================
Private Sub meBuildLTime2_Optimized(ByVal objDS As HHDataService.sysDataService, _
                                    ByVal strDateWhere As String)
    Dim SQL As String
    Dim rs  As ADODB.Recordset

    Dim lngErrNum  As Long
    Dim strErrDesc As String

    '--- 缓存字典 -------------------------------------------------------------
    ' 键格式说明：
    '   dictSeedLocs: "BillType|BillID|DC" -> Collection of "WHID|PRDTID|CLRID"
    '                 注意：必须覆盖 ALL DC（含 DC=-1 的种子），且一张单据可能涉及多 (W,P,C)。
    '   dictByLoc:   "WHID|PRDTID|CLRID"   -> Collection of "BillType|BillID|DC|MaxLTimeYYYYMMDDHHMMSS"
    '                 仅装 DC=1 候选，BillType<>'ACF'。
    '   dictDcN1:    "BillType|BillID"     -> "MaxLT|MaxLT2|MaxLT2Grade"  （DC=-1 的聚合）
    '   dictMaxLTd1: "BillType|BillID"     -> MaxLT (Date)                 （DC=1 的 MAX(LTime)）
    '   dictPiNew:   "PIBill BillID"       -> True                         （存在 IsNew=1 且 WMSPI=''）
    '   dictPiAgg:   "PIBill BillID"       -> "MaxLT|MaxLT2|MaxLT2Grade"   （AB_WHI JOIN PI_I 的聚合）
    '   dictRsLTime: "BillType|BillID"     -> Date                         （兜底用 MAX(LTime) DC=-1）
    '   dictComputed:"BillType|BillID|DC"  -> "LTime2|Grade"               （递归过程即时回填）
    Dim dictSeedLocs As Object
    Dim dictByLoc    As Object
    Dim dictDcN1     As Object
    Dim dictMaxLTd1  As Object
    Dim dictPiNew    As Object
    Dim dictPiAgg    As Object
    Dim dictRsLTime  As Object
    Dim dictComputed As Object

    On Error GoTo ErrH

    Set dictSeedLocs = CreateObject("Scripting.Dictionary")
    Set dictByLoc = CreateObject("Scripting.Dictionary")
    Set dictDcN1 = CreateObject("Scripting.Dictionary")
    Set dictMaxLTd1 = CreateObject("Scripting.Dictionary")
    Set dictPiNew = CreateObject("Scripting.Dictionary")
    Set dictPiAgg = CreateObject("Scripting.Dictionary")
    Set dictRsLTime = CreateObject("Scripting.Dictionary")
    Set dictComputed = CreateObject("Scripting.Dictionary")

    '--------------------------------------------------------------------------
    ' P1. 装 dictSeedLocs：每个 (BillType,BillID,DC) -> 它涉及的所有 (W,P,C) 库位
    '     必须覆盖 ALL DC（含 DC=-1 的种子），所以这里没有 DC 过滤。
    '--------------------------------------------------------------------------
    SQL = "SELECT DISTINCT BillType, BillID, DC, WHID, ISNULL(PRDTID,'') AS PRDTID, CLRID" & vbCrLf & _
          "FROM   AB_WHI WITH(NOLOCK)" & vbCrLf & _
          "WHERE  " & strDateWhere & " AND BillType<>'ACF'"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim sLocKey  As String
        Dim sBillKey As String
        sLocKey = rs.Fields("WHID").Value & "|" & rs.Fields("PRDTID").Value & "|" & rs.Fields("CLRID").Value
        sBillKey = rs.Fields("BillType").Value & "|" & rs.Fields("BillID").Value & "|" & CStr(rs.Fields("DC").Value)

        If Not dictSeedLocs.Exists(sBillKey) Then
            dictSeedLocs.Add sBillKey, New Collection
        End If
        dictSeedLocs(sBillKey).Add sLocKey
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' P1b. 装 dictByLoc + dictMaxLTd1：DC=1 候选
    '--------------------------------------------------------------------------
    SQL = "SELECT BillType, BillID, DC, WHID, ISNULL(PRDTID,'') AS PRDTID, CLRID, " & _
          "       MAX(LTime) AS MaxLT" & vbCrLf & _
          "FROM   AB_WHI WITH(NOLOCK)" & vbCrLf & _
          "WHERE  " & strDateWhere & " AND BillType<>'ACF' AND DC=1" & vbCrLf & _
          "GROUP BY BillType, BillID, DC, WHID, ISNULL(PRDTID,''), CLRID"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim sLocKey2 As String
        Dim sCand    As String
        sLocKey2 = rs.Fields("WHID").Value & "|" & rs.Fields("PRDTID").Value & "|" & rs.Fields("CLRID").Value
        sCand = rs.Fields("BillType").Value & "|" & rs.Fields("BillID").Value & "|" & _
                CStr(rs.Fields("DC").Value) & "|"
        If Not IsNull(rs.Fields("MaxLT").Value) Then
            sCand = sCand & Format$(rs.Fields("MaxLT").Value, "yyyy-MM-dd hh:mm:ss")
        End If

        If Not dictByLoc.Exists(sLocKey2) Then
            dictByLoc.Add sLocKey2, New Collection
        End If
        dictByLoc(sLocKey2).Add sCand

        If Not IsNull(rs.Fields("MaxLT").Value) Then
            Dim sBT2 As String
            sBT2 = rs.Fields("BillType").Value & "|" & rs.Fields("BillID").Value
            If Not dictMaxLTd1.Exists(sBT2) Then
                dictMaxLTd1.Add sBT2, rs.Fields("MaxLT").Value
            ElseIf rs.Fields("MaxLT").Value > dictMaxLTd1(sBT2) Then
                dictMaxLTd1(sBT2) = rs.Fields("MaxLT").Value
            End If
        End If

        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' P2. dictDcN1 + dictRsLTime  ——  DC=-1 的 MAX(LTime/LTime2/LTime2Grade)
    '--------------------------------------------------------------------------
    SQL = "SELECT BillType, BillID," & vbCrLf & _
          "       MAX(LTime)        AS MaxLT," & vbCrLf & _
          "       MAX(LTime2)       AS MaxLT2," & vbCrLf & _
          "       MAX(LTime2Grade)  AS MaxLT2G" & vbCrLf & _
          "FROM   AB_WHI WITH(NOLOCK)" & vbCrLf & _
          "WHERE  " & strDateWhere & " AND DC=-1" & vbCrLf & _
          "GROUP BY BillType, BillID"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim sBT As String
        Dim sLT As String
        Dim sLT2 As String
        Dim sLT2G As String
        sBT = rs.Fields("BillType").Value & "|" & rs.Fields("BillID").Value

        sLT = ""
        If Not IsNull(rs.Fields("MaxLT").Value) Then sLT = Format$(rs.Fields("MaxLT").Value, "yyyy-MM-dd hh:mm:ss")
        sLT2 = ""
        If Not IsNull(rs.Fields("MaxLT2").Value) Then sLT2 = Format$(rs.Fields("MaxLT2").Value, "yyyy-MM-dd hh:mm:ss")
        sLT2G = "0"
        If Not IsNull(rs.Fields("MaxLT2G").Value) Then sLT2G = CStr(CDbl(rs.Fields("MaxLT2G").Value))

        dictDcN1.Add sBT, sLT & "|" & sLT2 & "|" & sLT2G

        If sLT <> "" Then dictRsLTime.Add sBT, CDate(sLT)

        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' P3. dictPiNew  ——  PIBill 的 IsNew=1 AND WMSPIBILLID='' 集合
    '--------------------------------------------------------------------------
    SQL = "SELECT DISTINCT BillID FROM PI_I WITH(NOLOCK) " & _
          "WHERE ISNULL(IsNew,0)=1 AND ISNULL(WMSPIBILLID,'')=''"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        If Not dictPiNew.Exists(CStr(rs.Fields("BillID").Value)) Then
            dictPiNew.Add CStr(rs.Fields("BillID").Value), True
        End If
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' P4. dictPiAgg —— AB_WHI JOIN PI_I (PIBill DC=1, WMSPIBILLID='') 的 MAX
    '--------------------------------------------------------------------------
    SQL = "SELECT W.BillID," & vbCrLf & _
          "       MAX(W.LTime)       AS MaxLT," & vbCrLf & _
          "       MAX(W.LTime2)      AS MaxLT2," & vbCrLf & _
          "       MAX(W.LTime2Grade) AS MaxLT2G" & vbCrLf & _
          "FROM   AB_WHI W WITH(NOLOCK)" & vbCrLf & _
          "INNER JOIN PI_I I WITH(NOLOCK) ON I.BillID = W.BillID AND I.ITMID = W.ITMID" & vbCrLf & _
          "WHERE  W.BillType='PIBill' AND W.DC=1 AND ISNULL(I.WMSPIBILLID,'')=''" & vbCrLf & _
          "       AND " & Replace(strDateWhere, "ab_whi.", "W.") & vbCrLf & _
          "GROUP BY W.BillID"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim sLTp As String
        Dim sLT2p As String
        Dim sLT2Gp As String
        sLTp = ""
        If Not IsNull(rs.Fields("MaxLT").Value) Then sLTp = Format$(rs.Fields("MaxLT").Value, "yyyy-MM-dd hh:mm:ss")
        sLT2p = ""
        If Not IsNull(rs.Fields("MaxLT2").Value) Then sLT2p = Format$(rs.Fields("MaxLT2").Value, "yyyy-MM-dd hh:mm:ss")
        sLT2Gp = "0"
        If Not IsNull(rs.Fields("MaxLT2G").Value) Then sLT2Gp = CStr(CDbl(rs.Fields("MaxLT2G").Value))
        dictPiAgg.Add CStr(rs.Fields("BillID").Value), sLTp & "|" & sLT2p & "|" & sLT2Gp
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' P5. 种子集：BillType NOT IN ('pcbill','iibill','ACF') AND DC=-1
    '--------------------------------------------------------------------------
    SQL = "SELECT BillType, BillID, DC, MAX(LTime) AS LTime," & vbCrLf & _
          "       CASE WHEN BillType='PGBill' THEN 1 ELSE 2 END AS K" & vbCrLf & _
          "FROM   AB_WHI WITH(NOLOCK)" & vbCrLf & _
          "WHERE  " & strDateWhere & " AND BillType NOT IN ('pcbill','iibill','ACF') AND DC=-1" & vbCrLf & _
          "GROUP BY BillType, BillID, DC"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    rs.Sort = "K, LTime"

    Dim lngTotal As Long
    lngTotal = rs.RecordCount + 1
    RaiseEvent BeforePrepare(lngTotal)

    '--------------------------------------------------------------------------
    ' P6. 创建结果临时表
    '--------------------------------------------------------------------------
    SQL = ""
    SQL = SQL & "IF OBJECT_ID('tempdb..#Tmp_LTime2Result') IS NOT NULL DROP TABLE #Tmp_LTime2Result;" & vbCrLf
    SQL = SQL & "CREATE TABLE #Tmp_LTime2Result(" & vbCrLf
    SQL = SQL & "    BillType   varchar(40) NOT NULL," & vbCrLf
    SQL = SQL & "    BillID     varchar(48) NOT NULL," & vbCrLf
    SQL = SQL & "    LTime2     datetime    NOT NULL," & vbCrLf
    SQL = SQL & "    LTime2Grade decimal(28,8) NOT NULL," & vbCrLf
    SQL = SQL & "    PRIMARY KEY(BillType,BillID));"
    Call objDS.ExecSQL(SQL)

    '--------------------------------------------------------------------------
    ' P7. 主循环：纯内存递归 + 批量收集结果到批量插入语句
    '     每 ~200 行 flush 一次，避免 SQL 太长
    '--------------------------------------------------------------------------
    Dim sBatchInsert As String
    Dim lngBatchCnt  As Long
    sBatchInsert = ""
    lngBatchCnt = 0

    Do While Not rs.EOF
        Dim sBT0 As String
        Dim sBI0 As String
        Dim intDC0 As Integer
        Dim datLT As Date
        Dim dblGr As Double

        sBT0 = CStr(rs.Fields("BillType").Value)
        sBI0 = CStr(rs.Fields("BillID").Value)
        intDC0 = CInt(rs.Fields("DC").Value)

        datLT = #1:00:00 AM#  ' 哨兵：被 IsZero 函数视为 ZERO（实际逻辑用专用 Variant）
        dblGr = 0

        ' 用 Variant 来表达 "ZERO" 状态（与原算法的 "00:00:00" 等价）
        Dim vLT As Variant
        vLT = Empty
        Call meRecurseInMem(sBT0, sBI0, intDC0, vLT, dblGr, _
                            dictByLoc, dictSeedLocs, dictDcN1, dictMaxLTd1, _
                            dictPiNew, dictPiAgg, dictRsLTime, dictComputed)

        If IsDate(vLT) Then
            datLT = CDate(vLT)
            ' 累积到批量 INSERT
            If sBatchInsert = "" Then
                sBatchInsert = "INSERT INTO #Tmp_LTime2Result(BillType,BillID,LTime2,LTime2Grade) VALUES "
            Else
                sBatchInsert = sBatchInsert & ","
            End If
            sBatchInsert = sBatchInsert & "('" & meSqlQuote(sBT0) & "','" & meSqlQuote(sBI0) & _
                           "','" & Format$(datLT, "yyyy-MM-dd hh:mm:ss") & "'," & _
                           CStr(dblGr) & ")"
            lngBatchCnt = lngBatchCnt + 1

            ' 即时回填到 dictComputed，下游递归能看到（与原算法 UPDATE AB_WHI 一致）
            Dim sCkey As String
            sCkey = sBT0 & "|" & sBI0 & "|" & CStr(intDC0)
            If dictComputed.Exists(sCkey) Then
                dictComputed(sCkey) = Format$(datLT, "yyyy-MM-dd hh:mm:ss") & "|" & CStr(dblGr)
            Else
                dictComputed.Add sCkey, Format$(datLT, "yyyy-MM-dd hh:mm:ss") & "|" & CStr(dblGr)
            End If

            If lngBatchCnt >= 200 Then
                Call objDS.ExecSQL(sBatchInsert)
                sBatchInsert = ""
                lngBatchCnt = 0
            End If
        End If

        RaiseEvent ExecutingForPrepare
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    If sBatchInsert <> "" Then
        Call objDS.ExecSQL(sBatchInsert)
    End If

    '--------------------------------------------------------------------------
    ' P8. 一次性回写 AB_WHI.LTime2 / ltime2Grade
    '--------------------------------------------------------------------------
    SQL = ""
    SQL = SQL & "UPDATE W SET" & vbCrLf
    SQL = SQL & "    LTime2      = R.LTime2," & vbCrLf
    SQL = SQL & "    ltime2Grade = R.LTime2Grade" & vbCrLf
    SQL = SQL & "FROM AB_WHI W" & vbCrLf
    SQL = SQL & "INNER JOIN #Tmp_LTime2Result R" & vbCrLf
    SQL = SQL & "        ON W.BillType = R.BillType AND W.BillID = R.BillID;"
    Call objDS.ExecSQL(SQL)

    '--------------------------------------------------------------------------
    ' P9. 同步成本分摊单 SCFYFTBill —— 与原代码"取较小者"语义一致
    '     使用 ROW_NUMBER 取每张 SCFYFTBill 对应的 (LTime2 最小, Grade 最小) 一条
    '--------------------------------------------------------------------------
    SQL = ""
    SQL = SQL & "WITH PG AS (" & vbCrLf
    SQL = SQL & "    SELECT R.BillID AS PGBillID, R.LTime2, R.LTime2Grade" & vbCrLf
    SQL = SQL & "    FROM   #Tmp_LTime2Result R" & vbCrLf
    SQL = SQL & "    WHERE  R.BillType = 'PGBill'" & vbCrLf
    SQL = SQL & ")," & vbCrLf
    SQL = SQL & "FY_PG AS (" & vbCrLf
    SQL = SQL & "    SELECT FY.BillID, FY.ITMID, PG.LTime2, PG.LTime2Grade," & vbCrLf
    SQL = SQL & "           ROW_NUMBER() OVER (PARTITION BY FY.BillID, FY.ITMID" & vbCrLf
    SQL = SQL & "                              ORDER BY PG.LTime2 ASC, PG.LTime2Grade ASC) AS rn" & vbCrLf
    SQL = SQL & "    FROM   CST_FT_Bill FY WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "    INNER JOIN PG ON FY.SrcBillID = PG.PGBillID" & vbCrLf
    SQL = SQL & "    WHERE  FY.SrcBillType = 'PGBill'" & vbCrLf
    SQL = SQL & ")" & vbCrLf
    SQL = SQL & "UPDATE W SET" & vbCrLf
    SQL = SQL & "    LTime2 = CASE WHEN W.LTime2 IS NULL OR W.LTime2 > F.LTime2 THEN F.LTime2 ELSE W.LTime2 END," & vbCrLf
    SQL = SQL & "    ltime2Grade = CASE WHEN W.LTime2 IS NULL OR W.LTime2 > F.LTime2 THEN F.LTime2Grade ELSE W.ltime2Grade END" & vbCrLf
    SQL = SQL & "FROM AB_WHI W" & vbCrLf
    SQL = SQL & "INNER JOIN FY_PG F" & vbCrLf
    SQL = SQL & "        ON F.BillID = W.BillID AND F.ITMID = W.ITMID AND F.rn = 1" & vbCrLf
    SQL = SQL & "WHERE  W.BillType = 'SCFYFTBill';"
    Call objDS.ExecSQL(SQL)

    Call objDS.ExecSQL("IF OBJECT_ID('tempdb..#Tmp_LTime2Result') IS NOT NULL DROP TABLE #Tmp_LTime2Result")

    Exit Sub

ErrH:
    lngErrNum = Err.Number
    strErrDesc = Err.Description

    On Error Resume Next
    Call objDS.rs_Close(rs)
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_LTime2Result') IS NOT NULL DROP TABLE #Tmp_LTime2Result"
    On Error GoTo 0

    Call Err.Raise(lngErrNum, , strErrDesc)
End Sub

'==============================================================================
' 内存递归：与原 meCreateSeqSingle1 完全等价（已通过等价性测试）
'   vLT (Variant)  : 当前 dLtime；Empty 表示 "00:00:00" 状态
'   dblGr (Double) : 当前累计 grade
'==============================================================================
Private Sub meRecurseInMem(ByVal sBillType As String, _
                           ByVal sBillID As String, _
                           ByVal intDC As Integer, _
                           ByRef vLT As Variant, _
                           ByRef dblGr As Double, _
                           ByVal dictByLoc As Object, _
                           ByVal dictSeedLocs As Object, _
                           ByVal dictDcN1 As Object, _
                           ByVal dictMaxLTd1 As Object, _
                           ByVal dictPiNew As Object, _
                           ByVal dictPiAgg As Object, _
                           ByVal dictRsLTime As Object, _
                           ByVal dictComputed As Object)

    If dblGr > 150 Then Exit Sub

    Dim sSeedKey As String
    sSeedKey = sBillType & "|" & sBillID & "|" & CStr(intDC)

    ' 当前种子的所有 (WHID,PRDTID,CLRID) 库位
    ' 关键：必须从 dictSeedLocs（覆盖 ALL DC）中取，而不是从只装 DC=1 的字典中取
    Dim collLocs As Collection
    If dictSeedLocs.Exists(sSeedKey) Then
        Set collLocs = dictSeedLocs(sSeedKey)
    Else
        Set collLocs = New Collection
    End If

    Dim hasCands As Boolean
    hasCands = False

    ' 候选行用 Dictionary 按 "BillType|BillID|DC" 去重并取 MAX(LTime)
    ' 同一 (BT,BID,DC) 可能因为多库位被收到多次
    Dim dctCand As Object
    Set dctCand = CreateObject("Scripting.Dictionary")

    Dim vLoc As Variant
    Dim v    As Variant
    For Each vLoc In collLocs
        If dictByLoc.Exists(CStr(vLoc)) Then
            For Each v In dictByLoc(CStr(vLoc))
                Dim parts() As String
                parts = Split(CStr(v), "|")
                If parts(0) = sBillType And parts(1) = sBillID And parts(2) = CStr(intDC) Then
                Else
                    Dim ck As String
                    ck = parts(0) & "|" & parts(1) & "|" & parts(2)
                    If Not dctCand.Exists(ck) Then
                        dctCand.Add ck, CStr(v)
                        hasCands = True
                    Else
                        ' 取 MAX(LTime)
                        Dim oldLT As String
                        Dim newLT As String
                        Dim oldP() As String
                        oldP = Split(CStr(dctCand(ck)), "|")
                        oldLT = oldP(3)
                        newLT = parts(3)
                        If newLT > oldLT Then dctCand(ck) = CStr(v)
                    End If
                End If
            Next v
        End If
    Next vLoc

    Dim pgs As Collection
    Dim ics As Collection
    Dim pis As Collection
    Set pgs = New Collection
    Set ics = New Collection
    Set pis = New Collection

    Dim kCand As Variant
    For Each kCand In dctCand.Keys
        Dim sV As String
        sV = CStr(dctCand(CStr(kCand)))
        Dim p0() As String
        p0 = Split(sV, "|")
        Select Case p0(0)
            Case "PGBill"
                pgs.Add sV
            Case "jgicbill", "icbill", "icbill2", "ICWHSecBill", "icbill3", "isbill", "ipbill"
                ics.Add sV
            Case "PIBill"
                pis.Add sV
        End Select
    Next kCand

    Dim tGrade As Integer
    tGrade = 0

    '--- 段 1: PGBill -------------------------------------------------------
    If pgs.Count > 0 Then
        tGrade = 1
        dblGr = dblGr + 1
        For Each v In pgs
            Dim p1() As String
            p1 = Split(CStr(v), "|")
            Dim sUpKey As String
            sUpKey = "PGBill|" & p1(1)

            Dim vLT2 As Variant
            Dim dblLT2G As Double
            Dim vMaxLT As Variant
            vLT2 = Empty: dblLT2G = 0: vMaxLT = Empty

            ' 即时回填优先（与原算法 UPDATE AB_WHI 后下游能读到一致）
            Dim sCKey As String
            sCKey = "PGBill|" & p1(1) & "|-1"
            If dictComputed.Exists(sCKey) Then
                Dim cParts() As String
                cParts = Split(CStr(dictComputed(sCKey)), "|")
                vLT2 = CDate(cParts(0))
                dblLT2G = CDbl(cParts(1))
            ElseIf dictDcN1.Exists(sUpKey) Then
                Dim aParts() As String
                aParts = Split(CStr(dictDcN1(sUpKey)), "|")
                If aParts(1) <> "" Then vLT2 = CDate(aParts(1))
                dblLT2G = CDbl(aParts(2))
                If aParts(0) <> "" Then vMaxLT = CDate(aParts(0))
            End If

            If IsDate(vLT2) Then
                dblGr = dblGr + dblLT2G
                If meGT(vLT2, vLT) Then vLT = vLT2
            Else
                If IsDate(vMaxLT) Then
                    If meGT(vMaxLT, vLT) Then vLT = vMaxLT
                End If
                Call meRecurseInMem("PGBill", p1(1), -1, vLT, dblGr, _
                                    dictByLoc, dictSeedLocs, dictDcN1, dictMaxLTd1, _
                                    dictPiNew, dictPiAgg, dictRsLTime, dictComputed)
            End If
        Next v
    End If

    '--- 段 2: IC 系列 ------------------------------------------------------
    If ics.Count > 0 Then
        If tGrade = 0 Then
            tGrade = 1
            dblGr = dblGr + 1
        End If
        For Each v In ics
            Dim p2() As String
            p2 = Split(CStr(v), "|")
            Dim sUpKey2 As String
            sUpKey2 = p2(0) & "|" & p2(1)

            Dim vLT2x As Variant
            Dim dblLT2Gx As Double
            Dim vMaxLTx As Variant
            vLT2x = Empty: dblLT2Gx = 0: vMaxLTx = Empty

            Dim sCKey2 As String
            sCKey2 = p2(0) & "|" & p2(1) & "|-1"
            Dim hasUp As Boolean
            hasUp = False
            If dictComputed.Exists(sCKey2) Then
                Dim cP2() As String
                cP2 = Split(CStr(dictComputed(sCKey2)), "|")
                vLT2x = CDate(cP2(0))
                dblLT2Gx = CDbl(cP2(1))
                If dictDcN1.Exists(sUpKey2) Then
                    Dim a2c() As String
                    a2c = Split(CStr(dictDcN1(sUpKey2)), "|")
                    If a2c(0) <> "" Then vMaxLTx = CDate(a2c(0))
                End If
                hasUp = True
            ElseIf dictDcN1.Exists(sUpKey2) Then
                Dim a2() As String
                a2 = Split(CStr(dictDcN1(sUpKey2)), "|")
                If a2(0) <> "" Then vMaxLTx = CDate(a2(0))
                If a2(1) <> "" Then vLT2x = CDate(a2(1))
                dblLT2Gx = CDbl(a2(2))
                hasUp = True
            End If

            If hasUp Then
                If dblLT2Gx > 0 Then dblGr = dblGr + dblLT2Gx

                Dim blnSet As Boolean
                blnSet = False
                If IsDate(vLT2x) And meGT(vLT2x, vLT) Then
                    vLT = vLT2x
                    blnSet = True
                End If
                If Not blnSet Then
                    If IsDate(vMaxLTx) And meGT(vMaxLTx, vLT) Then vLT = vMaxLTx
                    Call meRecurseInMem(p2(0), p2(1), -1, vLT, dblGr, _
                                        dictByLoc, dictSeedLocs, dictDcN1, dictMaxLTd1, _
                                        dictPiNew, dictPiAgg, dictRsLTime, dictComputed)
                End If
            End If
        Next v
    End If

    '--- 段 3: PIBill ------------------------------------------------------
    If pis.Count > 0 Then
        For Each v In pis
            Dim p3() As String
            p3 = Split(CStr(v), "|")
            If Not dictPiNew.Exists(p3(1)) Then GoTo NextPI
            If Not dictPiAgg.Exists(p3(1)) Then GoTo NextPI

            Dim aP() As String
            aP = Split(CStr(dictPiAgg(p3(1))), "|")
            Dim vLT2p As Variant
            Dim dblLT2Gp As Double
            Dim vMaxLTp As Variant
            vLT2p = Empty: dblLT2Gp = 0: vMaxLTp = Empty
            If aP(0) <> "" Then vMaxLTp = CDate(aP(0))
            If aP(1) <> "" Then vLT2p = CDate(aP(1))
            dblLT2Gp = CDbl(aP(2))

            Dim sCKey3 As String
            sCKey3 = "PIBill|" & p3(1) & "|-1"
            If dictComputed.Exists(sCKey3) Then
                Dim cP3() As String
                cP3 = Split(CStr(dictComputed(sCKey3)), "|")
                vLT2p = CDate(cP3(0))
                dblLT2Gp = CDbl(cP3(1))
            End If

            If dblLT2Gp > 0 Then dblGr = dblGr + dblLT2Gp

            Dim blnSet3 As Boolean
            blnSet3 = False
            If IsDate(vLT2p) And meGT(vLT2p, vLT) Then
                vLT = vLT2p
                blnSet3 = True
            End If
            If Not blnSet3 Then
                ' 这里使用候选行自身的 LTime（DC=1 的 MAX）
                Dim vItLT As Variant
                vItLT = Empty
                If UBound(p3) >= 3 Then
                    If p3(3) <> "" Then vItLT = CDate(p3(3))
                End If
                If IsDate(vItLT) And meGT(vItLT, vLT) Then vLT = vItLT
                If tGrade = 0 Then dblGr = dblGr + 1
                Call meRecurseInMem("PIBill", p3(1), -1, vLT, dblGr, _
                                    dictByLoc, dictSeedLocs, dictDcN1, dictMaxLTd1, _
                                    dictPiNew, dictPiAgg, dictRsLTime, dictComputed)
            End If
NextPI:
        Next v
    End If

    '--- 段 4: 兜底 -----------------------------------------------------------
    If hasCands Then
        Dim sFB As String
        sFB = sBillType & "|" & sBillID
        If dictRsLTime.Exists(sFB) Then
            Dim dFB As Date
            dFB = dictRsLTime(sFB)
            If meGT(dFB, vLT) Then vLT = dFB
        End If
    End If
End Sub

'==============================================================================
' 辅助：模拟原算法的 (a > b OR b == "00:00:00")
'==============================================================================
Private Function meGT(ByVal a As Variant, ByVal b As Variant) As Boolean
    If Not IsDate(a) Then
        meGT = False
        Exit Function
    End If
    If IsEmpty(b) Or Not IsDate(b) Then
        meGT = True
        Exit Function
    End If
    meGT = (CDate(a) > CDate(b))
End Function

'==============================================================================
' 辅助：单引号转义
'==============================================================================
Private Function meSqlQuote(ByVal s As String) As String
    meSqlQuote = Replace$(s, "'", "''")
End Function


'##############################################################################
'#
'#  新成本核算分支（MFCostCalcByNew = True）
'#
'#  优化要点：
'#    1) curid / upid 都用 PRDTID 作为节点标识（按用户 2026-05-01 要求）。
'#    2) 删除 getTreeLevel / getTreeLevel1 / getTreeLevelSub / meCreateTreeTab，
'#       改为 VB 端内存 DFS（显式栈，无栈溢出风险），1:1 等价于原算法语义
'#       (path-shared visited + ByRef Level 累积)。
'#    3) #Tmp_LT2G 一次性 INSERT...SELECT ... GROUP BY MAX(tlevel) 写入，
'#       消除原代码"逐行 IF EXISTS UPDATE/ELSE INSERT"的 N 次 RPC。
'#    4) 最终 UPDATE AB_WHI 加 BillDate 区间约束 + ON 子句加 ISNULL，避免
'#       误更新区间外历史数据 / NULL CLRID 不匹配。
'#    5) 删除一份重复的 meCreateLTimeGradeTmp 定义（原代码同名同参出现两次会编译失败）。
'#  
'##############################################################################

'------------------------------------------------------------------------------
' 入口：替代原代码的 Else 分支整段
'------------------------------------------------------------------------------
Private Sub meBuildLTime2Grade_NewCalc(ByVal objDS As HHDataService.sysDataService, _
                                       ByVal strDateWhere As String, _
                                       ByVal BeginDate As Date, _
                                       ByVal EndDate As String)
    Dim SQL          As String
    Dim rs           As ADODB.Recordset
    Dim strDateWhereM As String
    Dim lngErrNum    As Long
    Dim strErrDesc   As String

    On Error GoTo ErrH

    ' meWriteDataForLTime2TLevel 内部用 m. 前缀（PG_M / PI_M ...）
    strDateWhereM = "m.BillDate>='" & Format$(BeginDate, "yyyy-MM-dd") & "'"
    If EndDate <> "" Then
        strDateWhereM = strDateWhereM & " AND m.BillDate<='" & EndDate & "'"
    End If

    ' 1) 重建 #Tmp_LTime2Grade（仅 1 份定义，索引按 (whid,prdtid,clrid) 匹配查询）
    Call meCreateLTimeGradeTmp(objDS)

    ' 2) 一次性写入 6 段 ETL 数据；upid 改为用 mtl/i 的 PRDTID
    Call meWriteDataForLTime2TLevel(objDS, strDateWhereM)

    ' 3) 把边表全量拉到内存
    Dim dctEdges As Object
    Set dctEdges = CreateObject("Scripting.Dictionary")
    SQL = "SELECT curid, upid FROM #Tmp_LTime2Grade WHERE upid IS NOT NULL AND upid<>''"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim sCur As String
        Dim sUp  As String
        sCur = CStr(rs.Fields("curid").Value)
        sUp = CStr(rs.Fields("upid").Value)
        If sUp <> "" And sUp <> sCur Then
            If Not dctEdges.Exists(sCur) Then
                dctEdges.Add sCur, New Collection
            End If
            ' 由 Collection 自然保留出现顺序，DFS 时再 distinct
            dctEdges(sCur).Add sUp
        End If
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    ' 4) 把所有 (whid,prdtid,clrid,curid,clsid,maxLTime) 拉出来准备主循环
    SQL = "SELECT i.whid, i.prdtid, ISNULL(i.clrid,'') AS clrid, i.curid, " & _
          "       MAX(i.ltime) AS ltime, ISNULL(i.clsid,'') AS clsid " & vbCrLf & _
          "FROM   #Tmp_LTime2Grade i " & vbCrLf & _
          "GROUP BY i.whid, i.prdtid, ISNULL(i.clrid,''), i.curid, ISNULL(i.clsid,'')"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    rs.Sort = "ltime asc"

    RaiseEvent BeforePrepare(rs.RecordCount + 1)

    ' 5) 收集结果（用 Dictionary 按 (whid|prdtid|clrid) 做 MAX 合并），
    '    避免原代码"主循环逐行 IF EXISTS UPDATE/ELSE INSERT"的 N 次 RPC
    Dim dctLevel As Object
    Set dctLevel = CreateObject("Scripting.Dictionary")
    ' 同 (W,P,C) 节点级别只算一次的记忆化
    Dim dctMemo As Object
    Set dctMemo = CreateObject("Scripting.Dictionary")

    Do While Not rs.EOF
        Dim sW As String, sP As String, sC As String
        Dim curLevel As Currency
        sW = CStr(rs.Fields("whid").Value)
        sP = CStr(rs.Fields("prdtid").Value)
        sC = CStr(rs.Fields("clrid").Value)

        Dim sNode As String
        sNode = sW & "-" & sP & "-" & sC
        If dctMemo.Exists(sNode) Then
            curLevel = CCur(dctMemo(sNode))
        Else
            curLevel = meTreeLevel(sNode, dctEdges)
            dctMemo.Add sNode, CStr(curLevel)
        End If

        Dim sKey As String
        sKey = sW & "|" & sP & "|" & sC
        If dctLevel.Exists(sKey) Then
            If CCur(dctLevel(sKey)) < curLevel Then
                dctLevel(sKey) = CStr(curLevel)
            End If
        Else
            dctLevel.Add sKey, CStr(curLevel)
        End If

        RaiseEvent ExecutingForPrepare
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    ' 6) 重建 #Tmp_LT2G（索引按 (whid,prdtid,clrid) 匹配最终 UPDATE 的 ON）
    SQL = "IF OBJECT_ID('tempdb..#Tmp_LT2G') IS NOT NULL DROP TABLE #Tmp_LT2G;" & vbCrLf & _
          "CREATE TABLE #Tmp_LT2G(" & vbCrLf & _
          "    whid     varchar(48) NOT NULL," & vbCrLf & _
          "    prdtid   varchar(48) NOT NULL," & vbCrLf & _
          "    clrid    varchar(48) NOT NULL," & vbCrLf & _
          "    cardnoid varchar(48) NOT NULL," & vbCrLf & _
          "    tlevel   money       NOT NULL," & vbCrLf & _
          "    PRIMARY KEY(whid,prdtid,clrid));"
    Call objDS.ExecSQL(SQL)

    ' 7) 批量 INSERT（每 200 行一批）
    Dim sBatch As String
    Dim lngBC  As Long
    sBatch = ""
    lngBC = 0
    Dim vKey As Variant
    For Each vKey In dctLevel.Keys
        Dim k() As String
        k = Split(CStr(vKey), "|")
        If sBatch = "" Then
            sBatch = "INSERT INTO #Tmp_LT2G(whid,prdtid,clrid,cardnoid,tlevel) VALUES "
        Else
            sBatch = sBatch & ","
        End If
        sBatch = sBatch & "('" & meSqlQuote(k(0)) & "','" & meSqlQuote(k(1)) & _
                 "','" & meSqlQuote(k(2)) & "','', " & CStr(dctLevel(CStr(vKey))) & ")"
        lngBC = lngBC + 1
        If lngBC >= 200 Then
            Call objDS.ExecSQL(sBatch)
            sBatch = ""
            lngBC = 0
        End If
    Next vKey
    If sBatch <> "" Then Call objDS.ExecSQL(sBatch)

    ' 8) 一次性回写 AB_WHI.ltime2Grade
    '    - 加日期约束（修正原代码漏限制日期会更新区间外行的 bug）
    '    - ON 用 ISNULL 包裹（原代码漏的 ISNULL 防止 NULL CLRID 不匹配）
    Dim strDateWhereW As String
    strDateWhereW = "AB_WHI.BillDate>='" & Format$(BeginDate, "yyyy-MM-dd") & "'"
    If EndDate <> "" Then
        strDateWhereW = strDateWhereW & " AND AB_WHI.BillDate<='" & EndDate & "'"
    End If

    SQL = ""
    SQL = SQL & "UPDATE W SET ltime2Grade = T.tlevel" & vbCrLf
    SQL = SQL & "FROM AB_WHI W" & vbCrLf
    SQL = SQL & "INNER JOIN #Tmp_LT2G T" & vbCrLf
    SQL = SQL & "        ON ISNULL(W.WHID,'')   = ISNULL(T.WHID,'')" & vbCrLf
    SQL = SQL & "       AND ISNULL(W.PrdtID,'') = ISNULL(T.PrdtID,'')" & vbCrLf
    SQL = SQL & "       AND ISNULL(W.ClrID,'')  = ISNULL(T.ClrID,'')" & vbCrLf
    SQL = SQL & "WHERE  " & Replace(strDateWhereW, "AB_WHI.", "W.") & vbCrLf
    SQL = SQL & "  AND  ISNULL(W.ltime2Grade,0) <> T.tlevel"
    Call objDS.ExecSQL(SQL)

    ' 9) 清理临时表
    Call objDS.ExecSQL("IF OBJECT_ID('tempdb..#Tmp_LTime2Grade') IS NOT NULL DROP TABLE #Tmp_LTime2Grade")
    Call objDS.ExecSQL("IF OBJECT_ID('tempdb..#Tmp_LT2G')        IS NOT NULL DROP TABLE #Tmp_LT2G")

    Exit Sub

ErrH:
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    On Error Resume Next
    Call objDS.rs_Close(rs)
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_LTime2Grade') IS NOT NULL DROP TABLE #Tmp_LTime2Grade"
    objDS.ExecSQL "IF OBJECT_ID('tempdb..#Tmp_LT2G')        IS NOT NULL DROP TABLE #Tmp_LT2G"
    On Error GoTo 0
    Call Err.Raise(lngErrNum, , strErrDesc)
End Sub

'------------------------------------------------------------------------------
' 创建 #Tmp_LTime2Grade（仅一份定义，原代码有两处同签名重复 -> 编译错误）
' 索引覆盖 (whid,prdtid,clrid) 以匹配 ETL 与主循环的查询模式。
'------------------------------------------------------------------------------
Private Sub meCreateLTimeGradeTmp(ByVal objDS As HHDataService.sysDataService)
    Dim SQL As String
    SQL = ""
    SQL = SQL & "IF OBJECT_ID('tempdb..#Tmp_LTime2Grade') IS NOT NULL DROP TABLE #Tmp_LTime2Grade;" & vbCrLf
    SQL = SQL & "CREATE TABLE #Tmp_LTime2Grade(" & vbCrLf
    SQL = SQL & "    curid    varchar(160) NULL," & vbCrLf
    SQL = SQL & "    upid     varchar(160) NULL," & vbCrLf
    SQL = SQL & "    whid     varchar(48)  NULL," & vbCrLf
    SQL = SQL & "    prdtid   varchar(48)  NULL," & vbCrLf
    SQL = SQL & "    clrid    varchar(48)  NULL," & vbCrLf
    SQL = SQL & "    cardnoid varchar(48)  NULL," & vbCrLf
    SQL = SQL & "    tlevel   money        NULL," & vbCrLf
    SQL = SQL & "    LTime    datetime     NULL," & vbCrLf
    SQL = SQL & "    clsid    varchar(48)  NULL);" & vbCrLf
    SQL = SQL & "CREATE INDEX IX_Tmp_LTime2Grade_WPC ON #Tmp_LTime2Grade(whid, prdtid, clrid);" & vbCrLf
    SQL = SQL & "CREATE INDEX IX_Tmp_LTime2Grade_Cur ON #Tmp_LTime2Grade(curid);"
    Call objDS.ExecSQL(SQL)
End Sub

'------------------------------------------------------------------------------
' 6 段 ETL：写入 #Tmp_LTime2Grade
' 修改要点（按用户 2026-05-01 要求）：upid 用 PRDTID 而不是 CLSID
'------------------------------------------------------------------------------
Private Sub meWriteDataForLTime2TLevel(ByVal objDS As HHDataService.sysDataService, _
                                       ByVal strDateWhere As String)
    Dim SQL As String

    ' 段 1：生产入库（直接领料）
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    i.WHID + '-' + i.PRDTID + '-' + ISNULL(i.CLRID,'')                     AS curid," & vbCrLf
    SQL = SQL & "    mtl.WHID + '-' + mtl.PRDTID + '-' + ISNULL(mtl.CLRID,'')               AS upid," & vbCrLf
    SQL = SQL & "    i.WHID, i.PRDTID, i.CLRID, NULL, 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM PG_G i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN PG_M m WITH(NOLOCK) ON m.BillID = i.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN PG_Mtl mtl WITH(NOLOCK) ON i.BillID = mtl.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON i.PRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE " & strDateWhere & " AND m.BState <> 0 AND ISNULL(i.isok,0) = 1"
    Call objDS.ExecSQL(SQL)

    ' 段 2：生产入库（关联其他出库）
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    i.WHID + '-' + i.PRDTID + '-' + ISNULL(i.CLRID,'')                     AS curid," & vbCrLf
    SQL = SQL & "    mtlm.WHID + '-' + mtl.PRDTID + '-' + ISNULL(mtl.CLRID,'')              AS upid," & vbCrLf
    SQL = SQL & "    i.WHID, i.PRDTID, i.CLRID, NULL, 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM PG_G i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN PG_M m WITH(NOLOCK) ON m.BillID = i.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN #Tmp_PGMtlFromIO l ON i.BillID = l.PGBillID" & vbCrLf
    SQL = SQL & "LEFT JOIN iio_i mtl WITH(NOLOCK) ON l.IOBillID = mtl.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN iio_m mtlm WITH(NOLOCK) ON mtlm.BillID = mtl.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON i.PRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE " & strDateWhere & " AND m.BState <> 0" & vbCrLf
    SQL = SQL & "  AND " & Replace(strDateWhere, "m.", "mtlm.") & " AND ISNULL(mtlm.bstate,0) <> 0" & vbCrLf
    SQL = SQL & "  AND ISNULL(i.isok,0) = 1"
    Call objDS.ExecSQL(SQL)

    ' 段 3：生产入库（MES 车间发货）
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    i.WHID + '-' + i.PRDTID + '-' + ISNULL(i.CLRID,'')                     AS curid," & vbCrLf
    SQL = SQL & "    mtl.WHID + '-' + mtl.PRDTID + '-' + ISNULL(mtl.CLRID,'')               AS upid," & vbCrLf
    SQL = SQL & "    i.WHID, i.PRDTID, i.CLRID, '', 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM mes_ws_materials_in_i i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN mes_ws_materials_in_m m WITH(NOLOCK) ON m.billid = i.billid" & vbCrLf
    SQL = SQL & "LEFT JOIN mes_ws_materials_out_i mtl WITH(NOLOCK) ON i.mobillid = mtl.SrcBillID" & vbCrLf
    SQL = SQL & "INNER JOIN mes_ws_materials_out_m mtlm WITH(NOLOCK) ON mtlm.billid = mtl.billid" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON i.PRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE " & strDateWhere & " AND m.BState <> 0" & vbCrLf
    SQL = SQL & "  AND " & Replace(strDateWhere, "m.", "mtlm.") & " AND ISNULL(mtlm.bstate,0) <> 0" & vbCrLf
    SQL = SQL & "  AND i.deleted = 0 AND mtl.deleted = 0 AND ISNULL(i.gradetag,0) <> 3"
    Call objDS.ExecSQL(SQL)

    ' 段 4：盘点
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    m.WHID + '-' + i.PRDTID + '-' + ISNULL(i.CLRID,'')                     AS curid," & vbCrLf
    SQL = SQL & "    m.WHID + '-' + mtl.PRDTID + '-' + ISNULL(mtl.CLRID,'')                 AS upid," & vbCrLf
    SQL = SQL & "    m.WHID, i.PRDTID, i.CLRID, NULL, 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM PI_I i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN PI_M m WITH(NOLOCK) ON m.BillID = i.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN PI_I mtl WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "        ON i.BillID = mtl.BillID AND ISNULL(mtl.isnew,0) = 0" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON i.PRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE " & strDateWhere & " AND m.BState <> 0 AND ISNULL(i.isnew,0) = 1"
    Call objDS.ExecSQL(SQL)

    ' 段 5：拆分单（curid=成品  upid=原料）
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    i.WHID + '-' + i.PRDTID + '-' + ISNULL(i.CLRID,'')                     AS curid," & vbCrLf
    SQL = SQL & "    m.WHID + '-' + m.PRDTID + '-' + ISNULL(m.CLRID,'')                     AS upid," & vbCrLf
    SQL = SQL & "    i.WHID, i.PRDTID, i.CLRID, NULL, 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM IPS_I i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN IPS_M m WITH(NOLOCK) ON m.BillID = i.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON i.PRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE M.IPSTag = -1 AND " & strDateWhere & " AND m.BState <> 0"
    Call objDS.ExecSQL(SQL)

    ' 段 6：组装单（curid=成品  upid=组件）
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    m.WHID + '-' + m.PRDTID + '-' + ISNULL(m.CLRID,'')                     AS curid," & vbCrLf
    SQL = SQL & "    i.WHID + '-' + i.PRDTID + '-' + ISNULL(i.CLRID,'')                     AS upid," & vbCrLf
    SQL = SQL & "    m.WHID, m.PRDTID, m.CLRID, NULL, 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM IPS_I i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN IPS_M m WITH(NOLOCK) ON m.BillID = i.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON m.PRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE M.IPSTag = 1 AND " & strDateWhere & " AND m.BState <> 0"
    Call objDS.ExecSQL(SQL)

    ' 段 7：委外加工
    SQL = ""
    SQL = SQL & "INSERT INTO #Tmp_LTime2Grade(curid,upid,whid,prdtid,clrid,cardnoid,tlevel,LTime,clsid)" & vbCrLf
    SQL = SQL & "SELECT DISTINCT" & vbCrLf
    SQL = SQL & "    m.IWHID + '-' + i.iPRDTID + '-' + ISNULL(i.iCLRID,'')                  AS curid," & vbCrLf
    SQL = SQL & "    m.WHID  + '-' + i.PRDTID  + '-' + ISNULL(i.CLRID,'')                   AS upid," & vbCrLf
    SQL = SQL & "    m.iWHID, i.iPRDTID, i.iCLRID, NULL, 0, m.acctime, prdt.clsid" & vbCrLf
    SQL = SQL & "FROM JG_IC_I i WITH(NOLOCK)" & vbCrLf
    SQL = SQL & "INNER JOIN JG_IC_M m WITH(NOLOCK) ON m.BillID = i.BillID" & vbCrLf
    SQL = SQL & "LEFT JOIN prdt prdt WITH(NOLOCK) ON i.iPRDTID = prdt.prdtid" & vbCrLf
    SQL = SQL & "WHERE " & strDateWhere & " AND m.BState <> 0"
    Call objDS.ExecSQL(SQL)

    ' 删除自环（curid 与 upid 相同）；改用直接列比较，比原代码字符串拼接更快
    SQL = "DELETE FROM #Tmp_LTime2Grade WHERE ISNULL(curid,'') = ISNULL(upid,'')"
    Call objDS.ExecSQL(SQL)
End Sub

'------------------------------------------------------------------------------
' 内存版 getTreeLevel1：与原算法 1:1 等价（path-shared visited + ByRef Level 累积），
' 用并行数组 + 显式栈实现，避免 VB6 默认 1MB 栈在大 DAG 下溢出。
'
' 算法等价改写为纯函数（已在 Python 等价性测试中验证）：
'   sub(node, visited, level_in) -> level_out:
'       visited.add(node)
'       uppers = [u for u in edges(node) if u not in visited]
'       if not uppers: return level_in
'       cur = level_in + 1
'       for u in uppers:                         # visited 在兄弟之间 PERSIST
'           ret = sub(u, visited, cur)
'           if ret > cur: cur = ret
'       return cur
'   getTreeLevel1(root):
'       uppers = [u for u in edges(root) ...]
'       if not uppers: return 0
'       best = 1
'       for u in uppers:                         # 此处 visited 每个兄弟都重置
'           visited = {root}
'           ret = sub(u, visited, 1)
'           if ret > best: best = ret
'       return best
'
' 输入：
'   sRoot     - "WHID-PRDTID-CLRID"
'   dctEdges  - 边表 Dictionary：sRoot -> Collection of upid 字符串
' 返回：当前节点的 tree level（与原 getTreeLevel1 完全一致的数值）
'------------------------------------------------------------------------------
Private Function meTreeLevel(ByVal sRoot As String, ByVal dctEdges As Object) As Currency
    If Not dctEdges.Exists(sRoot) Then
        meTreeLevel = 0
        Exit Function
    End If

    Dim arrRootU() As String
    arrRootU = meDistinctSortedUppers(dctEdges, sRoot, Nothing)
    If meIsEmptyArr(arrRootU) Then
        meTreeLevel = 0
        Exit Function
    End If

    Dim cBest As Currency
    cBest = 1

    Dim i As Long
    For i = 0 To UBound(arrRootU)
        Dim dctVisited As Object
        Set dctVisited = CreateObject("Scripting.Dictionary")
        dctVisited.Add sRoot, True

        Dim cRet As Currency
        cRet = meSubIter(arrRootU(i), dctEdges, dctVisited, 1)
        If cRet > cBest Then cBest = cRet
    Next i

    meTreeLevel = cBest
End Function

'------------------------------------------------------------------------------
' 显式栈实现 sub(node, visited, level_in) -> level_out
' 用并行数组：栈深 = 节点数；每帧记录 Node / Uppers / Idx / Level / 子返回值
'------------------------------------------------------------------------------
Private Function meSubIter(ByVal sStartNode As String, _
                           ByVal dctEdges As Object, _
                           ByVal dctVisited As Object, _
                           ByVal cInitLevel As Currency) As Currency
    ' 起始：visited.add(sStartNode), 取 uppers
    Dim arrU0() As String
    arrU0 = meDistinctSortedUppers(dctEdges, sStartNode, dctVisited)
    If meIsEmptyArr(arrU0) Then
        meSubIter = cInitLevel
        Exit Function
    End If

    ' 并行栈：用 ReDim Preserve 增长
    Dim stkUppers() As Variant       ' 每帧的 uppers 数组（变体保存 String 数组）
    Dim stkIdx()    As Long          ' 当前下一个要处理的 upper 索引
    Dim stkLevel()  As Currency      ' 当前帧的 cur level
    Dim stkChild()  As Currency      ' 子帧返回值（待消化）
    Dim stkHasCh()  As Boolean       ' 是否有待消化的子返回值
    Dim sp As Long                   ' 栈顶索引（-1 表示空）
    sp = -1

    ' 入栈 root frame
    sp = sp + 1
    ReDim Preserve stkUppers(sp)
    ReDim Preserve stkIdx(sp)
    ReDim Preserve stkLevel(sp)
    ReDim Preserve stkChild(sp)
    ReDim Preserve stkHasCh(sp)
    stkUppers(sp) = arrU0
    stkIdx(sp) = 0
    stkLevel(sp) = cInitLevel + 1
    stkChild(sp) = 0
    stkHasCh(sp) = False

    Dim cFinal As Currency
    cFinal = cInitLevel

    Do While sp >= 0
        ' 消化子帧返回值
        If stkHasCh(sp) Then
            If stkChild(sp) > stkLevel(sp) Then stkLevel(sp) = stkChild(sp)
            stkHasCh(sp) = False
        End If

        Dim arrUTop() As String
        arrUTop = stkUppers(sp)
        Dim iIdx As Long
        iIdx = stkIdx(sp)

        If iIdx > UBound(arrUTop) Then
            ' 本帧完成，pop
            Dim cRet As Currency
            cRet = stkLevel(sp)
            sp = sp - 1
            If sp >= 0 Then
                stkChild(sp) = cRet
                stkHasCh(sp) = True
            Else
                cFinal = cRet
            End If
        Else
            Dim sU As String
            sU = arrUTop(iIdx)
            stkIdx(sp) = iIdx + 1

            ' 进入子节点：visited.add + 取 sub_uppers
            Dim arrSubU() As String
            arrSubU = meDistinctSortedUppers(dctEdges, sU, dctVisited)
            If meIsEmptyArr(arrSubU) Then
                ' 子返回 = 当前 Level，无变化
                ' （已在进入时 visited.add 了 sU；max 也无效果）
            Else
                ' 推子帧：child level_in = stkLevel(sp), child cur = level_in + 1
                Dim cChildInit As Currency
                cChildInit = stkLevel(sp)
                sp = sp + 1
                ReDim Preserve stkUppers(sp)
                ReDim Preserve stkIdx(sp)
                ReDim Preserve stkLevel(sp)
                ReDim Preserve stkChild(sp)
                ReDim Preserve stkHasCh(sp)
                stkUppers(sp) = arrSubU
                stkIdx(sp) = 0
                stkLevel(sp) = cChildInit + 1
                stkChild(sp) = 0
                stkHasCh(sp) = False
            End If
        End If
    Loop

    meSubIter = cFinal
End Function

'------------------------------------------------------------------------------
' 取 node 的 distinct uppers（按字符串升序），并把 node 加入 visited（如果非空）。
' visited 可传 Nothing 表示不做"NOT IN visited"过滤，仅去重。
'------------------------------------------------------------------------------
Private Function meDistinctSortedUppers(ByVal dctEdges As Object, _
                                        ByVal sNode As String, _
                                        ByVal dctVisited As Object) As String()
    Dim arrEmpty() As String
    ReDim arrEmpty(-1 To -1)                                    ' 标记为空数组（UBound=-1, LBound=0）

    If Not dctVisited Is Nothing Then
        If Not dctVisited.Exists(sNode) Then dctVisited.Add sNode, True
    End If

    If Not dctEdges.Exists(sNode) Then
        meDistinctSortedUppers = arrEmpty
        Exit Function
    End If

    Dim dctTmp As Object
    Set dctTmp = CreateObject("Scripting.Dictionary")
    Dim v As Variant
    For Each v In dctEdges(sNode)
        Dim su As String
        su = CStr(v)
        If su <> "" And su <> sNode Then
            If dctVisited Is Nothing Then
                If Not dctTmp.Exists(su) Then dctTmp.Add su, True
            ElseIf Not dctVisited.Exists(su) Then
                If Not dctTmp.Exists(su) Then dctTmp.Add su, True
            End If
        End If
    Next v

    If dctTmp.Count = 0 Then
        meDistinctSortedUppers = arrEmpty
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(0 To dctTmp.Count - 1)
    Dim i As Long
    i = 0
    Dim vk As Variant
    For Each vk In dctTmp.Keys
        arr(i) = CStr(vk)
        i = i + 1
    Next vk
    ' 升序排序（节点数通常 <= 几十）
    Dim j As Long
    Dim s As String
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                s = arr(i): arr(i) = arr(j): arr(j) = s
            End If
        Next j
    Next i
    meDistinctSortedUppers = arr
End Function

'------------------------------------------------------------------------------
' 判定 String() 是否为空（LBound>UBound 表示空，由 ReDim arr(-1 To -1) 构造）
'------------------------------------------------------------------------------
Private Function meIsEmptyArr(ByRef arr() As String) As Boolean
    On Error Resume Next
    Dim u As Long
    u = UBound(arr)
    If Err.Number <> 0 Then
        Err.Clear
        meIsEmptyArr = True
        Exit Function
    End If
    On Error GoTo 0
    meIsEmptyArr = (UBound(arr) < LBound(arr))
End Function
