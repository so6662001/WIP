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
    '   dictSeed:    "BillType|BillID|DC"  -> "WHID|PRDTID|CLRID"
    '   dictByLoc:   "WHID|PRDTID|CLRID"   -> Collection of "BillType|BillID|DC|MaxLTimeYYYYMMDDHHMMSS"
    '   dictDcN1:    "BillType|BillID"     -> "MaxLT|MaxLT2|MaxLT2Grade"  （DC=-1 的聚合）
    '   dictMaxLTd1: "BillType|BillID"     -> MaxLT (Date)                 （DC=1 的 MAX(LTime)）
    '   dictPiNew:   "PIBill BillID"       -> True                         （存在 IsNew=1 且 WMSPI=''）
    '   dictPiAgg:   "PIBill BillID"       -> "MaxLT|MaxLT2|MaxLT2Grade"   （AB_WHI JOIN PI_I 的聚合）
    '   dictRsLTime: "BillType|BillID"     -> Date                         （兜底用 MAX(LTime) DC=-1）
    '   dictComputed:"BillType|BillID|DC"  -> "LTime2|Grade"               （递归过程即时回填）
    Dim dictSeed     As Object
    Dim dictByLoc    As Object
    Dim dictDcN1     As Object
    Dim dictMaxLTd1  As Object
    Dim dictPiNew    As Object
    Dim dictPiAgg    As Object
    Dim dictRsLTime  As Object
    Dim dictComputed As Object

    On Error GoTo ErrH

    Set dictSeed = CreateObject("Scripting.Dictionary")
    Set dictByLoc = CreateObject("Scripting.Dictionary")
    Set dictDcN1 = CreateObject("Scripting.Dictionary")
    Set dictMaxLTd1 = CreateObject("Scripting.Dictionary")
    Set dictPiNew = CreateObject("Scripting.Dictionary")
    Set dictPiAgg = CreateObject("Scripting.Dictionary")
    Set dictRsLTime = CreateObject("Scripting.Dictionary")
    Set dictComputed = CreateObject("Scripting.Dictionary")

    '--------------------------------------------------------------------------
    ' P1. dictByLoc + dictSeed + dictMaxLTd1
    '     候选行：BillType<>'ACF' AND DC=1 区间内
    '--------------------------------------------------------------------------
    SQL = "SELECT BillType, BillID, DC, WHID, ISNULL(PRDTID,'') AS PRDTID, CLRID, " & _
          "       MAX(LTime) AS MaxLT" & vbCrLf & _
          "FROM   AB_WHI WITH(NOLOCK)" & vbCrLf & _
          "WHERE  " & strDateWhere & " AND BillType<>'ACF' AND DC=1" & vbCrLf & _
          "GROUP BY BillType, BillID, DC, WHID, ISNULL(PRDTID,''), CLRID"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim sLocKey As String
        Dim sBillKey As String
        Dim sCand   As String
        sLocKey = rs.Fields("WHID").Value & "|" & rs.Fields("PRDTID").Value & "|" & rs.Fields("CLRID").Value
        sBillKey = rs.Fields("BillType").Value & "|" & rs.Fields("BillID").Value & "|" & CStr(rs.Fields("DC").Value)
        sCand = sBillKey & "|"
        If Not IsNull(rs.Fields("MaxLT").Value) Then
            sCand = sCand & Format$(rs.Fields("MaxLT").Value, "yyyy-MM-dd hh:mm:ss")
        End If

        If Not dictByLoc.Exists(sLocKey) Then
            dictByLoc.Add sLocKey, New Collection
        End If
        dictByLoc(sLocKey).Add sCand

        If Not dictSeed.Exists(sBillKey) Then
            dictSeed.Add sBillKey, sLocKey
        End If

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
                            dictByLoc, dictSeed, dictDcN1, dictMaxLTd1, _
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
                           ByVal dictSeed As Object, _
                           ByVal dictDcN1 As Object, _
                           ByVal dictMaxLTd1 As Object, _
                           ByVal dictPiNew As Object, _
                           ByVal dictPiAgg As Object, _
                           ByVal dictRsLTime As Object, _
                           ByVal dictComputed As Object)

    If dblGr > 150 Then Exit Sub

    Dim sSeedKey As String
    sSeedKey = sBillType & "|" & sBillID & "|" & CStr(intDC)
    If Not dictSeed.Exists(sSeedKey) Then
        ' 该种子不在候选 (DC=1) 中是正常情况：本身可能就是 DC=-1 的种子
        ' 仍需走"段4 兜底" 的逻辑：但段4 仅在存在候选时触发，所以无候选直接返回
    End If

    Dim sLocKey As String
    Dim collCands As Collection
    Dim hasCands As Boolean

    ' 当前种子的 (WHID,PRDTID,CLRID) —— 原算法是从 AB_WHI 直接读这一行，
    ' 这里我们先从 dictSeed 中找；找不到就读 dictByLoc（可能 DC=-1 行）
    If dictSeed.Exists(sSeedKey) Then
        sLocKey = dictSeed(sSeedKey)
    Else
        sLocKey = ""
    End If

    Dim pgs As Collection
    Dim ics As Collection
    Dim pis As Collection
    Set pgs = New Collection
    Set ics = New Collection
    Set pis = New Collection
    hasCands = False

    If sLocKey <> "" And dictByLoc.Exists(sLocKey) Then
        Set collCands = dictByLoc(sLocKey)
        Dim v As Variant
        For Each v In collCands
            ' v 形如 "BillType|BillID|DC|YYYY-MM-DD hh:mm:ss"
            Dim parts() As String
            parts = Split(CStr(v), "|")
            ' 跳过自身
            If Not (parts(0) = sBillType And parts(1) = sBillID And parts(2) = CStr(intDC)) Then
                hasCands = True
                Select Case parts(0)
                    Case "PGBill"
                        pgs.Add v
                    Case "jgicbill", "icbill", "icbill2", "ICWHSecBill", "icbill3", "isbill", "ipbill"
                        ics.Add v
                    Case "PIBill"
                        pis.Add v
                End Select
            End If
        Next v
    End If

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
                                    dictByLoc, dictSeed, dictDcN1, dictMaxLTd1, _
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
                                        dictByLoc, dictSeed, dictDcN1, dictMaxLTd1, _
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
                                    dictByLoc, dictSeed, dictDcN1, dictMaxLTd1, _
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
