Attribute VB_Name = "VouBatchWriter"
'==============================================================================
' Module      : VouBatchWriter.bas
' Description : ADO Recordset → 批量 INSERT VALUES 写入器（PR-4 第四波 C2）
'
'   用 INSERT VALUES 替代 sysDS.SaveData 的 ADO Recordset.Update 行级写入。
'   仅在 VouBatchOps.InBatchMode=True 时启用；False 时调用方走原 SaveData。
'
' 使用方法    :
'   Call VouBatchWriter.InsertSingleRow(objDS, MainData, "FVou_M_T")
'   Call VouBatchWriter.InsertMultiRows(objDS, ItemsData, "FVou_I_T", 200)
'
' 100% 等价   :
'   - 字段顺序：按 rs.Fields 顺序，与 ADO Update 一致
'   - NULL：IsNull(value) 输出 "NULL" 关键字
'   - 字符串：' 转义为 ''
'   - 日期：Format(yyyy-MM-dd HH:mm:ss)  ⚠️ 用大写 HH（24 小时制）
'   - 数值：用 meNumToStr（始终用 . 当小数点，不受区域影响）
'   - Boolean：1/0
'   - Decimal/Currency: 通过 meNumToStr 保留精度（FVou 表用 money/decimal(28,8)，4 位小数足够）
'
' Database    : 兼容 Microsoft SQL Server 2008
'==============================================================================
Option Explicit

' 单条 SQL 的最大字符数（SQL Server 默认 batch 大小 65536 个 token，保守用 65000 字节）
Private Const MAX_SQL_BYTES As Long = 65000


'==============================================================================
' Public：把 rs 当前行作为单行 INSERT 写入 tabName
'   等价于 ADO Recordset.AddNew + Update，但走 INSERT VALUES
'==============================================================================
Public Sub InsertSingleRow(ByVal objDS As HHDataService.sysDataService, _
                           ByVal rs As ADODB.Recordset, _
                           ByVal tabName As String)
    Dim sCols As String, sVals As String
    Dim fld As ADODB.Field

    For Each fld In rs.Fields
        If meShouldInclude(fld) Then
            If sCols <> "" Then
                sCols = sCols & ","
                sVals = sVals & ","
            End If
            sCols = sCols & "[" & fld.Name & "]"
            sVals = sVals & meQuoteValue(fld)
        End If
    Next fld

    Call objDS.ExecSQL("INSERT INTO [" & tabName & "](" & sCols & ") VALUES (" & sVals & ")")
End Sub


'==============================================================================
' Private：判断字段是否应包含在 INSERT 中
'   排除 IDENTITY 自增列（adFldRowID）；其它所有字段（含 Updatable=0 但非 IDENTITY）
'   都包含。这与 ADO Recordset Update 行为一致。
'   FVou_M / FVou_I 表用 BillID/ITMID 作为业务主键，无 IDENTITY 列，所以
'   实际上所有字段都会被 INSERT。
'==============================================================================
Private Function meShouldInclude(ByVal fld As ADODB.Field) As Boolean
    ' adFldRowID = IDENTITY/RowVersion 等系统自动生成列
    If (fld.Attributes And adFldRowID) <> 0 Then
        meShouldInclude = False
        Exit Function
    End If
    ' adFldRowVersion = timestamp 类型，由 SQL Server 自动维护
    If (fld.Attributes And adFldRowVersion) <> 0 Then
        meShouldInclude = False
        Exit Function
    End If
    meShouldInclude = True
End Function


'==============================================================================
' Public：把 rs 全量行作为批量 INSERT VALUES (...),(...),... 写入 tabName
'   分页：每 pageSize 行一批，或 SQL 字符串达到 MAX_SQL_BYTES 提前 flush
'==============================================================================
Public Sub InsertMultiRows(ByVal objDS As HHDataService.sysDataService, _
                          ByVal rs As ADODB.Recordset, _
                          ByVal tabName As String, _
                          Optional ByVal pageSize As Long = 200)
    Dim sHeader As String
    Dim sBuf As String
    Dim cnt As Long
    Dim sCols As String

    sHeader = ""
    sBuf = ""
    cnt = 0

    Call objDS.rs_MoveFirst(rs)
    Do While Not rs.EOF
        If sHeader = "" Then
            sCols = meBuildColList(rs)
            sHeader = "INSERT INTO [" & tabName & "](" & sCols & ") VALUES "
            sBuf = ""
            cnt = 0
        End If

        Dim sRow As String
        sRow = "(" & meBuildValuesRow(rs) & ")"

        ' 如果加上这行会超过 SQL 长度上限，先 flush
        If Len(sHeader) + Len(sBuf) + Len(sRow) + 2 > MAX_SQL_BYTES And cnt > 0 Then
            Call objDS.ExecSQL(sHeader & sBuf)
            sBuf = ""
            cnt = 0
        End If

        If sBuf <> "" Then sBuf = sBuf & ","
        sBuf = sBuf & sRow
        cnt = cnt + 1

        If cnt >= pageSize Then
            Call objDS.ExecSQL(sHeader & sBuf)
            sBuf = ""
            cnt = 0
        End If

        rs.MoveNext
    Loop

    If sBuf <> "" Then
        Call objDS.ExecSQL(sHeader & sBuf)
    End If
End Sub


'==============================================================================
' Private：构造 INSERT INTO 的列名清单（按 rs.Fields 顺序，含 [] 包裹）
'==============================================================================
Private Function meBuildColList(ByVal rs As ADODB.Recordset) As String
    Dim sCols As String
    Dim fld As ADODB.Field
    For Each fld In rs.Fields
        If meShouldInclude(fld) Then
            If sCols <> "" Then sCols = sCols & ","
            sCols = sCols & "[" & fld.Name & "]"
        End If
    Next fld
    meBuildColList = sCols
End Function


'==============================================================================
' Private：构造单行 VALUES 内容（按 rs.Fields 顺序）
'==============================================================================
Private Function meBuildValuesRow(ByVal rs As ADODB.Recordset) As String
    Dim sVals As String
    Dim fld As ADODB.Field
    For Each fld In rs.Fields
        If meShouldInclude(fld) Then
            If sVals <> "" Then sVals = sVals & ","
            sVals = sVals & meQuoteValue(fld)
        End If
    Next fld
    meBuildValuesRow = sVals
End Function


'==============================================================================
' Private：把 ADO Field 值转为 SQL VALUES 中的文字常量
'   等价于 ADO 内部 OLE DB 类型转换，但用文本表示
'
'   关键修复（审计 PR-4 时发现）：
'     1) 日期用 "HH:mm:ss"（24 小时）而非 "hh:mm:ss"（12 小时）
'        VB6 Format() 的 hh 是 12 小时制，下午 13:00 会被输出为 "01:00:00"
'     2) 数值转字符串用 meNumToStr，确保始终用 "." 当小数点
'        VB6 CStr() 在德语/俄语等区域用 "," 当小数点，会破坏 SQL 语法
'==============================================================================
Private Function meQuoteValue(ByVal fld As ADODB.Field) As String
    If IsNull(fld.Value) Then
        meQuoteValue = "NULL"
        Exit Function
    End If

    Select Case fld.Type
        Case adVarChar, adVarWChar, adChar, adWChar, _
             adLongVarChar, adLongVarWChar
            meQuoteValue = "N'" & Replace(CStr(fld.Value), "'", "''") & "'"

        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            ' ✅ 用大写 HH 强制 24 小时制
            ' VB6 Format() 中 hh = 12 小时制，HH = 24 小时制
            meQuoteValue = "'" & Format$(fld.Value, "yyyy-MM-dd HH:mm:ss") & "'"

        Case adBoolean
            meQuoteValue = IIf(CBool(fld.Value), "1", "0")

        Case adGUID
            ' GUID 字符串可能含 {}，SQL Server 接受不带 {} 的形式
            Dim sGUID As String
            sGUID = CStr(fld.Value)
            If Left$(sGUID, 1) = "{" Then sGUID = Mid$(sGUID, 2, Len(sGUID) - 2)
            meQuoteValue = "'" & sGUID & "'"

        Case adNumeric, adDecimal, adCurrency, adDouble, adSingle, _
             adInteger, adSmallInt, adBigInt, adTinyInt, _
             adUnsignedInt, adUnsignedSmallInt, adUnsignedBigInt, adUnsignedTinyInt
            ' ✅ 用 meNumToStr，区域无关
            meQuoteValue = meNumToStr(fld.Value)

        Case adBinary, adVarBinary, adLongVarBinary
            ' 二进制不应出现在 FVou_M / FVou_I（无此类字段）；防御性处理
            meQuoteValue = "0x" & meBytesToHex(fld.Value)

        Case Else
            ' 兜底当字符串
            meQuoteValue = "N'" & Replace(CStr(fld.Value), "'", "''") & "'"
    End Select
End Function


'==============================================================================
' Private：数值 → 区域无关的字符串（始终用 "." 当小数点）
'   VB6 CStr() 受 LCID 影响，在 zh-CN / de-DE 等区域可能输出 "1,5"。
'   Str() 函数始终用 "." 但正数前会留空格，需要 Trim。
'==============================================================================
Private Function meNumToStr(ByVal v As Variant) As String
    ' Str(v) 对所有数值类型（含 Currency / Double / Decimal）返回区域无关的
    ' "." 小数点字符串。正数前面有 1 空格，需要 Trim。
    meNumToStr = Trim$(Str$(v))
End Function


'==============================================================================
' Private：bytes → hex string（用于 adBinary 兜底）
'==============================================================================
Private Function meBytesToHex(ByVal v As Variant) As String
    Dim b()  As Byte
    Dim s    As String
    Dim i    As Long
    b = v
    For i = LBound(b) To UBound(b)
        s = s & Right("00" & Hex(b(i)), 2)
    Next i
    meBytesToHex = s
End Function
