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
'   - 日期：Format(yyyy-MM-dd hh:mm:ss)
'   - 数值：CStr 直接转
'   - Boolean：1/0
'   - Decimal/Currency: 通过 CStr 保留精度（VB6 Currency 内部 Int64 缩放 4 位）
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
        ' 跳过只读/计算列（与 ADO Update 一致）
        If (fld.Attributes And adFldUpdatable) <> 0 Or (fld.Attributes And adFldUnknownUpdatable) <> 0 Then
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
        If (fld.Attributes And adFldUpdatable) <> 0 Or (fld.Attributes And adFldUnknownUpdatable) <> 0 Then
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
        If (fld.Attributes And adFldUpdatable) <> 0 Or (fld.Attributes And adFldUnknownUpdatable) <> 0 Then
            If sVals <> "" Then sVals = sVals & ","
            sVals = sVals & meQuoteValue(fld)
        End If
    Next fld
    meBuildValuesRow = sVals
End Function


'==============================================================================
' Private：把 ADO Field 值转为 SQL VALUES 中的文字常量
'   等价于 ADO 内部 OLE DB 类型转换，但用文本表示
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
            ' SQL Server 接受 'yyyy-MM-dd hh:mm:ss' ISO 格式
            meQuoteValue = "'" & Format(fld.Value, "yyyy-MM-dd hh:mm:ss") & "'"

        Case adBoolean
            meQuoteValue = IIf(CBool(fld.Value), "1", "0")

        Case adGUID
            meQuoteValue = "'" & CStr(fld.Value) & "'"

        Case adNumeric, adDecimal, adCurrency, adDouble, adSingle, _
             adInteger, adSmallInt, adBigInt, adTinyInt, _
             adUnsignedInt, adUnsignedSmallInt, adUnsignedBigInt, adUnsignedTinyInt
            ' Currency: VB6 内部 Int64 / 10000，CStr 输出 "1234.5678"
            ' Decimal/Numeric: 同样用 CStr 保留精度
            meQuoteValue = CStr(fld.Value)

        Case adBinary, adVarBinary, adLongVarBinary
            ' 二进制不应出现在 FVou_M / FVou_I（无此类字段）；防御性处理
            meQuoteValue = "0x" & meBytesToHex(fld.Value)

        Case Else
            ' 兜底当字符串
            meQuoteValue = "N'" & Replace(CStr(fld.Value), "'", "''") & "'"
    End Select
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
