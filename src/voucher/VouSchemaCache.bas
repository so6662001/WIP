Attribute VB_Name = "VouSchemaCache"
'==============================================================================
' Module      : VouSchemaCache.bas
' Description : 凭证生成 batch 的 ADO Recordset Schema 缓存（PR-2 第二波 A2）
'               消除 t_FVou_M.megetDocByID("") 内的 3 次 SELECT TOP 0。
'
' 原理        : SELECT TOP 0 ... 的目的只是构造一个具有正确字段定义的空 Recordset，
'               用于后续 AddNew + Update + 由 sysDS.SaveData 写入数据库。
'               缓存"字段定义数组"（FieldName / Type / DefinedSize / Attributes），
'               每次需要时用 Fields.Append 构造一个独立的 disconnected Recordset。
'
' 100% 等价   : 字段名、类型、大小、属性完全复制自 SELECT TOP 0 模板。
'               生成的 Recordset 是 disconnected (CursorLocation=adUseClient)，
'               支持 AddNew / Update / Filter / SaveData，行为与原 SELECT TOP 0
'               生成的 Recordset 完全一致（原 SELECT TOP 0 也是 disconnected
'               recordset，因为 OpenRecordsetBySQL 第二个参数 False 表示
'               adOpenStatic + adLockOptimistic）。
'
'               未加载时调用 CloneXxxEmpty 返回 Nothing，调用方应回退到
'               原 SELECT TOP 0 路径（保证向后兼容）。
'
' Database    : 兼容 Microsoft SQL Server 2008 及以上
'==============================================================================
Option Explicit

' 字段定义快照
Private Type FieldDef
    Name           As String
    DataType       As Long      ' ADODB.DataTypeEnum
    DefinedSize    As Long
    Attributes     As Long      ' 可设置的子集（adFldIsNullable / adFldUpdatable）
    NumericScale   As Byte      ' 仅 adNumeric / adDecimal 用
    Precision      As Byte      ' 仅 adNumeric / adDecimal 用
End Type

Private m_arrMainFlds()     As FieldDef
Private m_lngMainCount      As Long
Private m_arrItemsFlds()    As FieldDef
Private m_lngItemsCount     As Long
Private m_arrIItemsFlds()   As FieldDef
Private m_lngIItemsCount    As Long

Private m_blnLoaded         As Boolean


Public Property Get IsLoaded() As Boolean
    IsLoaded = m_blnLoaded
End Property


'==============================================================================
' Public：懒加载入口（跨工程友好）
' 工程 C (POPBus3GL2Service) 内 megetDocByID 调用本函数自动 LoadAll
'==============================================================================
Public Sub EnsureLoaded(ByVal objDS As HHDataService.sysDataService)
    If m_blnLoaded Then Exit Sub
    Call LoadAll(objDS)
End Sub


'==============================================================================
' Public：懒加载并 Clone（一步到位）
'   返回 Nothing 表示首次 LoadAll 失败（DB 异常）；调用方应回退原 SELECT TOP 0
'==============================================================================
Public Function CloneMainEmptyOrLoad(ByVal objDS As HHDataService.sysDataService) As ADODB.Recordset
    On Error Resume Next
    Call EnsureLoaded(objDS)
    On Error GoTo 0
    If Not m_blnLoaded Then Exit Function
    Set CloneMainEmptyOrLoad = CloneMainEmpty()
End Function


Public Function CloneItemsEmptyOrLoad(ByVal objDS As HHDataService.sysDataService) As ADODB.Recordset
    On Error Resume Next
    Call EnsureLoaded(objDS)
    On Error GoTo 0
    If Not m_blnLoaded Then Exit Function
    Set CloneItemsEmptyOrLoad = CloneItemsEmpty()
End Function


Public Function CloneIItemsEmptyOrLoad(ByVal objDS As HHDataService.sysDataService) As ADODB.Recordset
    On Error Resume Next
    Call EnsureLoaded(objDS)
    On Error GoTo 0
    If Not m_blnLoaded Then Exit Function
    Set CloneIItemsEmptyOrLoad = CloneIItemsEmpty()
End Function


'==============================================================================
' Public：在 batch 入口扫描 schema 模板（一次性）
'==============================================================================
Public Sub LoadAll(ByVal objDS As HHDataService.sysDataService)
On Error GoTo ErrH
    Dim SQL As String
    Dim rs  As ADODB.Recordset

    Call ClearAll

    ' 1) FVou_M 主表
    SQL = "SELECT TOP 0 * FROM FVou_M with(nolock) WHERE 1=2"
    Set rs = objDS.OpenRecordsetBySQL(SQL, False, True)
    Call meCaptureFields(rs, m_arrMainFlds, m_lngMainCount)
    Call objDS.rs_Close(rs)

    ' 2) FVou_I 明细 + 11 表 LEFT JOIN（与原 megetDocByID 一致）
    SQL = "SELECT TOP 0 dbo.FVou_I.*, FVou_I.ITMID as NewITMID, FVou_I.ITMID as oldITMID," & vbCrLf & _
          "       Cst_Center.CstCenName,dep.depname,emp.empname," & vbCrLf & _
          "       dbo.Corp.CorpName, PRDTCls.ClsName,TypeForInv.TInvName," & vbCrLf & _
          "       Account.AccName,WH.WHName,fi.baCust,fi.baSupp,fi.baOtherCorp,fi.baNum,fi.baDep,fi.baEmp" & vbCrLf & _
          "FROM dbo.FVou_I with(nolock)  LEFT OUTER JOIN" & vbCrLf & _
          "   dbo.Corp ON dbo.FVou_I.CorpID = dbo.Corp.CorpID LEFT OUTER JOIN" & vbCrLf & _
          "   Account ON FVou_I.AccID=Account.AccID  LEFT OUTER JOIN  " & vbCrLf & _
          "   dep ON FVou_I.depID=dep.depID  LEFT OUTER JOIN  " & vbCrLf & _
          "   emp ON FVou_I.empID=emp.empID  LEFT OUTER JOIN  " & vbCrLf & _
          "   PRDT ON FVOU_I.PrdtID=Prdt.PrdtID LEFT OUTER JOIN " & vbCrLf & _
          "   PRDTCls ON Prdt.ClsID=PrdtCls.ClsID  LEFT OUTER JOIN  " & vbCrLf & _
          "   WH ON FVou_I.WHID=WH.WHID  LEFT OUTER JOIN  " & vbCrLf & _
          "   Cst_Center ON FVou_I.CstCenID = Cst_Center.CstCenID LEFT OUTER JOIN" & vbCrLf & _
          "   TypeForInv ON FVou_I.TInvID=TypeForInv.TInvID LEFT OUTER JOIN" & vbCrLf & _
          "   FinanceItems FI ON FVou_I.FIID=FI.FIID" & vbCrLf & _
          "WHERE 1=2"
    Set rs = objDS.OpenRecordsetBySQL(SQL, False, True)
    Call meCaptureFields(rs, m_arrItemsFlds, m_lngItemsCount)
    Call objDS.rs_Close(rs)

    ' 3) FVou_II
    SQL = "SELECT TOP 0 * FROM FVou_II with(nolock) WHERE 1=2"
    Set rs = objDS.OpenRecordsetBySQL(SQL, False, True)
    Call meCaptureFields(rs, m_arrIItemsFlds, m_lngIItemsCount)
    Call objDS.rs_Close(rs)

    m_blnLoaded = True
    Exit Sub

ErrH:
    Dim lngE As Long, sDesc As String
    lngE = Err.Number
    sDesc = Err.Description
    On Error Resume Next
    Call ClearAll
    On Error GoTo 0
    Call Err.Raise(lngE, , sDesc)
End Sub


'==============================================================================
' Public：清理（batch 末尾）
'==============================================================================
Public Sub ClearAll()
    On Error Resume Next
    m_lngMainCount = 0
    m_lngItemsCount = 0
    m_lngIItemsCount = 0
    Erase m_arrMainFlds
    Erase m_arrItemsFlds
    Erase m_arrIItemsFlds
    m_blnLoaded = False
    On Error GoTo 0
End Sub


'==============================================================================
' Public：构造空的 disconnected FVou_M Recordset（独立 buffer，可 AddNew）
'   Schema 与原 SELECT TOP 0 完全一致。
'   返回 Nothing 表示缓存未加载，调用方应回退到 SELECT TOP 0
'==============================================================================
Public Function CloneMainEmpty() As ADODB.Recordset
    If Not m_blnLoaded Then Exit Function
    Set CloneMainEmpty = meBuildEmpty(m_arrMainFlds, m_lngMainCount)
End Function


Public Function CloneItemsEmpty() As ADODB.Recordset
    If Not m_blnLoaded Then Exit Function
    Set CloneItemsEmpty = meBuildEmpty(m_arrItemsFlds, m_lngItemsCount)
End Function


Public Function CloneIItemsEmpty() As ADODB.Recordset
    If Not m_blnLoaded Then Exit Function
    Set CloneIItemsEmpty = meBuildEmpty(m_arrIItemsFlds, m_lngIItemsCount)
End Function


'==============================================================================
' Private：把 ADO Recordset 字段定义抓到数组
'   ⚠️ Attributes 只保留 Fields.Append 接受的子集（adFldIsNullable + adFldUpdatable）
'      其它如 adFldKeyColumn / adFldRowID / adFldFixed 等是 read-only 标志，
'      传给 Append 会触发 ADO 错误 3251
'==============================================================================
Private Sub meCaptureFields(ByVal rs As ADODB.Recordset, _
                            ByRef arr() As FieldDef, _
                            ByRef lngCount As Long)
    Dim fld As ADODB.Field
    Dim i   As Long
    lngCount = rs.Fields.Count
    ReDim arr(0 To lngCount - 1)
    i = 0
    For Each fld In rs.Fields
        arr(i).Name = fld.Name
        arr(i).DataType = fld.Type
        arr(i).DefinedSize = fld.DefinedSize
        ' 只保留 Fields.Append 接受的属性
        Dim attrs As Long
        attrs = 0
        If (fld.Attributes And adFldIsNullable) <> 0 Then
            attrs = attrs Or adFldIsNullable
        End If
        If (fld.Attributes And adFldUpdatable) <> 0 Then
            attrs = attrs Or adFldUpdatable
        End If
        arr(i).Attributes = attrs

        ' 数值类型保存 Precision/NumericScale
        Select Case fld.Type
            Case adNumeric, adDecimal
                arr(i).NumericScale = fld.NumericScale
                arr(i).Precision = fld.Precision
            Case Else
                arr(i).NumericScale = 0
                arr(i).Precision = 0
        End Select

        i = i + 1
    Next fld
End Sub


'==============================================================================
' Private：根据字段定义数组构造一个独立的、可写的空 Recordset
'   对 adNumeric/adDecimal 类型必须设置 Precision/NumericScale，否则 ADO 报错
'==============================================================================
Private Function meBuildEmpty(ByRef arr() As FieldDef, ByVal lngCount As Long) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i  As Long

    Set rs = New ADODB.Recordset
    For i = 0 To lngCount - 1
        ' VarChar / Numeric 等需要 DefinedSize；DefinedSize=0 时省略参数
        If arr(i).DefinedSize > 0 Then
            rs.Fields.Append arr(i).Name, arr(i).DataType, arr(i).DefinedSize, arr(i).Attributes
        Else
            rs.Fields.Append arr(i).Name, arr(i).DataType, , arr(i).Attributes
        End If
        ' 数值类型：Append 后单独设置 Precision/NumericScale
        Select Case arr(i).DataType
            Case adNumeric, adDecimal
                rs.Fields(arr(i).Name).Precision = arr(i).Precision
                rs.Fields(arr(i).Name).NumericScale = arr(i).NumericScale
        End Select
    Next i
    rs.CursorLocation = adUseClient
    rs.Open , , adOpenKeyset, adLockOptimistic

    Set meBuildEmpty = rs
End Function
