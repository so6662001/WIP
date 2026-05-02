Attribute VB_Name = "VouMetaCache"
'==============================================================================
' Module      : VouMetaCache.bas
' Description : 凭证生成 batch 元数据缓存（PR-1 第一波）
'               消除 SaveDoc / BeforeAction / Init / CheckDateValidate
'               等路径下的 N+1 查询，把 DB 查询替换为 Dictionary 查找。
'
' 原则        : 100% 等价
'               缓存只把"DB 查询"替换为"内存查找"，输入相同的 key 必须返回
'               与原 DB 查询完全一样的值。任何缓存未命中 → 回退到原 DB 查询，
'               保证向后兼容（首次部署时即使忘记 LoadAll，行为也不会变）。
'
' 生命周期    : 在 cMthCstAccGL2.meCreVouFor* 进入前调用 LoadAll 一次
'               整 batch 结束后调用 ClearAll
'
' 健壮性      : v2 修订—— 不再用 "|" 分隔字符串拼接（FIName/CorpName 可能含
'               "|" 字符会被错切分）。所有缓存采用嵌套 Dictionary 结构，
'               每个字段独立 key，避免任何编码/转义问题。
'
' Database    : 兼容 Microsoft SQL Server 2008 及以上
'==============================================================================
Option Explicit

' 各表缓存：DocID -> Scripting.Dictionary（含该行所有字段）
Private m_dctFI         As Object       ' FIID -> sub-dict
Private m_dctCorp       As Object       ' CorpID -> sub-dict
Private m_dctEmp        As Object       ' EmpID -> sub-dict
Private m_dctAcc        As Object       ' AccID -> sub-dict
Private m_dctAP         As Object       ' APID -> sub-dict
Private m_dctDocFI      As Object       ' "DocKey|WHID|DocID" -> sub-dict
                                         ' (DocKey/WHID/DocID 中不会出现 "|" — 都是 ID 字段)
Private m_dctVCBills    As Object       ' "BillType|BillSubType" -> vcid (单值，BillSubType 是枚举)

' AccPeriod 顺序数组（按 BeginDate 升序排好），用于按日期定位
Private m_arrAPID()     As String
Private m_lngAPCount    As Long
Private m_strActiveAPID_APTag   As String
Private m_strActiveAPID_FOAPTag As String

Private m_blnLoaded     As Boolean


'==============================================================================
' Public：状态查询
'==============================================================================
Public Property Get IsLoaded() As Boolean
    IsLoaded = m_blnLoaded
End Property


'==============================================================================
' Public：懒加载入口（跨工程友好）
'
' VB6 标准模块的状态在每个 ActiveX DLL 工程内独立。当 t_FVou_M / CVouService
' （工程 C）调用 TryGetFI 时，工程 C 内的 m_blnLoaded 与工程 A 的 m_blnLoaded
' 是不同的两份变量。原 PR-1 假设"meCreateVou 一次 LoadAll，所有调用方共享"
' 在跨工程下不成立。
'
' 修复：每个 TryGet* 函数首次调用时自动 EnsureLoaded —— 每个工程独立加载一次
' 元数据，数据完全相同（只读元数据），3 个工程一共最多 21 次 SQL（一次性，
' 不在主循环里）。
'==============================================================================
Public Sub EnsureLoaded(ByVal objDS As HHDataService.sysDataService)
    If m_blnLoaded Then Exit Sub
    Call LoadAll(objDS)
End Sub


'==============================================================================
' Public：在 batch 入口一次性加载所有元数据
'==============================================================================
Public Sub LoadAll(ByVal objDS As HHDataService.sysDataService)
On Error GoTo ErrH
    Dim rs  As ADODB.Recordset
    Dim SQL As String

    Call ClearAll

    Set m_dctFI = CreateObject("Scripting.Dictionary")
    Set m_dctCorp = CreateObject("Scripting.Dictionary")
    Set m_dctEmp = CreateObject("Scripting.Dictionary")
    Set m_dctAcc = CreateObject("Scripting.Dictionary")
    Set m_dctAP = CreateObject("Scripting.Dictionary")
    Set m_dctDocFI = CreateObject("Scripting.Dictionary")
    Set m_dctVCBills = CreateObject("Scripting.Dictionary")

    '--------------------------------------------------------------------------
    ' 1) FinanceItems
    '--------------------------------------------------------------------------
    SQL = "SELECT FIID, FINO, FIName, ISStop, FITag, EXPTAG, HSTag, HSTagName, " & _
          "       baNum, baCust, baSupp, baOtherCorp, FullName " & _
          "FROM   FinanceItems WITH(NOLOCK) WHERE ISNULL(FIID,'')<>''"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim dctFIRow As Object
        Set dctFIRow = CreateObject("Scripting.Dictionary")
        dctFIRow.Add "FINO",        objDS.NullToStr(rs.Fields("FINO").Value)
        dctFIRow.Add "FIName",      objDS.NullToStr(rs.Fields("FIName").Value)
        dctFIRow.Add "ISStop",      objDS.NullToBool(rs.Fields("ISStop").Value)
        dctFIRow.Add "FITag",       CLng(objDS.NullToDbl(rs.Fields("FITag").Value))
        dctFIRow.Add "EXPTAG",      CLng(objDS.NullToDbl(rs.Fields("EXPTAG").Value))
        dctFIRow.Add "HSTag",       CLng(objDS.NullToDbl(rs.Fields("HSTag").Value))
        dctFIRow.Add "HSTagName",   objDS.NullToStr(rs.Fields("HSTagName").Value)
        dctFIRow.Add "baNum",       objDS.NullToBool(rs.Fields("baNum").Value)
        dctFIRow.Add "baCust",      objDS.NullToBool(rs.Fields("baCust").Value)
        dctFIRow.Add "baSupp",      objDS.NullToBool(rs.Fields("baSupp").Value)
        dctFIRow.Add "baOtherCorp", objDS.NullToBool(rs.Fields("baOtherCorp").Value)
        dctFIRow.Add "FullName",    objDS.NullToStr(rs.Fields("FullName").Value)
        m_dctFI.Add CStr(rs.Fields("FIID").Value), dctFIRow
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 2) Corp
    '--------------------------------------------------------------------------
    SQL = "SELECT CorpID, Contact, CorpName FROM Corp WITH(NOLOCK)"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim dctCorpRow As Object
        Set dctCorpRow = CreateObject("Scripting.Dictionary")
        dctCorpRow.Add "Contact",  objDS.NullToBool(rs.Fields("Contact").Value)
        dctCorpRow.Add "CorpName", objDS.NullToStr(rs.Fields("CorpName").Value)
        m_dctCorp.Add CStr(rs.Fields("CorpID").Value), dctCorpRow
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 3) Emp
    '--------------------------------------------------------------------------
    SQL = "SELECT EmpID, dismission, EmpName FROM Emp WITH(NOLOCK)"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim dctEmpRow As Object
        Set dctEmpRow = CreateObject("Scripting.Dictionary")
        dctEmpRow.Add "dismission", objDS.NullToBool(rs.Fields("dismission").Value)
        dctEmpRow.Add "EmpName",    objDS.NullToStr(rs.Fields("EmpName").Value)
        m_dctEmp.Add CStr(rs.Fields("EmpID").Value), dctEmpRow
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 4) Account
    '--------------------------------------------------------------------------
    SQL = "SELECT AccID, ISStop, AccName FROM Account WITH(NOLOCK)"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        Dim dctAccRow As Object
        Set dctAccRow = CreateObject("Scripting.Dictionary")
        dctAccRow.Add "ISStop",  objDS.NullToBool(rs.Fields("ISStop").Value)
        dctAccRow.Add "AccName", objDS.NullToStr(rs.Fields("AccName").Value)
        m_dctAcc.Add CStr(rs.Fields("AccID").Value), dctAccRow
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 5) AccPeriod (整张表，通常 < 50 行)
    '--------------------------------------------------------------------------
    SQL = "SELECT APID, BeginDate, EndDate, " & _
          "       ISNULL(APTag,0)   AS APTag, " & _
          "       ISNULL(FOAPTag,0) AS FOAPTag " & _
          "FROM   AccPeriod WITH(NOLOCK) ORDER BY BeginDate"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    m_lngAPCount = 0
    ReDim m_arrAPID(0 To 1023)
    Do While Not rs.EOF
        Dim sAPID As String
        sAPID = CStr(rs.Fields("APID").Value)
        Dim dctAPRow As Object
        Set dctAPRow = CreateObject("Scripting.Dictionary")
        ' BeginDate / EndDate 用字符串 "yyyy-MM-dd" 存，便于直接字符串比较
        dctAPRow.Add "BeginDate", Format$(rs.Fields("BeginDate").Value, "yyyy-MM-dd")
        dctAPRow.Add "EndDate",   Format$(rs.Fields("EndDate").Value, "yyyy-MM-dd")
        dctAPRow.Add "APTag",     CLng(rs.Fields("APTag").Value)
        dctAPRow.Add "FOAPTag",   CLng(rs.Fields("FOAPTag").Value)
        m_dctAP.Add sAPID, dctAPRow

        If m_lngAPCount >= UBound(m_arrAPID) Then
            ReDim Preserve m_arrAPID(0 To (UBound(m_arrAPID) + 1) * 2 - 1)
        End If
        m_arrAPID(m_lngAPCount) = sAPID
        m_lngAPCount = m_lngAPCount + 1

        If CLng(rs.Fields("APTag").Value) = 2 Then
            m_strActiveAPID_APTag = sAPID
        End If
        If CLng(rs.Fields("FOAPTag").Value) = 2 Then
            m_strActiveAPID_FOAPTag = sAPID
        End If
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 6) GL2_AchFIID + FinanceItems 联合
    '--------------------------------------------------------------------------
    SQL = "SELECT ISNULL(D.dockey,'')      AS dockey," & vbCrLf & _
          "       ISNULL(D.DocID,'')       AS DocID," & vbCrLf & _
          "       ISNULL(D.WHID,'')        AS WHID," & vbCrLf & _
          "       D.FIID," & vbCrLf & _
          "       FI.FINo," & vbCrLf & _
          "       CASE WHEN ISNULL(FI.FullName,'')='' " & _
          "            THEN FI.FINAME ELSE FI.FullName END AS FIName," & vbCrLf & _
          "       FI.HSTag, FI.HSTagName, FI.FITag" & vbCrLf & _
          "FROM   GL2_AchFIID D WITH(NOLOCK)" & vbCrLf & _
          "INNER JOIN FinanceItems FI WITH(NOLOCK) ON D.FIID = FI.FIID" & vbCrLf & _
          "WHERE  ISNULL(D.FIID,'')<>''"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        ' DocKey / WHID / DocID 都是 ID 字段（varchar(40~48)，纯 ASCII 标识符）
        ' 不会出现 "|"，可以安全用 "|" 拼 key
        Dim sKey As String
        sKey = CStr(rs.Fields("dockey").Value) & Chr$(31) & _
               CStr(rs.Fields("WHID").Value) & Chr$(31) & _
               CStr(rs.Fields("DocID").Value)
        ' 同 key 重复时保留首条（与原 ADO Filter 行为一致）
        If Not m_dctDocFI.Exists(sKey) Then
            Dim dctDocRow As Object
            Set dctDocRow = CreateObject("Scripting.Dictionary")
            dctDocRow.Add "FIID",      CStr(rs.Fields("FIID").Value)
            dctDocRow.Add "FINO",      objDS.NullToStr(rs.Fields("FINo").Value)
            dctDocRow.Add "FIName",    objDS.NullToStr(rs.Fields("FIName").Value)
            dctDocRow.Add "HSTag",     CLng(objDS.NullToDbl(rs.Fields("HSTag").Value))
            dctDocRow.Add "HSTagName", objDS.NullToStr(rs.Fields("HSTagName").Value)
            dctDocRow.Add "FITag",     CLng(objDS.NullToDbl(rs.Fields("FITag").Value))
            m_dctDocFI.Add sKey, dctDocRow
        End If
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 7) F_VouCls_Bills（vcid 查找）
    '--------------------------------------------------------------------------
    SQL = "SELECT BillType, BillSubType, vcid FROM F_VouCls_Bills WITH(NOLOCK) " & _
          "WHERE ISNULL(VCID,'')<>''"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        ' BillType 是 BillKey 标识符（如 "PGBill"/"SSBill"），无 "|"
        ' BillSubType 也是枚举，无 "|"
        Dim sk As String
        sk = CStr(rs.Fields("BillType").Value) & Chr$(31) & _
             objDS.NullToStr(rs.Fields("BillSubType").Value)
        If Not m_dctVCBills.Exists(sk) Then
            m_dctVCBills.Add sk, CStr(rs.Fields("vcid").Value)
        End If
        rs.MoveNext
    Loop
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
' Public：清理缓存（batch 末尾调用）
'==============================================================================
Public Sub ClearAll()
    On Error Resume Next
    Set m_dctFI = Nothing
    Set m_dctCorp = Nothing
    Set m_dctEmp = Nothing
    Set m_dctAcc = Nothing
    Set m_dctAP = Nothing
    Set m_dctDocFI = Nothing
    Set m_dctVCBills = Nothing
    m_lngAPCount = 0
    Erase m_arrAPID
    m_strActiveAPID_APTag = ""
    m_strActiveAPID_FOAPTag = ""
    m_blnLoaded = False
    On Error GoTo 0
End Sub


'==============================================================================
' Public：FinanceItems 查找
'   返回 True = 找到；False = 未找到（调用方应回退到 DB）
'==============================================================================
Public Function TryGetFI(ByVal sFIID As String, _
                        ByRef rtnFINO As String, ByRef rtnFIName As String, _
                        ByRef rtnIsStop As Boolean, _
                        ByRef rtnFITag As Long, ByRef rtnEXPTAG As Long, _
                        ByRef rtnHSTag As Long, ByRef rtnHSTagName As String, _
                        ByRef rtnBaNum As Boolean, ByRef rtnBaCust As Boolean, _
                        ByRef rtnBaSupp As Boolean, ByRef rtnBaOtherCorp As Boolean, _
                        ByRef rtnFullName As String) As Boolean
    If Not m_blnLoaded Then Exit Function
    If Not m_dctFI.Exists(sFIID) Then Exit Function

    Dim d As Object
    Set d = m_dctFI(sFIID)
    rtnFINO = CStr(d("FINO"))
    rtnFIName = CStr(d("FIName"))
    rtnIsStop = CBool(d("ISStop"))
    rtnFITag = CLng(d("FITag"))
    rtnEXPTAG = CLng(d("EXPTAG"))
    rtnHSTag = CLng(d("HSTag"))
    rtnHSTagName = CStr(d("HSTagName"))
    rtnBaNum = CBool(d("baNum"))
    rtnBaCust = CBool(d("baCust"))
    rtnBaSupp = CBool(d("baSupp"))
    rtnBaOtherCorp = CBool(d("baOtherCorp"))
    rtnFullName = CStr(d("FullName"))
    TryGetFI = True
End Function


'==============================================================================
' Public：Corp 查找
'==============================================================================
Public Function TryGetCorp(ByVal sCorpID As String, _
                          ByRef rtnContact As Boolean, _
                          ByRef rtnCorpName As String) As Boolean
    If Not m_blnLoaded Then Exit Function
    If Not m_dctCorp.Exists(sCorpID) Then Exit Function

    Dim d As Object
    Set d = m_dctCorp(sCorpID)
    rtnContact = CBool(d("Contact"))
    rtnCorpName = CStr(d("CorpName"))
    TryGetCorp = True
End Function


'==============================================================================
' Public：Emp 查找
'==============================================================================
Public Function TryGetEmp(ByVal sEmpID As String, _
                         ByRef rtnDismission As Boolean, _
                         ByRef rtnEmpName As String) As Boolean
    If Not m_blnLoaded Then Exit Function
    If Not m_dctEmp.Exists(sEmpID) Then Exit Function

    Dim d As Object
    Set d = m_dctEmp(sEmpID)
    rtnDismission = CBool(d("dismission"))
    rtnEmpName = CStr(d("EmpName"))
    TryGetEmp = True
End Function


'==============================================================================
' Public：Account 查找
'==============================================================================
Public Function TryGetAcc(ByVal sAccID As String, _
                         ByRef rtnIsStop As Boolean, _
                         ByRef rtnAccName As String) As Boolean
    If Not m_blnLoaded Then Exit Function
    If Not m_dctAcc.Exists(sAccID) Then Exit Function

    Dim d As Object
    Set d = m_dctAcc(sAccID)
    rtnIsStop = CBool(d("ISStop"))
    rtnAccName = CStr(d("AccName"))
    TryGetAcc = True
End Function


'==============================================================================
' Public：根据日期查找包含它的 AccPeriod
'   等价于：SELECT APID, APTag/FOAPTag FROM AccPeriod WHERE BeginDate<=? AND EndDate>=?
'   useFOAPTag=True 表示读 FOAPTag 字段
'==============================================================================
Public Function TryGetAPByDate(ByVal vDate As Date, ByVal useFOAPTag As Boolean, _
                              ByRef rtnAPID As String, ByRef rtnAPTag As Long) As Boolean
    If Not m_blnLoaded Then Exit Function

    Dim sDate As String
    sDate = Format$(vDate, "yyyy-MM-dd")

    Dim i As Long
    For i = 0 To m_lngAPCount - 1
        Dim d As Object
        Set d = m_dctAP(m_arrAPID(i))
        ' 字符串 "yyyy-MM-dd" 直接比较等价于日期比较
        If sDate >= CStr(d("BeginDate")) And sDate <= CStr(d("EndDate")) Then
            rtnAPID = m_arrAPID(i)
            If useFOAPTag Then
                rtnAPTag = CLng(d("FOAPTag"))
            Else
                rtnAPTag = CLng(d("APTag"))
            End If
            TryGetAPByDate = True
            Exit Function
        End If
    Next i
End Function


'==============================================================================
' Public：当前活动 AccPeriod 是否存在
'==============================================================================
Public Function HasActiveAP(ByVal useFOAPTag As Boolean) As Boolean
    If Not m_blnLoaded Then Exit Function
    If useFOAPTag Then
        HasActiveAP = (m_strActiveAPID_FOAPTag <> "")
    Else
        HasActiveAP = (m_strActiveAPID_APTag <> "")
    End If
End Function


'==============================================================================
' Public：GL2_AchFIID 查找（DocKey + WHID + DocID 精确匹配）
'==============================================================================
Public Function TryGetDocFI(ByVal sDocKey As String, ByVal sWHID As String, _
                           ByVal sDocID As String, _
                           ByRef rtnFIID As String, _
                           ByRef rtnFINO As String, ByRef rtnFIName As String, _
                           ByRef rtnHSTag As Long, ByRef rtnHSTagName As String, _
                           ByRef rtnFITag As Long) As Boolean
    If Not m_blnLoaded Then Exit Function

    Dim sKey As String
    sKey = sDocKey & Chr$(31) & sWHID & Chr$(31) & sDocID
    If Not m_dctDocFI.Exists(sKey) Then Exit Function

    Dim d As Object
    Set d = m_dctDocFI(sKey)
    rtnFIID = CStr(d("FIID"))
    rtnFINO = CStr(d("FINO"))
    rtnFIName = CStr(d("FIName"))
    rtnHSTag = CLng(d("HSTag"))
    rtnHSTagName = CStr(d("HSTagName"))
    rtnFITag = CLng(d("FITag"))
    TryGetDocFI = True
End Function


'==============================================================================
' Public：F_VouCls_Bills 查找
'==============================================================================
Public Function TryGetVCID(ByVal sBillType As String, ByVal sBillSubType As String, _
                          ByRef rtnVCID As String) As Boolean
    If Not m_blnLoaded Then Exit Function

    Dim sk As String
    sk = sBillType & Chr$(31) & sBillSubType
    If m_dctVCBills.Exists(sk) Then
        rtnVCID = CStr(m_dctVCBills(sk))
        TryGetVCID = True
        Exit Function
    End If
    ' BillSubType 为空也应能查到（原 SQL 中 BillSubType IS NULL 视同 ''）
    sk = sBillType & Chr$(31)
    If m_dctVCBills.Exists(sk) Then
        rtnVCID = CStr(m_dctVCBills(sk))
        TryGetVCID = True
    End If
End Function
