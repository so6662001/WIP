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
' Database    : 兼容 Microsoft SQL Server 2008 及以上
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' FinanceItems：FIID -> "FINO|FIName|ISStop|FITag|EXPTAG|HSTag|HSTagName|baNum|baCust|baSupp|baOtherCorp|FullName"
'   ISStop / baNum / baCust / baSupp / baOtherCorp 用 "0"/"1" 表示 Boolean
'   FITag / EXPTAG / HSTag 转为 Long 字符串
'------------------------------------------------------------------------------
Private m_dctFI         As Object

'------------------------------------------------------------------------------
' Corp：CorpID -> "Contact|CorpName"
'------------------------------------------------------------------------------
Private m_dctCorp       As Object

'------------------------------------------------------------------------------
' Emp：EmpID -> "dismission|EmpName"
'------------------------------------------------------------------------------
Private m_dctEmp        As Object

'------------------------------------------------------------------------------
' Account：AccID -> "ISStop|AccName"
'------------------------------------------------------------------------------
Private m_dctAcc        As Object

'------------------------------------------------------------------------------
' AccPeriod 全表：APID -> "BeginDate(yyyy-MM-dd)|EndDate(yyyy-MM-dd)|APTag|FOAPTag"
'------------------------------------------------------------------------------
Private m_dctAP         As Object
'   按 BeginDate 升序排好的 APID 数组（用于线性查找日期落在哪个期间）
'   AccPeriod 通常 < 50 行，线性扫描足够
Private m_arrAPID()     As String
Private m_lngAPCount    As Long
'   active period（APTag=2 或 FOAPTag=2 的 APID）
Private m_strActiveAPID_APTag   As String
Private m_strActiveAPID_FOAPTag As String

'------------------------------------------------------------------------------
' GL2_AchFIID + FinanceItems 联合：'DocKey|WHID|DocID' -> "FIID|FINO|FIName|HSTag|HSTagName|FITag"
'   原 t_FVou_M.Init() 里加载到 mrsDocFIRel 的内容
'   NULL 在 key 中统一规范化为 ''
'------------------------------------------------------------------------------
Private m_dctDocFI      As Object

'------------------------------------------------------------------------------
' F_VouCls_Bills：'BillType|BillSubType' -> "vcid"
'   t_FVou_M.CreateVou 里查 vcid 用
'------------------------------------------------------------------------------
Private m_dctVCBills    As Object

'------------------------------------------------------------------------------
' 状态
'------------------------------------------------------------------------------
Private m_blnLoaded     As Boolean


'==============================================================================
' Public：状态查询
'==============================================================================
Public Property Get IsLoaded() As Boolean
    IsLoaded = m_blnLoaded
End Property


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
        m_dctFI.Add CStr(rs.Fields("FIID").Value), _
            objDS.NullToStr(rs.Fields("FINO").Value) & "|" & _
            objDS.NullToStr(rs.Fields("FIName").Value) & "|" & _
            IIf(objDS.NullToBool(rs.Fields("ISStop").Value), "1", "0") & "|" & _
            CStr(CLng(objDS.NullToDbl(rs.Fields("FITag").Value))) & "|" & _
            CStr(CLng(objDS.NullToDbl(rs.Fields("EXPTAG").Value))) & "|" & _
            CStr(CLng(objDS.NullToDbl(rs.Fields("HSTag").Value))) & "|" & _
            objDS.NullToStr(rs.Fields("HSTagName").Value) & "|" & _
            IIf(objDS.NullToBool(rs.Fields("baNum").Value), "1", "0") & "|" & _
            IIf(objDS.NullToBool(rs.Fields("baCust").Value), "1", "0") & "|" & _
            IIf(objDS.NullToBool(rs.Fields("baSupp").Value), "1", "0") & "|" & _
            IIf(objDS.NullToBool(rs.Fields("baOtherCorp").Value), "1", "0") & "|" & _
            objDS.NullToStr(rs.Fields("FullName").Value)
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 2) Corp
    '--------------------------------------------------------------------------
    SQL = "SELECT CorpID, Contact, CorpName FROM Corp WITH(NOLOCK)"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        m_dctCorp.Add CStr(rs.Fields("CorpID").Value), _
            IIf(objDS.NullToBool(rs.Fields("Contact").Value), "1", "0") & "|" & _
            objDS.NullToStr(rs.Fields("CorpName").Value)
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 3) Emp
    '--------------------------------------------------------------------------
    SQL = "SELECT EmpID, dismission, EmpName FROM Emp WITH(NOLOCK)"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        m_dctEmp.Add CStr(rs.Fields("EmpID").Value), _
            IIf(objDS.NullToBool(rs.Fields("dismission").Value), "1", "0") & "|" & _
            objDS.NullToStr(rs.Fields("EmpName").Value)
        rs.MoveNext
    Loop
    Call objDS.rs_Close(rs)

    '--------------------------------------------------------------------------
    ' 4) Account
    '--------------------------------------------------------------------------
    SQL = "SELECT AccID, ISStop, AccName FROM Account WITH(NOLOCK)"
    Set rs = objDS.OpenRecordsetBySQL(SQL, True, True)
    Do While Not rs.EOF
        m_dctAcc.Add CStr(rs.Fields("AccID").Value), _
            IIf(objDS.NullToBool(rs.Fields("ISStop").Value), "1", "0") & "|" & _
            objDS.NullToStr(rs.Fields("AccName").Value)
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
    ReDim m_arrAPID(0 To 1023)                          ' 预分配
    Do While Not rs.EOF
        Dim sAPID As String
        sAPID = CStr(rs.Fields("APID").Value)
        m_dctAP.Add sAPID, _
            Format$(rs.Fields("BeginDate").Value, "yyyy-MM-dd") & "|" & _
            Format$(rs.Fields("EndDate").Value, "yyyy-MM-dd") & "|" & _
            CStr(CLng(rs.Fields("APTag").Value)) & "|" & _
            CStr(CLng(rs.Fields("FOAPTag").Value))
        If m_lngAPCount > UBound(m_arrAPID) Then
            ReDim Preserve m_arrAPID(0 To UBound(m_arrAPID) * 2)
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
    ' 6) GL2_AchFIID + FinanceItems 联合（同 t_FVou_M.Init）
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
        Dim sKey As String
        sKey = CStr(rs.Fields("dockey").Value) & "|" & _
               CStr(rs.Fields("WHID").Value) & "|" & _
               CStr(rs.Fields("DocID").Value)
        ' 同 key 重复时保留首条（与原 ADO Filter 行为一致：取第一行命中）
        If Not m_dctDocFI.Exists(sKey) Then
            m_dctDocFI.Add sKey, _
                CStr(rs.Fields("FIID").Value) & "|" & _
                objDS.NullToStr(rs.Fields("FINo").Value) & "|" & _
                objDS.NullToStr(rs.Fields("FIName").Value) & "|" & _
                CStr(CLng(objDS.NullToDbl(rs.Fields("HSTag").Value))) & "|" & _
                objDS.NullToStr(rs.Fields("HSTagName").Value) & "|" & _
                CStr(CLng(objDS.NullToDbl(rs.Fields("FITag").Value)))
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
        Dim sk As String
        sk = CStr(rs.Fields("BillType").Value) & "|" & _
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
    Call ClearAll
    Call Err.Raise(Err.Number, , Err.Description)
End Sub


'==============================================================================
' Public：清理缓存（batch 末尾调用）
'==============================================================================
Public Sub ClearAll()
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
End Sub


'==============================================================================
' Public：FinanceItems 查找
'   返回 True = 找到，并通过 ByRef 输出各字段
'   返回 False = 未找到（调用方应回退到 DB 或抛错，与原行为一致）
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

    Dim p() As String
    p = Split(CStr(m_dctFI(sFIID)), "|")
    rtnFINO = p(0)
    rtnFIName = p(1)
    rtnIsStop = (p(2) = "1")
    rtnFITag = CLng(p(3))
    rtnEXPTAG = CLng(p(4))
    rtnHSTag = CLng(p(5))
    rtnHSTagName = p(6)
    rtnBaNum = (p(7) = "1")
    rtnBaCust = (p(8) = "1")
    rtnBaSupp = (p(9) = "1")
    rtnBaOtherCorp = (p(10) = "1")
    If UBound(p) >= 11 Then rtnFullName = p(11)
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

    Dim p() As String
    p = Split(CStr(m_dctCorp(sCorpID)), "|")
    rtnContact = (p(0) = "1")
    rtnCorpName = p(1)
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

    Dim p() As String
    p = Split(CStr(m_dctEmp(sEmpID)), "|")
    rtnDismission = (p(0) = "1")
    rtnEmpName = p(1)
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

    Dim p() As String
    p = Split(CStr(m_dctAcc(sAccID)), "|")
    rtnIsStop = (p(0) = "1")
    rtnAccName = p(1)
    TryGetAcc = True
End Function


'==============================================================================
' Public：根据日期查找包含它的 AccPeriod
'   等价于：SELECT APID, APTag/FOAPTag FROM AccPeriod WHERE BeginDate<=? AND EndDate>=?
'   useFOAPTag=True 表示读 FOAPTag 字段（GL2_ERPFODiffCarry=True 场景）
'==============================================================================
Public Function TryGetAPByDate(ByVal vDate As Date, ByVal useFOAPTag As Boolean, _
                              ByRef rtnAPID As String, ByRef rtnAPTag As Long) As Boolean
    If Not m_blnLoaded Then Exit Function

    Dim sDate As String
    sDate = Format$(vDate, "yyyy-MM-dd")

    ' 线性扫描（AccPeriod 通常 < 50 行）
    Dim i As Long
    For i = 0 To m_lngAPCount - 1
        Dim p() As String
        p = Split(CStr(m_dctAP(m_arrAPID(i))), "|")
        ' p(0)=BeginDate, p(1)=EndDate, p(2)=APTag, p(3)=FOAPTag
        If sDate >= p(0) And sDate <= p(1) Then
            rtnAPID = m_arrAPID(i)
            If useFOAPTag Then
                rtnAPTag = CLng(p(3))
            Else
                rtnAPTag = CLng(p(2))
            End If
            TryGetAPByDate = True
            Exit Function
        End If
    Next i
End Function


'==============================================================================
' Public：当前活动 AccPeriod 是否存在（APTag=2 或 FOAPTag=2）
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
    sKey = sDocKey & "|" & sWHID & "|" & sDocID
    If Not m_dctDocFI.Exists(sKey) Then Exit Function

    Dim p() As String
    p = Split(CStr(m_dctDocFI(sKey)), "|")
    rtnFIID = p(0)
    rtnFINO = p(1)
    rtnFIName = p(2)
    rtnHSTag = CLng(p(3))
    rtnHSTagName = p(4)
    rtnFITag = CLng(p(5))
    TryGetDocFI = True
End Function


'==============================================================================
' Public：F_VouCls_Bills 查找
'==============================================================================
Public Function TryGetVCID(ByVal sBillType As String, ByVal sBillSubType As String, _
                          ByRef rtnVCID As String) As Boolean
    If Not m_blnLoaded Then Exit Function

    Dim sk As String
    sk = sBillType & "|" & sBillSubType
    If m_dctVCBills.Exists(sk) Then
        rtnVCID = CStr(m_dctVCBills(sk))
        TryGetVCID = True
        Exit Function
    End If
    ' BillSubType 为空也应能查到（原 SQL 中 BillSubType IS NULL 视同 ''）
    sk = sBillType & "|"
    If m_dctVCBills.Exists(sk) Then
        rtnVCID = CStr(m_dctVCBills(sk))
        TryGetVCID = True
    End If
End Function
