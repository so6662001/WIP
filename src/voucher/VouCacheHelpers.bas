Attribute VB_Name = "VouCacheHelpers"
'==============================================================================
' Module      : VouCacheHelpers.bas
' Description : 凭证 batch 元数据缓存的"等价回退"辅助 helper（PR-1 第一波）
'
'   用法：
'     原 CVouService.BeforeAction 内的 4 段 IN 列表 SELECT + 后续 Filter
'     全部替换为对本模块 helper 的调用：
'       Call VouCacheHelpers.CheckFI   (objDS, ItemsData, blnUseCache, rsfcode, rsFIBak, ...)
'       Call VouCacheHelpers.CheckCorp (objDS, ItemsData, blnUseCache, rsCorps, ...)
'       Call VouCacheHelpers.CheckEmp  (objDS, ItemsData, blnUseCache, rsEmps, ...)
'       Call VouCacheHelpers.CheckAcc  (objDS, ItemsData, blnUseCache, rsAccs, ...)
'
'   两条路径（cache / 非 cache）的输出必须 100% 一致：
'     - ItemsData 字段写回值
'     - 错误信息文案（strErrInfo 累积）
'     - rsFIBak 的内存缓存写入（与原代码一致）
'==============================================================================
Option Explicit

'==============================================================================
' Public：FI 检查 + 字段写回（替换 BeforeAction 内的 rsfcode.Filter 段）
'==============================================================================
Public Sub CheckFI(ByVal objDS As HHDataService.sysDataService, _
                   ByVal ItemsData As ADODB.Recordset, _
                   ByVal blnUseCache As Boolean, _
                   ByRef rsfcode As ADODB.Recordset, _
                   ByRef rsFIBak As ADODB.Recordset, _
                   ByVal ActCHName As String, _
                   ByVal strBillInfo As String, _
                   ByRef strErrInfo As String)
    Dim sFIID As String
    sFIID = objDS.NullToStr(ItemsData.Fields("FIID").Value)
    
    Dim blnNewFI As Boolean
    blnNewFI = True
    
    If Not rsFIBak Is Nothing Then
        rsFIBak.Filter = "fiid='" & sFIID & "'"
        If rsFIBak.RecordCount > 0 Then blnNewFI = False
    End If
    
    Dim sFINO As String, sFIName As String, sHSTagName As String, sFullName As String
    Dim blnIsStop As Boolean, blnBaNum As Boolean, blnBaCust As Boolean
    Dim blnBaSupp As Boolean, blnBaOtherCorp As Boolean
    Dim lngFITag As Long, lngEXPTAG As Long, lngHSTag As Long
    Dim blnFound As Boolean
    
    If blnNewFI Then
        If blnUseCache Then
            blnFound = VouMetaCache.TryGetFI(sFIID, sFINO, sFIName, blnIsStop, _
                lngFITag, lngEXPTAG, lngHSTag, sHSTagName, _
                blnBaNum, blnBaCust, blnBaSupp, blnBaOtherCorp, sFullName)
        Else
            rsfcode.Filter = "fiid='" & sFIID & "'"
            blnFound = (rsfcode.RecordCount > 0)
            If blnFound Then
                sFINO = objDS.NullToStr(rsfcode.Fields("FINO").Value)
                sFIName = objDS.NullToStr(rsfcode.Fields("FIName").Value)
                blnIsStop = objDS.NullToBool(rsfcode.Fields("isStop").Value)
                lngFITag = CLng(objDS.NullToDbl(rsfcode.Fields("FITag").Value))
                lngEXPTAG = CLng(objDS.NullToDbl(rsfcode.Fields("EXPTAG").Value))
                lngHSTag = CInt(objDS.NullToDbl(rsfcode.Fields("hstag").Value))
                sHSTagName = objDS.NullToStr(rsfcode.Fields("hstagname").Value)
                blnBaNum = objDS.NullToBool(rsfcode.Fields("baNum").Value)
                blnBaCust = objDS.NullToBool(rsfcode.Fields("baCust").Value)
                blnBaSupp = objDS.NullToBool(rsfcode.Fields("baSupp").Value)
                blnBaOtherCorp = objDS.NullToBool(rsfcode.Fields("baOtherCorp").Value)
                sFullName = objDS.NullToStr(rsfcode.Fields("fullname").Value)
            End If
        End If
        
        If Not blnFound Then
            strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
                " 条分录中的会计科目[" & ItemsData.Fields("FINO").Value & _
                "]在科目档案中已经找不到，可能被删除！" & vbCrLf
            Exit Sub
        End If
        
        If blnIsStop Then
            strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
                " 条分录中的会计科目[" & ItemsData.Fields("FINO").Value & _
                "]已经被停用，无法" & ActCHName & "！" & vbCrLf
        End If
        
        If lngFITag = 6 Then
            If lngEXPTAG = 1 Then
                strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
                    " 条分录中的会计科目[" & ItemsData.Fields("FINO").Value & _
                    "]是采购费用类科目，无法" & ActCHName & _
                    "，请修改成非采购费用类科目！" & vbCrLf
            End If
        End If
        
        If objDS.NullToStr(ItemsData.Fields("FIName").Value) = "" Then
            ItemsData.Fields("FIName").Value = Replace(IIf(sFullName = "", sFIName, sFullName), vbCrLf, "")
        End If
        ItemsData.Fields("FINO").Value = sFINO
        ItemsData.Fields("hstag").Value = CInt(lngHSTag)
        ItemsData.Fields("hstagname").Value = sHSTagName
        ItemsData.Fields("FITag").Value = CDbl(lngFITag)
        ItemsData.Fields("baNum").Value = blnBaNum
        ItemsData.Fields("baCust").Value = blnBaCust
        ItemsData.Fields("baSupp").Value = blnBaSupp
        ItemsData.Fields("baOtherCorp").Value = blnBaOtherCorp
        
        ' 写入 rsFIBak（与原代码一致：跨 row 缓存）
        If rsFIBak Is Nothing Then
            ' 原代码用 objDS.CopyRsStruct(rsfcode)；当 cache 路径无 rsfcode 时
            ' 我们手动构造一个等价 schema
            Set rsFIBak = meCreateFIBakRs()
        End If
        rsFIBak.AddNew
        rsFIBak.Fields("fiid").Value = sFIID
        rsFIBak.Fields("fino").Value = sFINO
        rsFIBak.Fields("finame").Value = ItemsData.Fields("FIName").Value
        rsFIBak.Fields("hstag").Value = CInt(lngHSTag)
        rsFIBak.Fields("hstagname").Value = sHSTagName
        rsFIBak.Fields("FITag").Value = CDbl(lngFITag)
        rsFIBak.Fields("baNum").Value = blnBaNum
        rsFIBak.Fields("baCust").Value = blnBaCust
        rsFIBak.Fields("baSupp").Value = blnBaSupp
        rsFIBak.Fields("baOtherCorp").Value = blnBaOtherCorp
        rsFIBak.Update
    Else
        ' 命中 rsFIBak（原代码也走这条 else 分支）
        If objDS.NullToStr(ItemsData.Fields("FIName").Value) = "" Then
            ItemsData.Fields("FIName").Value = objDS.NullToStr(rsFIBak.Fields("FIName").Value)
        End If
        ItemsData.Fields("FINO").Value = objDS.NullToStr(rsFIBak.Fields("FINO").Value)
        ItemsData.Fields("hstag").Value = CInt(objDS.NullToDbl(rsFIBak.Fields("hstag").Value))
        ItemsData.Fields("hstagname").Value = objDS.NullToStr(rsFIBak.Fields("hstagname").Value)
        ItemsData.Fields("FITag").Value = objDS.NullToDbl(rsFIBak.Fields("FITag").Value)
        ItemsData.Fields("baNum").Value = objDS.NullToBool(rsFIBak.Fields("baNum").Value)
        ItemsData.Fields("baCust").Value = objDS.NullToBool(rsFIBak.Fields("baCust").Value)
        ItemsData.Fields("baSupp").Value = objDS.NullToBool(rsFIBak.Fields("baSupp").Value)
        ItemsData.Fields("baOtherCorp").Value = objDS.NullToBool(rsFIBak.Fields("baOtherCorp").Value)
    End If
End Sub


'==============================================================================
' Public：Account 检查（替换 BeforeAction 内的 rsAccs.Filter 段）
'==============================================================================
Public Sub CheckAcc(ByVal objDS As HHDataService.sysDataService, _
                    ByVal ItemsData As ADODB.Recordset, _
                    ByVal blnUseCache As Boolean, _
                    ByRef rsAccs As ADODB.Recordset, _
                    ByVal ActCHName As String, _
                    ByVal strBillInfo As String, _
                    ByRef strErrInfo As String)
    Dim sAccID As String
    sAccID = objDS.NullToStr(ItemsData.Fields("AccID").Value)
    If sAccID = "" Then Exit Sub
    
    Dim blnFound As Boolean
    Dim blnIsStop As Boolean
    Dim sAccName As String
    
    If blnUseCache Then
        blnFound = VouMetaCache.TryGetAcc(sAccID, blnIsStop, sAccName)
    Else
        rsAccs.Filter = "accid='" & sAccID & "'"
        blnFound = (rsAccs.RecordCount > 0)
        If blnFound Then
            blnIsStop = objDS.NullToBool(rsAccs.Fields("isStop").Value)
            sAccName = objDS.NullToStr(rsAccs.Fields("AccName").Value)
        End If
    End If
    
    If Not blnFound Then
        strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
            " 条分录[" & ItemsData.Fields("FINO").Value & "]中的帐户[" & sAccName & _
            "]已经找不到，无法" & ActCHName & "！"
        Exit Sub
    End If
    If blnIsStop Then
        strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
            " 条分录[" & ItemsData.Fields("FINO").Value & "]中的帐户[" & sAccName & _
            "]已经被停用，无法" & ActCHName & "！"
    End If
End Sub


'==============================================================================
' Public：Corp 检查
'   strCTagName = 客户/供应商/相关单位 由调用方根据 corptag 决定
'==============================================================================
Public Sub CheckCorp(ByVal objDS As HHDataService.sysDataService, _
                     ByVal ItemsData As ADODB.Recordset, _
                     ByVal blnUseCache As Boolean, _
                     ByRef rsCorps As ADODB.Recordset, _
                     ByVal ActCHName As String, _
                     ByVal strBillInfo As String, _
                     ByRef strErrInfo As String)
    Dim sCorpID As String
    sCorpID = objDS.NullToStr(ItemsData.Fields("CorpID").Value)
    If sCorpID = "" Then Exit Sub
    
    Dim sCTagName As String
    Select Case CLng(objDS.NullToDbl(ItemsData.Fields("corptag").Value))
        Case 3: sCTagName = "客户"
        Case 4: sCTagName = "供应商"
        Case 5: sCTagName = "相关单位"
    End Select
    
    Dim blnFound As Boolean, blnContact As Boolean, sCorpName As String
    
    If blnUseCache Then
        blnFound = VouMetaCache.TryGetCorp(sCorpID, blnContact, sCorpName)
    Else
        rsCorps.Filter = "Corpid='" & sCorpID & "'"
        blnFound = (rsCorps.RecordCount > 0)
        If blnFound Then
            blnContact = objDS.NullToBool(rsCorps.Fields("Contact").Value)
            sCorpName = objDS.NullToStr(rsCorps.Fields("CorpName").Value)
        End If
    End If
    
    If Not blnFound Then
        strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
            " 条分录[" & ItemsData.Fields("FINO").Value & "]中的" & sCTagName & "[" & sCorpName & _
            "]已经找不到，无法" & ActCHName & "！" & vbCrLf
        Exit Sub
    End If
    If Not blnContact Then
        strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
            " 条分录[" & ItemsData.Fields("FINO").Value & "]中的" & sCTagName & " [" & sCorpName & _
            "]已经被停用，无法" & ActCHName & "！" & vbCrLf
    End If
End Sub


'==============================================================================
' Public：Emp 检查
'==============================================================================
Public Sub CheckEmp(ByVal objDS As HHDataService.sysDataService, _
                    ByVal ItemsData As ADODB.Recordset, _
                    ByVal blnUseCache As Boolean, _
                    ByRef rsEmps As ADODB.Recordset, _
                    ByVal ActCHName As String, _
                    ByVal strBillInfo As String, _
                    ByRef strErrInfo As String)
    Dim sEmpID As String
    sEmpID = objDS.NullToStr(ItemsData.Fields("empid").Value)
    If sEmpID = "" Then Exit Sub
    
    Dim blnFound As Boolean, blnDismission As Boolean, sEmpName As String
    
    If blnUseCache Then
        blnFound = VouMetaCache.TryGetEmp(sEmpID, blnDismission, sEmpName)
    Else
        rsEmps.Filter = "empid='" & sEmpID & "'"
        blnFound = (rsEmps.RecordCount > 0)
        If blnFound Then
            blnDismission = objDS.NullToBool(rsEmps.Fields("dismission").Value)
            sEmpName = objDS.NullToStr(rsEmps.Fields("EmpName").Value)
        End If
    End If
    
    ' 注意：原代码这里 rsEmps.RecordCount=0 时不报错（个人找不到不致命）
    ' 但 dismission=True 时报错。保留原行为。
    If Not blnFound Then Exit Sub
    
    If blnDismission Then
        strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
            " 条分录[" & ItemsData.Fields("FINO").Value & "]中的个人[" & sEmpName & _
            "]已经被停用，无法" & ActCHName & "！" & vbCrLf
    End If
End Sub


'==============================================================================
' Private：构造与原 rsfcode 等价 schema 的内存 rsFIBak
'==============================================================================
Private Function meCreateFIBakRs() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Fields.Append "fiid",         adVarChar, 48,  adFldIsNullable
    rs.Fields.Append "fino",         adVarChar, 256, adFldIsNullable
    rs.Fields.Append "finame",       adVarChar, 512, adFldIsNullable
    rs.Fields.Append "hstag",        adInteger, , adFldIsNullable
    rs.Fields.Append "hstagname",    adVarChar, 64,  adFldIsNullable
    rs.Fields.Append "FITag",        adInteger, , adFldIsNullable
    rs.Fields.Append "baNum",        adBoolean, , adFldIsNullable
    rs.Fields.Append "baCust",       adBoolean, , adFldIsNullable
    rs.Fields.Append "baSupp",       adBoolean, , adFldIsNullable
    rs.Fields.Append "baOtherCorp",  adBoolean, , adFldIsNullable
    rs.CursorLocation = adUseClient
    rs.Open , , adOpenKeyset, adLockOptimistic
    Set meCreateFIBakRs = rs
End Function
