Attribute VB_Name = "VouCacheHelpers"
'==============================================================================
' Module      : VouCacheHelpers.bas
' Description : 凭证 batch 元数据缓存的"等价回退"辅助 helper（PR-1 第一波）
'
' 提供两组 helper：
'
'   【组 A】CVouService.BeforeAction 元数据校验：
'     - CheckFI   (cache / db 双路径 → 输出 100% 等价于原 BeforeAction)
'     - CheckCorp (同上)
'     - CheckEmp  (同上)
'     - CheckAcc  (同上)
'
'   【组 B】t_FVou_M.GetFIIDByPrdt* 等 16 个 FI 查询函数：
'     - LookupDocFI_Strict (5 步 fallback：精确 → WHID → ClsID → ParentCls → 全空)
'     - LookupDocFI_NoStrict (4 步 fallback，找不到不抛错)
'     - LookupDocFI_KeyOnly (DocKey + DocID 两步，用于 Account/CGExpFI/TInv* 类)
'     这三个 helper 把原 t_FVou_M 中**全部** GetFIIDBy* 函数的 cache 路径统一封装；
'     调用方只需在函数入口加一行：If LookupDocFI_*(...) Then Exit Function。
'
'   两条路径（cache / db）的输出必须 100% 等价：
'     - 返回字段值
'     - 错误信息文案
'     - 异常类型（包括原代码 bug 引发的 ADO 异常 — 复刻保留）
'==============================================================================
Option Explicit

'==============================================================================
' Public：FI 检查 + 字段写回（替换 BeforeAction 内的 rsfcode.Filter 段）
'
' 原代码顺序 (BeforeAction 一次循环内)：
'   1. 判断 rsFIBak 是否命中
'   2. 写 ItemsData 字段（FIName/FINO/hstag/hstagname/FITag/baNum/baCust/baSupp/baOtherCorp）
'   3. ItemsData.Update    ← 由调用方负责
'   4. 若 blnNewFI，AddNew rsFIBak 并 Update
'   5. Select Case ItemsData.Fields("HSTag").Value (由调用方负责)
'
' 本 helper 只做 1, 2, 4。调用方负责 3 和 5。
'
' 跨工程懒加载：本 helper 内部自动 EnsureLoaded VouMetaCache，调用方不用
' 关心 IsLoaded。如 LoadAll 失败（DB 异常）会向上抛错。
'==============================================================================
Public Sub CheckFI(ByVal objDS As HHDataService.sysDataService, _
                   ByVal ItemsData As ADODB.Recordset, _
                   ByRef rsfcode As ADODB.Recordset, _
                   ByRef rsFIBak As ADODB.Recordset, _
                   ByVal ActCHName As String, _
                   ByVal strBillInfo As String, _
                   ByRef strErrInfo As String)
    Call VouMetaCache.EnsureLoaded(objDS)
    Dim blnUseCache As Boolean
    blnUseCache = VouMetaCache.IsLoaded
    Dim sFIID As String
    sFIID = objDS.NullToStr(ItemsData.Fields("FIID").Value)

    ' 与原代码一致：FIID 为空时不在此 helper 处理（外层已处理）
    If sFIID = "" Then Exit Sub

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
        '--- Step 1: 找 FIID 行（cache 或 db）------------------------------
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
                lngHSTag = CLng(objDS.NullToDbl(rsfcode.Fields("hstag").Value))
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

        '--- Step 2: 校验 + 写 ItemsData 字段 -----------------------------
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

        '--- Step 4: rsFIBak 跨行缓存（原代码就在这里写，而且是在 ItemsData.Update 之后） --
        ' 注意：调用方应在调用本 helper 后立刻 ItemsData.Update。
        '       本 helper 内写 rsFIBak 的时机与原代码一致 —— 在 ItemsData 字段
        '       写入之后即可，不依赖 ItemsData.Update 是否已发生（rsFIBak 是
        '       完全独立的内存 Recordset）。
        If rsFIBak Is Nothing Then
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
        '--- 命中 rsFIBak：只写 ItemsData，不再校验（与原代码一致）---------
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
'
' ⚠️ 原代码已知 bug：
'   原 BeforeAction 错误信息引用 rsAccs.Fields("CorpName")，但 Account
'   表 SELECT 只查了 AccID/ISStop/AccName 三列，不含 CorpName 字段。
'   访问会抛 ADO "Item cannot be found in this collection" 异常。生产
'   极少触发（账户被删/停用且仍在用），所以长期未暴露。
'
'   100% 等价处置：保留 db 路径原 bug 行为；cache 路径下若**未命中**
'   或**命中但 isStop=True**，fallback 到 db 路径（必须传入 rsAccs），
'   触发同样的 ADO 异常 —— 生产中行为完全一致。
'==============================================================================
Public Sub CheckAcc(ByVal objDS As HHDataService.sysDataService, _
                    ByVal ItemsData As ADODB.Recordset, _
                    ByRef rsAccs As ADODB.Recordset, _
                    ByVal ActCHName As String, _
                    ByVal strBillInfo As String, _
                    ByRef strErrInfo As String)
    Dim sAccID As String
    sAccID = objDS.NullToStr(ItemsData.Fields("AccID").Value)
    If sAccID = "" Then Exit Sub

    Call VouMetaCache.EnsureLoaded(objDS)
    Dim blnUseCache As Boolean
    blnUseCache = VouMetaCache.IsLoaded

    Dim blnFound As Boolean
    Dim blnIsStop As Boolean
    Dim sAccName As String     ' 仅 cache 路径用于 debug 上下文，不影响最终错误信息

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

    ' 错误分支：必须严格等价于原代码 —— 用 rsAccs.Fields("CorpName") 触发
    ' 原 ADO 异常。无论 cache/db 路径都走 rsAccs（cache 路径下 rsAccs
    ' 仍可能为 Nothing，此时 raise 同型异常以保持等价）。
    If Not blnFound Then
        Call meRaiseAccErrLikeOriginal(rsAccs, ItemsData, ActCHName, strBillInfo, "找不到", strErrInfo)
        Exit Sub
    End If
    If blnIsStop Then
        Call meRaiseAccErrLikeOriginal(rsAccs, ItemsData, ActCHName, strBillInfo, "被停用", strErrInfo)
    End If
End Sub


'==============================================================================
' Private：复刻原 BeforeAction 中 rsAccs.Fields("CorpName") 的 bug 行为
'   原代码无论账户是否找到/停用，都直接拼接 rsAccs.Fields("CorpName")
'   而该字段不在 Account 表查询中 —— 任何访问都会抛 ADO 异常。
'==============================================================================
Private Sub meRaiseAccErrLikeOriginal(ByVal rsAccs As ADODB.Recordset, _
                                      ByVal ItemsData As ADODB.Recordset, _
                                      ByVal ActCHName As String, _
                                      ByVal strBillInfo As String, _
                                      ByVal strReason As String, _
                                      ByRef strErrInfo As String)
    Dim sCorpName As String
    If rsAccs Is Nothing Then
        ' cache 路径下 rsAccs 不存在 —— 直接抛 ADO 等价异常以保持原 bug 行为
        Call Err.Raise(3265, "VouCacheHelpers.CheckAcc", _
            "Item cannot be found in the collection corresponding to the requested name or ordinal.")
    End If
    ' db 路径：访问 rsAccs.Fields("CorpName") 会抛 3265
    sCorpName = rsAccs.Fields("CorpName").Value
    ' 几乎到不了这一行（Fields("CorpName") 不存在）
    strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
        " 条分录[" & ItemsData.Fields("FINO").Value & "]中的帐户[" & sCorpName & _
        "]已经" & strReason & "，无法" & ActCHName & "！"
End Sub


'==============================================================================
' Public：Corp 检查
'
' ⚠️ 原代码已知 bug：
'   原代码在 rsCorps.RecordCount=0 时访问 rsCorps.Fields("CorpName").Value，
'   ADO Filter 后无当前行会抛 BOF/EOF 异常 (ADO error 3021)。
'
'   为保持 100% 等价：cache 路径下若未命中，抛同型 ADO 异常（用户在生产
'   触发时已习惯于看到该错误）。
'==============================================================================
Public Sub CheckCorp(ByVal objDS As HHDataService.sysDataService, _
                     ByVal ItemsData As ADODB.Recordset, _
                     ByRef rsCorps As ADODB.Recordset, _
                     ByVal ActCHName As String, _
                     ByVal strBillInfo As String, _
                     ByRef strErrInfo As String)
    Dim sCorpID As String
    sCorpID = objDS.NullToStr(ItemsData.Fields("CorpID").Value)
    If sCorpID = "" Then Exit Sub

    Call VouMetaCache.EnsureLoaded(objDS)
    Dim blnUseCache As Boolean
    blnUseCache = VouMetaCache.IsLoaded

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
        ' 等价于原代码 BOF/EOF 访问 Fields("CorpName").Value 抛 3021
        If blnUseCache Then
            Call Err.Raise(3021, "VouCacheHelpers.CheckCorp", _
                "Either BOF or EOF is True, or the current record has been deleted; the operation requested by the application requires a current record.")
        Else
            ' db 路径：原代码原地拼字符串触发同样异常
            strErrInfo = strErrInfo & strBillInfo & "第 " & ItemsData.Fields("itmid").Value & _
                " 条分录[" & ItemsData.Fields("FINO").Value & "]中的" & sCTagName & "[" & _
                rsCorps.Fields("CorpName").Value & _
                "]已经找不到，无法" & ActCHName & "！" & vbCrLf
        End If
        Exit Sub
    End If
    If Not blnContact Then
        ' 找到但被停用：rsCorps 当前行有效，CorpName 可正常读
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
                    ByRef rsEmps As ADODB.Recordset, _
                    ByVal ActCHName As String, _
                    ByVal strBillInfo As String, _
                    ByRef strErrInfo As String)
    Dim sEmpID As String
    sEmpID = objDS.NullToStr(ItemsData.Fields("empid").Value)
    If sEmpID = "" Then Exit Sub

    Call VouMetaCache.EnsureLoaded(objDS)
    Dim blnUseCache As Boolean
    blnUseCache = VouMetaCache.IsLoaded

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


'##############################################################################
'#
'# 组 B：t_FVou_M.GetFIIDBy* 系列函数的 cache 路径统一封装
'#
'# 原 t_FVou_M 中有 16 个 GetFIIDBy* 函数，每个都依赖 mrsDocFIRel（实例级
'# Recordset 缓存）。PR-1 引入 VouMetaCache 后这些函数都可以在 cache 已
'# 加载时跳过 mrsDocFIRel，本组 helper 把所有 cache 路径集中实现。
'#
'# 调用方法（典型用法）：
'#   Public Function GetFIIDByPrdt(...) As String
'#       If VouCacheHelpers.LookupDocFI_Strict(objDS, "PrdtCls", "存货", ...) Then
'#           GetFIIDByPrdt = ...   ' helper 已通过 ByRef 输出 FIID/FINO/FIName
'#           Exit Function
'#       End If
'#       ' --- cache 未加载 / 未命中 → 原 mrsDocFIRel 路径不变 ---
'#       ...
'#   End Function
'#
'##############################################################################

'==============================================================================
' Public：5 步 fallback 查找（精确 → WHID-only → ClsID-only → ParentCls → 全空 → Raise）
'   覆盖原代码 GetFIIDByPrdt / GetFIIDByPrdt1 / GetKCFIID / GetIRILFIID /
'                GetIJFIID / GetSSATMFIID / GetSSCostFIID / GetMFCostFIID
'   返回 True 表示 cache 已加载并命中（rtnFIID/rtnFINO/rtnFIName 已通过 ByRef 输出）
'   返回 False 表示 cache 未加载，调用方应回退到原 mrsDocFIRel 路径
'   isStrict=True 时找不到抛错（GetFIIDByPrdt 行为）；False 时静默返回 False（GetFIIDByPrdt1 行为）
'==============================================================================
Public Function LookupDocFI_5Step(ByVal objDS As HHDataService.sysDataService, _
                                  ByVal DocKey As String, ByVal DocCHName As String, _
                                  ByVal WHID As String, ByVal WHName As String, _
                                  ByVal ClsID As String, ByVal ClsName As String, _
                                  ByVal isStrict As Boolean, _
                                  ByRef rtnFIID As String, _
                                  ByRef rtnFINO As String, ByRef rtnFIName As String) As Boolean
    Call VouMetaCache.EnsureLoaded(objDS)
    If Not VouMetaCache.IsLoaded Then Exit Function

    Dim sFIID As String, sFINO As String, sFIName As String
    Dim sHSTagName As String
    Dim lngHSTag As Long, lngFITag As Long

    ' Step 1: DocKey + WHID + ClsID 精确
    If VouMetaCache.TryGetDocFI(DocKey, WHID, ClsID, sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
        rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
        LookupDocFI_5Step = True
        Exit Function
    End If
    ' Step 2: DocKey + WHID + DocID=''
    If VouMetaCache.TryGetDocFI(DocKey, WHID, "", sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
        rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
        LookupDocFI_5Step = True
        Exit Function
    End If
    ' Step 3: DocKey + WHID='' + DocID=ClsID
    If VouMetaCache.TryGetDocFI(DocKey, "", ClsID, sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
        rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
        LookupDocFI_5Step = True
        Exit Function
    End If
    ' Step 4: 上级品类（必须查 DB getPrdtClsTree；此处 cache 路径主动调原函数 fallback）
    '          ⚠️ 调用方负责实现"上级品类查询" 走 cache 而不是 mrsDocFIRel —— 见
    '          LookupDocFI_ParentCls
    Dim sFIIDFromParent As String, sFINOFromParent As String, sFINameFromParent As String
    If LookupDocFI_ParentCls(objDS, DocKey, WHID, ClsID, _
                             sFIIDFromParent, sFINOFromParent, sFINameFromParent) Then
        rtnFIID = sFIIDFromParent: rtnFINO = sFINOFromParent: rtnFIName = sFINameFromParent
        LookupDocFI_5Step = True
        Exit Function
    End If
    ' Step 5: 全空
    If VouMetaCache.TryGetDocFI(DocKey, "", "", sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
        rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
        LookupDocFI_5Step = True
        Exit Function
    End If

    ' 全部未命中
    If isStrict Then
        Call Err.Raise(ERROR_FORSYSTEM, , WHName & " 的 " & ClsName & _
            " 尚未设置" & DocCHName & "会计科目，无法生成记账凭证！")
    End If
    ' isStrict=False：返回 True 表示"cache 路径已处理但未找到"
    ' （不抛错，rtnFIID 保持空，调用方判断空字符串即可）
    rtnFIID = "": rtnFINO = "": rtnFIName = ""
    LookupDocFI_5Step = True
End Function


'==============================================================================
' Public：上级品类 fallback（替代 t_FVou_M.getFIIDByClsIDFromParentCls 的 cache 路径）
'   原函数依赖 mrsDocFIRel + mrsWHClsFI + getPrdtClsTree 函数；
'   cache 路径下 mrsDocFIRel 是 Nothing，必须改用 VouMetaCache.TryGetDocFI 替代 Filter；
'   getPrdtClsTree 是 SQL Server function，仍需查 DB（不可避免）。
'   返回 True = 找到（rtnFIID 已设）；False = 未找到（继续 step5）
'==============================================================================
Public Function LookupDocFI_ParentCls(ByVal objDS As HHDataService.sysDataService, _
                                      ByVal DocKey As String, _
                                      ByVal WHID As String, ByVal ClsID As String, _
                                      ByRef rtnFIID As String, _
                                      ByRef rtnFINO As String, _
                                      ByRef rtnFIName As String) As Boolean
On Error GoTo ErrH
    Dim rsCls As ADODB.Recordset
    Dim sParentClsID As String
    Dim sFIID As String, sFINO As String, sFIName As String
    Dim sHSTagName As String
    Dim lngHSTag As Long, lngFITag As Long

    ' getPrdtClsTree 是 SQL function，必须查 DB
    Set rsCls = objDS.OpenRecordsetBySQL( _
        "SELECT ClsID FROM getPrdtClsTree('" & ClsID & "',1,-1,0) a " & _
        "WHERE clsid<>'' ORDER BY TreeLevel", True, True)

    Do While Not rsCls.EOF
        sParentClsID = CStr(rsCls.Fields("ClsID").Value)
        ' 与原代码 2 步 Filter 等价：
        '   1) DocKey + WHID + ParentClsID
        '   2) DocKey + WHID='' + ParentClsID
        If VouMetaCache.TryGetDocFI(DocKey, WHID, sParentClsID, _
                                    sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
            rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
            LookupDocFI_ParentCls = True
            Call objDS.rs_Close(rsCls)
            Exit Function
        End If
        If VouMetaCache.TryGetDocFI(DocKey, "", sParentClsID, _
                                    sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
            rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
            LookupDocFI_ParentCls = True
            Call objDS.rs_Close(rsCls)
            Exit Function
        End If
        rsCls.MoveNext
    Loop

ErrH:
    On Error Resume Next
    Call objDS.rs_Close(rsCls)
    On Error GoTo 0
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, , Err.Description)
End Function


'==============================================================================
' Public：DocKey + DocID 两步 fallback（DocID 精确 → DocID='' → 抛错或返回空）
'   覆盖原代码 GetFIIDByKey / GetAccFIID / GetCGExpYFZKFIID / GetFactYJTax_XXFIID /
'             GetFactYJTax_JXFIID / GetYJTax_TInvFIID / GetYJTax_XXFIID / GetYJTax_JXFIID
'   raiseOnNotFound=True 时找不到抛错（GetAccFIID 等行为）；False 时返回空（NotExistsRaiseErr=False）
'   raiseMsg 是抛错时的提示内容（由调用方拼好）
'==============================================================================
Public Function LookupDocFI_KeyOnly(ByVal objDS As HHDataService.sysDataService, _
                                    ByVal DocKey As String, ByVal DocCHName As String, _
                                    ByVal DocID As String, ByVal DocName As String, _
                                    ByVal raiseOnNotFound As Boolean, _
                                    ByVal raiseMsg As String, _
                                    ByRef rtnFIID As String, _
                                    ByRef rtnFINO As String, _
                                    ByRef rtnFIName As String) As Boolean
    Call VouMetaCache.EnsureLoaded(objDS)
    If Not VouMetaCache.IsLoaded Then Exit Function

    Dim sFIID As String, sFINO As String, sFIName As String
    Dim sHSTagName As String
    Dim lngHSTag As Long, lngFITag As Long

    ' Step 1: DocKey + DocID 精确
    If VouMetaCache.TryGetDocFI(DocKey, "", DocID, sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
        rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
        LookupDocFI_KeyOnly = True
        Exit Function
    End If
    ' Step 2: DocKey + DocID=''
    If VouMetaCache.TryGetDocFI(DocKey, "", "", sFIID, sFINO, sFIName, lngHSTag, sHSTagName, lngFITag) Then
        rtnFIID = sFIID: rtnFINO = sFINO: rtnFIName = sFIName
        LookupDocFI_KeyOnly = True
        Exit Function
    End If

    If raiseOnNotFound Then
        Call Err.Raise(ERROR_FORSYSTEM, , raiseMsg)
    End If
    rtnFIID = "": rtnFINO = "": rtnFIName = ""
    LookupDocFI_KeyOnly = True
End Function
