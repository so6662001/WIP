# PR-1 优化补丁应用清单

本 PR 把 6 处 N+1 查询替换为 `VouMetaCache.bas` 的内存查找，并把 `meAddRowI` 反射赋值展开为显式属性赋值。**所有修改保持调用接口和返回值 100% 等价**。

应用顺序：
1. 把 `VouMetaCache.bas` 直接添加到 VB6 工程（标准模块）
2. 按下面的 6 处 patch 修改对应的现有 `.cls` 文件

---

## Patch 1：`cMthCstAccGL2.cls` — 在 `meCreateVou` 入口加载 / 末尾清理

### 修改前

```vb
Public Sub meCreateVou(ByVal ComID As String, ByVal BeginDate As Date, ByVal EndDate As String, ByVal objDS As sysDataService, ByVal APID As String, _
                        ByVal isPreGenIDOnly As Boolean, ByVal isSaveToTransitionalTable As Boolean)
On Error GoTo ErrH
    Dim blnGL2  As Boolean
    Dim SQL     As String
    
    If EndDate = "" Then
        EndDate = Format(Date, "yyyy-MM-dd")
    End If
    
    If Not (Me.AppParameters.UseGL2 And Me.ReCreateGL) Then
        GoTo ErrH
    End If
    If objGL2 Is Nothing Then
        Set objGL2 = New cMthCstAccGL2
        ...
    End If
    objGL2.isPreGenIDOnly = isPreGenIDOnly
    objGL2.isSaveToTransitionalTable = isSaveToTransitionalTable
    
    '生成会计凭证 只需要处理成本 2017-11-09
    If Me.AppParameters.PSBillAutoCreateVou Then
        Call objGL2.meCreVouForSS(objDS)
    End If
    ...
    Call objGL2.CreateCstDiffVou(objDS, mEmpID)
    
ErrH:
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, , Err.Description)
    End If
End Sub
```

### 修改后

```vb
Public Sub meCreateVou(ByVal ComID As String, ByVal BeginDate As Date, ByVal EndDate As String, ByVal objDS As sysDataService, ByVal APID As String, _
                        ByVal isPreGenIDOnly As Boolean, ByVal isSaveToTransitionalTable As Boolean)
On Error GoTo ErrH
    Dim blnGL2  As Boolean
    Dim SQL     As String
    Dim blnMetaLoadedHere As Boolean        ' === PR-1 新增 ===
    
    If EndDate = "" Then
        EndDate = Format(Date, "yyyy-MM-dd")
    End If
    
    If Not (Me.AppParameters.UseGL2 And Me.ReCreateGL) Then
        GoTo ErrH
    End If
    If objGL2 Is Nothing Then
        Set objGL2 = New cMthCstAccGL2
        ...
    End If
    objGL2.isPreGenIDOnly = isPreGenIDOnly
    objGL2.isSaveToTransitionalTable = isSaveToTransitionalTable
    
    ' === PR-1 新增：整 batch 元数据预加载 ===
    If Not VouMetaCache.IsLoaded Then
        Call VouMetaCache.LoadAll(objDS)
        blnMetaLoadedHere = True            ' 标记由本次入口加载，退出时由本次清理
    End If
    
    '生成会计凭证 只需要处理成本 2017-11-09
    If Me.AppParameters.PSBillAutoCreateVou Then
        Call objGL2.meCreVouForSS(objDS)
    End If
    ...
    Call objGL2.CreateCstDiffVou(objDS, mEmpID)
    
ErrH:
    ' === PR-1 新增：仅当本入口加载时才负责清理 ===
    If blnMetaLoadedHere Then
        Call VouMetaCache.ClearAll
    End If
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, , Err.Description)
    End If
End Sub
```

---

## Patch 2：`t_FVou_M.cls` — `Init()` 优先用缓存

### 修改前

```vb
Private Sub Init(ByVal objDS As HHDataService.sysDataService, ByVal DocKey As String)
    Dim SQL As String
    
    SQL = "SELECT isnull(dockey,'') as dockey,  ISNULL(D.DocID,'') AS DocID, ISNULL(D.WHID,'') AS WHID, D.FIID, FI.FINo, CASE WHEN ISNULL(FI.FullName,'')='' THEN FI.FINAME ELSE FI.FullName END as FIName, FI.HSTag, FI.HSTagName, FI.FITag " & vbCrLf
    SQL = SQL & "FROM dbo.GL2_AchFIID D WITH(NOLOCK) INNER JOIN" & vbCrLf
    SQL = SQL & "   dbo.FinanceItems FI ON D.FIID = FI.FIID" & vbCrLf
    SQL = SQL & "WHERE ISNULL(D.FIID,'')<>''"
    
    Set mrsDocFIRel = objDS.OpenRecordsetBySQL(SQL, True, True)
End Sub
```

### 修改后（保持原行为；当 VouMetaCache 已加载时跳过 DB 查询）

```vb
Private Sub Init(ByVal objDS As HHDataService.sysDataService, ByVal DocKey As String)
    Dim SQL As String
    
    ' === PR-1 修改：如果 VouMetaCache 已加载，则不再读 DB
    '     mrsDocFIRel 不创建，所有 GetFIIDByPrdt / GetFIIDByPrdt1 / GetAccFIID
    '     等下游都改为先调 VouMetaCache.TryGetDocFI ===
    If VouMetaCache.IsLoaded Then
        ' mrsDocFIRel 保持 Nothing，使用模块缓存
        Exit Sub
    End If
    
    SQL = "SELECT isnull(dockey,'') as dockey,  ISNULL(D.DocID,'') AS DocID, ISNULL(D.WHID,'') AS WHID, D.FIID, FI.FINo, CASE WHEN ISNULL(FI.FullName,'')='' THEN FI.FINAME ELSE FI.FullName END as FIName, FI.HSTag, FI.HSTagName, FI.FITag " & vbCrLf
    SQL = SQL & "FROM dbo.GL2_AchFIID D WITH(NOLOCK) INNER JOIN" & vbCrLf
    SQL = SQL & "   dbo.FinanceItems FI ON D.FIID = FI.FIID" & vbCrLf
    SQL = SQL & "WHERE ISNULL(D.FIID,'')<>''"
    
    Set mrsDocFIRel = objDS.OpenRecordsetBySQL(SQL, True, True)
End Sub
```

---

## Patch 3：`t_FVou_M.cls` — 16 个 GetFIIDBy* 函数 cache 路径

每个函数在原代码**第一行**加 cache 路径调用 helper。helper 已在 `VouCacheHelpers.bas` 组 B 中实现完整 5 步 / 2 步 fallback 逻辑（含 ParentCls 走 cache）。

### Patch 3-1：`GetFIIDByPrdt` (5 步严格)

```vb
Public Function GetFIIDByPrdt(ByVal objDS As HHDataService.sysDataService, _
                              ByVal DocKey As String, ByVal DocCHName As String, _
                              ByVal WHID As String, ByVal WHName As String, _
                              ByVal ClsID As String, ByVal ClsName As String, _
                              ByRef rtnFINO As String, ByRef rtnFIName As String) As String

    ' === PR-1 cache fast path ===
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_5Step(objDS, DocKey, DocCHName, WHID, WHName, _
                                         ClsID, ClsName, True, _
                                         sFIID, rtnFINO, rtnFIName) Then
        GetFIIDByPrdt = sFIID
        Exit Function
    End If

    ' === 原路径（cache 未加载）保持不变 ===
    If mrsDocFIRel Is Nothing Then
        Call Init(objDS, DocKey)
    End If
    ' ...原代码全部保留...
End Function
```

### Patch 3-2：`GetFIIDByPrdt1` (4 步非严格)

```vb
Public Function GetFIIDByPrdt1(ByVal objDS As HHDataService.sysDataService, _
                               ByVal DocKey As String, ByVal DocCHName As String, _
                               ByVal WHID As String, ByVal WHName As String, _
                               ByVal ClsID As String, ByVal ClsName As String, _
                               ByRef rtnFINO As String, ByRef rtnFIName As String) As String

    ' === PR-1 cache fast path（isStrict=False，找不到不抛错）===
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_5Step(objDS, DocKey, DocCHName, WHID, WHName, _
                                         ClsID, ClsName, False, _
                                         sFIID, rtnFINO, rtnFIName) Then
        GetFIIDByPrdt1 = sFIID
        Exit Function
    End If

    ' === 原路径保持不变 ===
End Function
```

### Patch 3-3：`getFIIDByClsIDFromParentCls`（被 5 步调用，必须改）

```vb
Private Function getFIIDByClsIDFromParentCls(ByVal objDS As HHDataService.sysDataService, _
                                             ByVal DocKey As String, ByVal WHID As String, _
                                             ByVal ClsID As String, _
                                             ByRef rtnFINO As String, ByRef rtnFIName As String) As String
    ' === PR-1 cache fast path ===
    If VouMetaCache.IsLoaded Then
        Dim sFIID As String
        If VouCacheHelpers.LookupDocFI_ParentCls(objDS, DocKey, WHID, ClsID, _
                                                 sFIID, rtnFINO, rtnFIName) Then
            getFIIDByClsIDFromParentCls = sFIID
        End If
        Exit Function
    End If

    ' === 原路径保持不变 ===
End Function
```

### Patch 3-4：其它包装函数无需修改

`GetKCFIID` / `GetIRILFIID` / `GetIJFIID` / `GetSSATMFIID` / `GetSSCostFIID` / `GetMFCostFIID` 都是简单包装，内部调用 `GetFIIDByPrdt` / `GetFIIDByPrdt1`，**自动受益于 Patch 3-1/3-2**，无需单独修改。

### Patch 3-5：`GetAccFIID` (DocKey+DocID 两步严格)

```vb
Public Function GetAccFIID(ByVal objDS As HHDataService.sysDataService, _
                           ByVal AccID As String, ByVal AccName As String, _
                           ByRef rtnFINO As String, ByRef rtnFIName As String) As String
    ' === PR-1 cache fast path ===
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_KeyOnly(objDS, "Account", "货币资金", _
                                           AccID, AccName, _
                                           True, AccName & " 尚未设置货币资金会计科目，无法生成记账凭证！", _
                                           sFIID, rtnFINO, rtnFIName) Then
        GetAccFIID = sFIID
        Exit Function
    End If

    ' === 原路径保持不变 ===
End Function
```

### Patch 3-6：`GetCGExpYFZKFIID`

```vb
Public Function GetCGExpYFZKFIID(...) As String
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_KeyOnly(objDS, "CGExpFI", "采购费用", _
                                           CGExpID, CGExpName, _
                                           True, CGExpName & " 尚未设置采购费用会计科目，无法生成记账凭证！", _
                                           sFIID, rtnFINO, rtnFIName) Then
        GetCGExpYFZKFIID = sFIID
        Exit Function
    End If
    ' === 原路径保持不变 ===
End Function
```

### Patch 3-7 ~ 3-10：`GetFactYJTax_XXFIID` / `GetFactYJTax_JXFIID` / `GetYJTax_XXFIID` / `GetYJTax_JXFIID`

这 4 个函数内部都需要先判 `Me.AppParameters.YJTaxFIIDByTInvID = False` 直接返回常量，再判 `TInvID = ""`返回空，**这两个早期 return 必须先判，再走 cache fast path**：

```vb
Public Function GetFactYJTax_XXFIID(...) As String
    If Me.AppParameters.YJTaxFIIDByTInvID = False Then
        GetFactYJTax_XXFIID = Me.AppParameters.GL2_YJTaxXFIID_F
        Exit Function
    End If
    If TInvID = "" Then
        rtnFINO = "": rtnFIName = "": GetFactYJTax_XXFIID = ""
        Exit Function
    End If

    ' === PR-1 cache fast path ===
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_KeyOnly(objDS, "TInvXXFact", "应交税费销项", _
                                           TInvID, TInvName, _
                                           True, TInvName & " 尚未设置应交税费销项会计科目，无法生成记账凭证！", _
                                           sFIID, rtnFINO, rtnFIName) Then
        GetFactYJTax_XXFIID = sFIID
        Exit Function
    End If

    ' === 原路径保持不变 ===
End Function
```

> 其它三个 `GetFactYJTax_JXFIID` / `GetYJTax_XXFIID` / `GetYJTax_JXFIID` 同模式，DocKey 分别为 `"TInvJXFact"` / `"TInvXX"` / `"TInvJX"`，错误文案见原代码。

### Patch 3-11：`GetYJTax_TInvFIID`

```vb
Public Function GetYJTax_TInvFIID(...) As String
    If TInvID = "" Then
        rtnFINO = "": rtnFIName = "": GetYJTax_TInvFIID = ""
        Exit Function
    End If

    ' === PR-1 cache fast path ===
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_KeyOnly(objDS, DocKey, DocCHName, _
                                           TInvID, TInvName, _
                                           True, TInvName & " 尚未设置" & _
                                           IIf(Me.AppParameters.GL2_UseJITIYJTax, "计提", "") & _
                                           "应交税费" & DocCHName & "，无法生成记账凭证！", _
                                           sFIID, rtnFINO, rtnFIName) Then
        GetYJTax_TInvFIID = sFIID
        Exit Function
    End If

    ' === 原路径保持不变 ===
End Function
```

### Patch 3-12：`GetFIIDByKey`

```vb
Public Function GetFIIDByKey(..., Optional ByVal NotExistsRaiseErr As Boolean = True) As String
    ' === PR-1 cache fast path ===
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_KeyOnly(objDS, DocKey, DocCHName, _
                                           DocID, DocName, _
                                           NotExistsRaiseErr, _
                                           DocName & " 尚未设置" & DocCHName & "会计科目，无法生成记账凭证！", _
                                           sFIID, rtnFINO, rtnFIName) Then
        GetFIIDByKey = sFIID
        Exit Function
    End If

    ' === 原路径保持不变 ===
End Function
```

### 总结：Patch 3 共影响 **12 处**

| 函数 | 走的 helper |
|---|---|
| `GetFIIDByPrdt` | `LookupDocFI_5Step` (isStrict=True) |
| `GetFIIDByPrdt1` | `LookupDocFI_5Step` (isStrict=False) |
| `getFIIDByClsIDFromParentCls` | `LookupDocFI_ParentCls` |
| `GetKCFIID` / `GetIRILFIID` / `GetIJFIID` / `GetSSATMFIID` / `GetSSCostFIID` / `GetMFCostFIID` | 包装函数，自动继承 |
| `GetAccFIID` | `LookupDocFI_KeyOnly` |
| `GetCGExpYFZKFIID` | `LookupDocFI_KeyOnly` |
| `GetFactYJTax_XXFIID` / `GetFactYJTax_JXFIID` / `GetYJTax_XXFIID` / `GetYJTax_JXFIID` / `GetYJTax_TInvFIID` | `LookupDocFI_KeyOnly` (5 处) |
| `GetFIIDByKey` | `LookupDocFI_KeyOnly` |

> 任何 helper 返回 `False` 时（cache 未加载），调用方必须**保留原 mrsDocFIRel 路径完全不动**。这保证 cache 未启用时行为零变化。

---

## Patch 4：`CVouService.cls` — `BeforeAction` 中 4 张元数据表查询替换

### 修改前

```vb
Private Sub BeforeAction(...)
    ...
    If strFIIDList <> "" Then
        strFIIDList = Left(strFIIDList, Len(strFIIDList) - 1)
        SQL = "SELECT FIID,FIName,FINO,ISStop,FITag,ExpTag,HSTag,HSTagName,baNum,baCust,baSupp,baOtherCorp,fullname FROM FinanceItems WHERE FIID in (" & strFIIDList & ")"
        Set rsfcode = objDS.OpenRecordsetBySQL(SQL, True, True)
    End If
    If strCorpIDList <> "" Then
        strCorpIDList = Left(strCorpIDList, Len(strCorpIDList) - 1)
        SQL = "SELECT Corpid,Contact,CorpName FROM Corp WHERE Corpid in (" & strCorpIDList & ")"
        Set rsCorps = objDS.OpenRecordsetBySQL(SQL, True, True)
    End If
    If strEmpIDList <> "" Then
        strEmpIDList = Left(strEmpIDList, Len(strEmpIDList) - 1)
        SQL = "SELECT empid, dismission,EmpName FROM Emp WHERE EmpID IN (" & strEmpIDList & ")"
        Set rsEmps = objDS.OpenRecordsetBySQL(SQL, True, True)
    End If
    If strAccIDList <> "" Then
        strAccIDList = Left(strAccIDList, Len(strAccIDList) - 1)
        SQL = "SELECT AccID,ISStop,AccName FROM Account WHERE AccID in (" & strAccIDList & ")"
        Set rsAccs = objDS.OpenRecordsetBySQL(SQL, True, True)
    End If
    
    Call objDS.rs_MoveFirst(ItemsData)
    Do While Not ItemsData.EOF
        ' ... 后续用 rsfcode / rsCorps / rsEmps / rsAccs 的 Filter
        rsfcode.Filter = "fiid='" & .Fields("fiid").Value & "'"
        ...
    Loop
End Sub
```

### 修改后

新增私有辅助过程 `meBeforeAction_FillRow`，并在 `BeforeAction` 入口判断缓存：

```vb
Private Sub BeforeAction(...)
    ...
    Dim blnUseCache As Boolean
    blnUseCache = VouMetaCache.IsLoaded            ' === PR-1 新增 ===
    
    If Not blnUseCache Then
        ' 原 4 段 SQL 查询保持不变
        If strFIIDList <> "" Then
            ...
            Set rsfcode = objDS.OpenRecordsetBySQL(...)
        End If
        ' Corp / Emp / Acc 同样保留
    End If
    
    Call objDS.rs_MoveFirst(ItemsData)
    Do While Not ItemsData.EOF
        ' ... 字段访问统一改为通过 helper：
        Call meCheckFIID(objDS, blnUseCache, rsfcode, rsFIBak, ItemsData, ActCHName, strBillInfo, strErrInfo)
        Call meCheckAccID(objDS, blnUseCache, rsAccs, ItemsData, ActCHName, strBillInfo, strErrInfo)
        Call meCheckCorp(objDS, blnUseCache, rsCorps, ItemsData, ActCHName, strBillInfo, strErrInfo)
        Call meCheckEmp(objDS, blnUseCache, rsEmps, ItemsData, ActCHName, strBillInfo, strErrInfo)
        ItemsData.MoveNext
    Loop
End Sub
```

> 完整 helper 实现见 `CVouService_PR1_helpers.bas`。helper 的两条路径（cache 与非 cache）必须返回完全相同的 ItemsData 字段写入和 strErrInfo 内容。

---

## Patch 5：`CVouService.cls` — `CheckDateValidate` 走缓存

### 修改前

```vb
Public Function CheckDateValidate(ByVal vDate As Date, Optional ByVal vCnn As ADODB.Connection = Nothing) As String
    ...
    Set rsTmp = objDS.OpenRecordsetBySQL("SELECT APID," & strAPTagFld & " as APTag FROM AccPeriod WHERE BeginDate<='" & ... )
    If rsTmp.RecordCount = 0 Then
        Call Err.Raise(ERROR_FORSYSTEM, , "无法继续，因为您录入的日期不在本年度会计期间内！")
    End If
    Select Case CInt(objDS.NullToDbl(rsTmp.Fields("APTag").Value))
        Case 1
            Call Err.Raise(...)
        Case -1
            Call Err.Raise(...)
    End Select
    strPID = rsTmp.Fields("APID").Value
    Call objDS.rs_Close(rsTmp)

    Set rsTmp = objDS.OpenRecordsetBySQL("SELECT EndDate FROM AccPeriod WHERE " & strAPTagFld & "=2 ", True, True)
    If rsTmp.RecordCount = 0 Then
        Call Err.Raise(ERROR_FORSYSTEM, , "会计期间被破坏，系统无法继续！")
    End If
    CheckDateValidate = strPID
    ...
End Function
```

### 修改后

```vb
Public Function CheckDateValidate(ByVal vDate As Date, Optional ByVal vCnn As ADODB.Connection = Nothing) As String
On Error GoTo ErrH
    Dim rsTmp   As ADODB.Recordset
    Dim objDS   As sysDataService
    Dim strPID  As String
    Dim strAPTagFld As String

    strAPTagFld = IIf(Me.AppParameters.GL2_ERPFODiffCarry, "FOAPTag", "APTag")
    
    ' === PR-1 新增：优先走缓存 ===
    If VouMetaCache.IsLoaded Then
        Dim sAPID As String
        Dim lngAPTag As Long
        If Not VouMetaCache.TryGetAPByDate(vDate, Me.AppParameters.GL2_ERPFODiffCarry, sAPID, lngAPTag) Then
            Call Err.Raise(ERROR_FORSYSTEM, , "无法继续，因为您录入的日期不在本年度会计期间内！")
        End If
        Select Case CInt(lngAPTag)
            Case 1
                Call Err.Raise(ERROR_FORSYSTEM, , "您不能录入已月结会计期间的凭证！")
            Case -1
                Call Err.Raise(ERROR_FORSYSTEM, , "您不能录入以前会计期间的凭证！")
        End Select
        If Not VouMetaCache.HasActiveAP(Me.AppParameters.GL2_ERPFODiffCarry) Then
            Call Err.Raise(ERROR_FORSYSTEM, , "会计期间被破坏，系统无法继续！")
        End If
        CheckDateValidate = sAPID
        Exit Function
    End If

    ' === 原路径 ===
    Set objDS = New sysDataService
    objDS.ConnectionString = Me.ConnectionString
    If Not vCnn Is Nothing Then
        Set objDS.Connection = vCnn
    Else
        Call objDS.OpenConnection
    End If
    Set rsTmp = objDS.OpenRecordsetBySQL("SELECT APID," & strAPTagFld & " as APTag FROM AccPeriod WHERE BeginDate<='" & Format(vDate, "yyyy-MM-dd") & "' AND EndDate>='" & Format(vDate, "yyyy-MM-dd") & "'", True, True)
    If rsTmp.RecordCount = 0 Then
        Call Err.Raise(ERROR_FORSYSTEM, , "无法继续，因为您录入的日期不在本年度会计期间内！")
    End If
    Select Case CInt(objDS.NullToDbl(rsTmp.Fields("APTag").Value))
        Case 1
            Call Err.Raise(ERROR_FORSYSTEM, , "您不能录入已月结会计期间的凭证！")
        Case -1
            Call Err.Raise(ERROR_FORSYSTEM, , "您不能录入以前会计期间的凭证！")
    End Select
    strPID = rsTmp.Fields("APID").Value
    Call objDS.rs_Close(rsTmp)
    Set rsTmp = objDS.OpenRecordsetBySQL("SELECT EndDate FROM AccPeriod WHERE " & strAPTagFld & "=2 ", True, True)
    If rsTmp.RecordCount = 0 Then
        Call Err.Raise(ERROR_FORSYSTEM, , "会计期间被破坏，系统无法继续！")
    End If
    CheckDateValidate = strPID

ErrH:
    If Not objDS Is Nothing Then
        Call objDS.rs_Close(rsTmp)
        If vCnn Is Nothing Then Call objDS.CloseConnection
        Set objDS = Nothing
    End If
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, , Err.Description)
    End If
End Function
```

---

## Patch 6：`t_FVou_M.cls` — `meAddRowI` 反射赋值改为显式赋值

### 修改前

```vb
arrFlds = Array("itmid", "fiid", "fino", "finame", "idesc", "CstCenID", _
                "corpid", "corptag", "PrdtID", "PrdtName", "AccID", "Weight", "QTY", "MDQTY", "MCQTY", "Pri", _
                "DATM", "catm", "DATM_F", "catm_f", "AccRate", "SrcItmID", "isFixFI", "TInvID", "ClsID", "ASSInfo", "WHID", "DepID", "EmpID", _
                "MFCstPGBillID", "MFCstPGBillNo", "MFCstMOBillID", "MFCstMOBillNo", "PrjID", "PrjName", "JTID", "JTName")
                
For i = 0 To IDatas.Count - 1
    ...
    For j = 0 To UBound(arrFlds)
        Call meAddRowI(objVou.ItemsData, IDatas.Item(i + 1), arrFlds(j))
    Next
    ...
Next

Private Sub meAddRowI(ByVal rsItem As ADODB.Recordset, ByVal iData As t_FVou_I, ByVal FldName As String)
    rsItem.Fields(FldName).Value = CallByName(iData, FldName, VbGet)
End Sub
```

### 修改后

把整段 `For j` 循环替换为对 `meAddRowI_All` 的单次调用：

```vb
For i = 0 To IDatas.Count - 1
    ...
    Call meAddRowI_All(objVou.ItemsData, IDatas.Item(i + 1))     ' === PR-1 修改 ===
    ...
Next
```

新增的 `meAddRowI_All`：

```vb
'==============================================================================
' PR-1 新增：消除 CallByName 反射，按 arrFlds 顺序逐字段直接赋值
' 字段顺序与原 arrFlds 完全一致，输出 100% 等价
'==============================================================================
Private Sub meAddRowI_All(ByVal rsItem As ADODB.Recordset, ByVal iData As t_FVou_I)
    With rsItem
        .Fields("itmid").Value = iData.ITMID
        .Fields("fiid").Value = iData.FIID
        .Fields("fino").Value = iData.FINO
        .Fields("finame").Value = iData.FIName
        .Fields("idesc").Value = iData.IDesc
        .Fields("CstCenID").Value = iData.CstCenID
        .Fields("corpid").Value = iData.CorpID
        .Fields("corptag").Value = iData.CorpTag
        .Fields("PrdtID").Value = iData.PrdtID
        .Fields("PrdtName").Value = iData.PrdtName
        .Fields("AccID").Value = iData.AccID
        .Fields("Weight").Value = iData.Weight
        .Fields("QTY").Value = iData.QTY
        .Fields("MDQTY").Value = iData.MDQTY
        .Fields("MCQTY").Value = iData.MCQTY
        .Fields("Pri").Value = iData.Pri
        .Fields("DATM").Value = iData.DATM
        .Fields("catm").Value = iData.CATM
        .Fields("DATM_F").Value = iData.DATM_F
        .Fields("catm_f").Value = iData.CATM_F
        .Fields("AccRate").Value = iData.AccRate
        .Fields("SrcItmID").Value = iData.SrcITMID
        .Fields("isFixFI").Value = iData.isFixFI
        .Fields("TInvID").Value = iData.TInvID
        .Fields("ClsID").Value = iData.ClsID
        .Fields("ASSInfo").Value = iData.AssInfo
        .Fields("WHID").Value = iData.WHID
        .Fields("DepID").Value = iData.depid
        .Fields("EmpID").Value = iData.EmpID
        .Fields("MFCstPGBillID").Value = iData.MFCstPGBillID
        .Fields("MFCstPGBillNo").Value = iData.MFCstPGBillNo
        .Fields("MFCstMOBillID").Value = iData.MFCstMOBillID
        .Fields("MFCstMOBillNo").Value = iData.MFCstMOBillNo
        .Fields("PrjID").Value = iData.PrjID
        .Fields("PrjName").Value = iData.PrjName
        .Fields("JTID").Value = iData.JTID
        .Fields("JTName").Value = iData.JTName
    End With
End Sub
```

> `meAddRowI` 旧版本可以保留不删（其他地方还在用），不影响新路径。

---

## 部署顺序

1. 把 `VouMetaCache.bas` 加入 VB6 工程（标准模块）
2. 应用 Patch 6（最简单，最小风险）— 部署，回归测试
3. 应用 Patch 5 — 部署，回归
4. 应用 Patch 4 — 部署，回归
5. 应用 Patch 2 + 3 — 部署，回归
6. 最后应用 Patch 1（启用 LoadAll/ClearAll 入口）— 全部缓存生效

> 之所以最后才应用 Patch 1：前面几个 Patch 在缓存未加载时都自动回退到原 DB 路径，逐步部署时每一步都安全。最后启用入口时，所有缓存路径同时生效。

---

## 等价性约束摘要

- **VouMetaCache 任何 TryGet*** 函数返回 False（即缓存未命中）→ 调用方必须回退到原 DB 路径，行为完全一致
- **VouMetaCache 加载的字段值** 来自相同的 SQL（FIID、FINO、FIName、ISStop、FITag、…），与原代码 SELECT 的字段一一对应
- **错误文案** 在缓存路径下与原路径完全相同（已逐字对照）
- **凭证号生成、ID 生成、Recordset.Update、stored proc 调用**：不动
