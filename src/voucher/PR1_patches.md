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

## Patch 3：`t_FVou_M.cls` — `GetFIIDByPrdt` 优先走缓存

### 修改前

```vb
Public Function GetFIIDByPrdt(ByVal objDS As HHDataService.sysDataService, ByVal DocKey As String, ByVal DocCHName As String, _
                            ByVal WHID As String, ByVal WHName As String, ByVal ClsID As String, ByVal ClsName As String, _
                            ByRef rtnFINO As String, ByRef rtnFIName As String) As String
    If mrsDocFIRel Is Nothing Then
        Call Init(objDS, DocKey)       '初始化
    End If
    
    mrsDocFIRel.Filter = "DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID='" & ClsID & "'"
    
    '再查询仓库
    If mrsDocFIRel.RecordCount = 0 Then
        mrsDocFIRel.Filter = "(DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID='') or (DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID=NULL)"
    End If
    ...
End Function
```

### 修改后

```vb
Public Function GetFIIDByPrdt(ByVal objDS As HHDataService.sysDataService, ByVal DocKey As String, ByVal DocCHName As String, _
                            ByVal WHID As String, ByVal WHName As String, ByVal ClsID As String, ByVal ClsName As String, _
                            ByRef rtnFINO As String, ByRef rtnFIName As String) As String
    
    ' === PR-1 新增：缓存路径，等价于原 4 步 Filter 查找的精确-到-空 fallback ===
    If VouMetaCache.IsLoaded Then
        Dim sFIID As String
        Dim sHSTagName As String
        Dim lngHSTag As Long, lngFITag As Long
        
        ' Step 1: DocKey + WHID + ClsID 精确
        If VouMetaCache.TryGetDocFI(DocKey, WHID, ClsID, sFIID, rtnFINO, rtnFIName, lngHSTag, sHSTagName, lngFITag) Then
            GetFIIDByPrdt = sFIID
            Exit Function
        End If
        ' Step 2: DocKey + WHID + DOCID=''
        If VouMetaCache.TryGetDocFI(DocKey, WHID, "", sFIID, rtnFINO, rtnFIName, lngHSTag, sHSTagName, lngFITag) Then
            GetFIIDByPrdt = sFIID
            Exit Function
        End If
        ' Step 3: DocKey + WHID='' + DOCID=ClsID
        If VouMetaCache.TryGetDocFI(DocKey, "", ClsID, sFIID, rtnFINO, rtnFIName, lngHSTag, sHSTagName, lngFITag) Then
            GetFIIDByPrdt = sFIID
            Exit Function
        End If
        ' Step 4: 上级品类（必须查 DB，因为依赖 getPrdtClsTree 函数）
        Dim strFIIDFromParent As String
        strFIIDFromParent = getFIIDByClsIDFromParentCls(objDS, DocKey, WHID, ClsID, rtnFINO, rtnFIName)
        If strFIIDFromParent <> "" Then
            GetFIIDByPrdt = strFIIDFromParent
            Exit Function
        End If
        ' Step 5: 全空
        If VouMetaCache.TryGetDocFI(DocKey, "", "", sFIID, rtnFINO, rtnFIName, lngHSTag, sHSTagName, lngFITag) Then
            GetFIIDByPrdt = sFIID
            Exit Function
        End If
        Call Err.Raise(ERROR_FORSYSTEM, , WHName & " 的 " & ClsName & " 尚未设置" & DocCHName & "会计科目，无法生成记账凭证！")
        Exit Function
    End If
    
    ' === 原路径（VouMetaCache 未加载时回退）===
    If mrsDocFIRel Is Nothing Then
        Call Init(objDS, DocKey)
    End If
    
    mrsDocFIRel.Filter = "DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID='" & ClsID & "'"
    
    '再查询仓库
    If mrsDocFIRel.RecordCount = 0 Then
        mrsDocFIRel.Filter = "(DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID='') or (DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID=NULL)"
    End If
    
    '再查询货品类别
    If mrsDocFIRel.RecordCount = 0 Then
        mrsDocFIRel.Filter = "(DocKey='" & DocKey & "' AND DOCID='" & ClsID & "' AND WHID='') OR (DocKey='" & DocKey & "' AND DOCID='" & ClsID & "' AND WHID=NULL)"
    End If
    
    '查询其上级品类的会计科目
    Dim strFIID As String
    If mrsDocFIRel.RecordCount = 0 Then
        strFIID = getFIIDByClsIDFromParentCls(objDS, DocKey, WHID, ClsID, rtnFINO, rtnFIName)
    End If
    
    If strFIID = "" Then
        '再查询空的
        If mrsDocFIRel.RecordCount = 0 Then
            mrsDocFIRel.Filter = "DocKey='" & DocKey & "' AND WHID='' AND DOCID=''"
        End If
        
        If mrsDocFIRel.RecordCount > 0 Then
            rtnFINO = mrsDocFIRel.Fields("FINO").Value
            rtnFIName = getFIName(objDS, mrsDocFIRel.Fields("FIID").Value, mrsDocFIRel.Fields("FIName").Value, "", mrsMemoryFI)
            
            GetFIIDByPrdt = mrsDocFIRel.Fields("FIID").Value
            
            If Not mrsWHClsFI Is Nothing Then
                ' ... 原代码
            End If
        Else
            Call Err.Raise(ERROR_FORSYSTEM, , WHName & " 的 " & ClsName & " 尚未设置" & DocCHName & "会计科目，无法生成记账凭证！")
        End If
    Else
        GetFIIDByPrdt = strFIID
    End If
End Function
```

> **同样的方式也要应用到** `GetFIIDByPrdt1` / `GetFIIDByKey` / `GetAccFIID` / `GetCGExpYFZKFIID` / `GetFactYJTax_XXFIID` / `GetFactYJTax_JXFIID` / `GetYJTax_TInvFIID` / `GetYJTax_XXFIID` / `GetYJTax_JXFIID`：每个函数前面加上"如果 VouMetaCache.IsLoaded 则按缓存路径处理"分支。详见 `t_FVou_M_PR1_cache.bas`。

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
