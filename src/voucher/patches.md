# PR-1 应用 patches

## 跨工程架构

凭证生成涉及 3 个 VB6 ActiveX DLL 工程：

```
POPBus3FileService.dll  (工程 A)  meCreateVou + meCreVouForXX
POPBus3GL2IDC.dll       (工程 B)  IDCService + 各 DAL
POPBus3GL2Service.dll   (工程 C)  t_FVou_M + CVouService
```

PR-1 改动**全部集中在工程 C**（除了 .bas 文件需要加到三个工程外）。

## 文件部署

| .bas 文件 | 加到工程 |
|---|---|
| `VouMetaCache.bas` | A + B + C 三个工程 |
| `VouCacheHelpers.bas` | C |

按顺序应用以下修改。每条都标注**所属工程**。

---

## P1.1：`t_FVou_M.cls` — `Init()` 缓存优先（**工程 C**）

### 修改前

```vb
Private Sub Init(ByVal objDS As HHDataService.sysDataService, ByVal DocKey As String)
    Dim SQL As String
    SQL = "SELECT isnull(dockey,'') as dockey,  ISNULL(D.DocID,'') AS DocID, ISNULL(D.WHID,'') AS WHID, D.FIID, FI.FINo, ..."
    SQL = SQL & "FROM dbo.GL2_AchFIID D WITH(NOLOCK) INNER JOIN" & vbCrLf
    SQL = SQL & "   dbo.FinanceItems FI ON D.FIID = FI.FIID" & vbCrLf
    SQL = SQL & "WHERE ISNULL(D.FIID,'')<>''"
    Set mrsDocFIRel = objDS.OpenRecordsetBySQL(SQL, True, True)
End Sub
```

### 修改后

```vb
Private Sub Init(ByVal objDS As HHDataService.sysDataService, ByVal DocKey As String)
    Dim SQL As String

    ' === PR-1 A4：cache 已加载时跳过 GL2_AchFIID 全表 SELECT ===
    Call VouMetaCache.EnsureLoaded(objDS)
    If VouMetaCache.IsLoaded Then
        ' mrsDocFIRel 保持 Nothing，下游 GetFIIDBy* 走 VouCacheHelpers
        Exit Sub
    End If

    ' === 原路径（cache 加载失败回退）===
    SQL = "SELECT isnull(dockey,'') as dockey, ..."
    Set mrsDocFIRel = objDS.OpenRecordsetBySQL(SQL, True, True)
End Sub
```

---

## P1.2：`t_FVou_M.cls` — `GetFIIDByPrdt` 5 步 fallback（**工程 C**）

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

    ' === 原 mrsDocFIRel 路径全部保留作为 fallback ===
    If mrsDocFIRel Is Nothing Then
        Call Init(objDS, DocKey)
    End If
    mrsDocFIRel.Filter = "DocKey='" & DocKey & "' AND WHID='" & WHID & "' AND DOCID='" & ClsID & "'"
    ' ... 原代码所有逻辑保留 ...
End Function
```

---

## P1.3：`t_FVou_M.cls` — `GetFIIDByPrdt1` 同 P1.2，但 `isStrict=False`

```vb
Public Function GetFIIDByPrdt1(...) As String
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_5Step(objDS, DocKey, DocCHName, WHID, WHName, _
                                         ClsID, ClsName, False, _
                                         sFIID, rtnFINO, rtnFIName) Then
        GetFIIDByPrdt1 = sFIID
        Exit Function
    End If
    ' === 原路径保留 ===
End Function
```

---

## P1.4：`t_FVou_M.cls` — `getFIIDByClsIDFromParentCls`（被 5 步调用）

```vb
Private Function getFIIDByClsIDFromParentCls(ByVal objDS As HHDataService.sysDataService, _
                                             ByVal DocKey As String, ByVal WHID As String, _
                                             ByVal ClsID As String, _
                                             ByRef rtnFINO As String, ByRef rtnFIName As String) As String
    ' === PR-1 cache fast path ===
    Call VouMetaCache.EnsureLoaded(objDS)
    If VouMetaCache.IsLoaded Then
        Dim sFIID As String
        If VouCacheHelpers.LookupDocFI_ParentCls(objDS, DocKey, WHID, ClsID, _
                                                 sFIID, rtnFINO, rtnFIName) Then
            getFIIDByClsIDFromParentCls = sFIID
        End If
        Exit Function
    End If
    ' === 原 mrsDocFIRel 路径保留 ===
End Function
```

---

## P1.5：`t_FVou_M.cls` — `GetAccFIID` / `GetCGExpYFZKFIID` / `GetFIIDByKey` 等

简单 2 步 fallback（DocID 精确 → DocID=''）：

```vb
Public Function GetAccFIID(ByVal objDS As HHDataService.sysDataService, _
                           ByVal AccID As String, ByVal AccName As String, _
                           ByRef rtnFINO As String, ByRef rtnFIName As String) As String
    Dim sFIID As String
    If VouCacheHelpers.LookupDocFI_KeyOnly(objDS, "Account", "货币资金", _
                                           AccID, AccName, _
                                           True, AccName & " 尚未设置货币资金会计科目，无法生成记账凭证！", _
                                           sFIID, rtnFINO, rtnFIName) Then
        GetAccFIID = sFIID
        Exit Function
    End If
    ' === 原路径保留 ===
End Function
```

同模式应用到：

| 函数 | DocKey | 错误文案 |
|---|---|---|
| `GetCGExpYFZKFIID` | `"CGExpFI"` | `CGExpName & " 尚未设置采购费用会计科目..."` |
| `GetFactYJTax_XXFIID` | `"TInvXXFact"` | `TInvName & " 尚未设置应交税费销项会计科目..."` |
| `GetFactYJTax_JXFIID` | `"TInvJXFact"` | 进项 |
| `GetYJTax_XXFIID` | `"TInvXX"` | 计提应交税费销项 |
| `GetYJTax_JXFIID` | `"TInvJX"` | 计提应交税费进项 |
| `GetYJTax_TInvFIID` | 调用方传入 DocKey | 调用方传入文案 |
| `GetFIIDByKey` | 调用方传入 DocKey | 调用方传入文案 |

---

## P1.6：`t_FVou_M.cls` — `meAddRowI_All` 替代 CallByName（**工程 C**）

`CreateVou` 内原 For 循环：

```vb
For j = 0 To UBound(arrFlds)
    Call meAddRowI(objVou.ItemsData, IDatas.Item(i + 1), arrFlds(j))
Next
```

替换为：

```vb
Call meAddRowI_All(objVou.ItemsData, IDatas.Item(i + 1))
```

新增私有过程：

```vb
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

> 旧 `meAddRowI` 可保留（其它地方可能在用），不影响。

---

## P1.7：`CVouService.cls` — `BeforeAction` 调用 helper（**工程 C**）

```vb
Private Sub BeforeAction(...)
    ' === 保留 4 段 IN-list SQL 作为 fallback（cache 未加载时使用）===
    If strFIIDList <> "" Then
        strFIIDList = Left(strFIIDList, Len(strFIIDList) - 1)
        SQL = "SELECT FIID,FIName,FINO,ISStop,FITag,ExpTag,HSTag,HSTagName,baNum,baCust,baSupp,baOtherCorp,fullname FROM FinanceItems WHERE FIID in (" & strFIIDList & ")"
        Set rsfcode = objDS.OpenRecordsetBySQL(SQL, True, True)
    End If
    ' ... 同样保留 Corp / Emp / Acc 的 IN-list SQL ...

    Call objDS.rs_MoveFirst(ItemsData)
    Do While Not ItemsData.EOF
        ' helper 内部首先 EnsureLoaded VouMetaCache
        ' 命中 → 用缓存；未命中 → 用 rsfcode 等
        Call VouCacheHelpers.CheckFI  (objDS, ItemsData, rsfcode, rsFIBak, ActCHName, strBillInfo, strErrInfo)
        Call VouCacheHelpers.CheckAcc (objDS, ItemsData, rsAccs, ActCHName, strBillInfo, strErrInfo)
        Call VouCacheHelpers.CheckCorp(objDS, ItemsData, rsCorps, ActCHName, strBillInfo, strErrInfo)
        Call VouCacheHelpers.CheckEmp (objDS, ItemsData, rsEmps, ActCHName, strBillInfo, strErrInfo)

        ItemsData.Update    ' === 保留原 ItemsData.Update ===

        ' ... 后续 Select Case ItemsData.Fields("HSTag").Value 保持原代码 ...

        ItemsData.MoveNext
    Loop
End Sub
```

---

## P1.8：`CVouService.cls` — `CheckDateValidate` 走缓存（**工程 C**）

```vb
Public Function CheckDateValidate(ByVal vDate As Date, Optional ByVal vCnn As ADODB.Connection = Nothing) As String
On Error GoTo ErrH
    Dim rsTmp   As ADODB.Recordset
    Dim objDS   As sysDataService
    Dim strPID  As String
    Dim strAPTagFld As String

    strAPTagFld = IIf(Me.AppParameters.GL2_ERPFODiffCarry, "FOAPTag", "APTag")

    ' === PR-1 cache fast path ===
    Set objDS = New sysDataService
    objDS.ConnectionString = Me.ConnectionString
    If Not vCnn Is Nothing Then
        Set objDS.Connection = vCnn
    Else
        Call objDS.OpenConnection
    End If
    Call VouMetaCache.EnsureLoaded(objDS)
    If VouMetaCache.IsLoaded Then
        Dim sAPID As String
        Dim lngAPTag As Long
        If Not VouMetaCache.TryGetAPByDate(vDate, Me.AppParameters.GL2_ERPFODiffCarry, sAPID, lngAPTag) Then
            Call Err.Raise(ERROR_FORSYSTEM, , "无法继续，因为您录入的日期不在本年度会计期间内！")
        End If
        Select Case CInt(lngAPTag)
            Case 1: Call Err.Raise(ERROR_FORSYSTEM, , "您不能录入已月结会计期间的凭证！")
            Case -1: Call Err.Raise(ERROR_FORSYSTEM, , "您不能录入以前会计期间的凭证！")
        End Select
        If Not VouMetaCache.HasActiveAP(Me.AppParameters.GL2_ERPFODiffCarry) Then
            Call Err.Raise(ERROR_FORSYSTEM, , "会计期间被破坏，系统无法继续！")
        End If
        CheckDateValidate = sAPID
        GoTo ErrH
    End If

    ' === 原路径（cache 未加载时回退）===
    Set rsTmp = objDS.OpenRecordsetBySQL("SELECT APID," & strAPTagFld & " as APTag FROM AccPeriod WHERE BeginDate<='" & Format(vDate, "yyyy-MM-dd") & "' AND EndDate>='" & Format(vDate, "yyyy-MM-dd") & "'", True, True)
    ' ... 原代码所有逻辑保留 ...

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

## 部署核对清单

- [ ] `VouMetaCache.bas` 添加到工程 **A**（POPBus3FileService）
- [ ] `VouMetaCache.bas` 添加到工程 **B**（POPBus3GL2IDC）
- [ ] `VouMetaCache.bas` 添加到工程 **C**（POPBus3GL2Service）
- [ ] `VouCacheHelpers.bas` 添加到工程 **C**（POPBus3GL2Service）
- [ ] 应用 P1.1～P1.5 到 `t_FVou_M.cls`（工程 C）
- [ ] 应用 P1.6 到 `t_FVou_M.CreateVou`（工程 C）
- [ ] 应用 P1.7 到 `CVouService.BeforeAction`（工程 C）
- [ ] 应用 P1.8 到 `CVouService.CheckDateValidate`（工程 C）
- [ ] 编译三个工程并部署 DLL
- [ ] 灰度测试：跑一个月凭证生成对比 baseline 行为
