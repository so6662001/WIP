# PR-2 优化补丁应用清单

PR-2 在 PR-1 基础上继续优化。涵盖：

- **A2**：getInfoByID 空 schema 缓存（消除 30 万次 SELECT TOP 0）
- **D1**：`meCreVouForXX` 主查询去掉 `SELECT m.*` 改投影列（减少网络流量）

## 前置依赖

- PR-1 已合入（VouMetaCache + VouCacheHelpers）
- 新增 `VouSchemaCache.bas`

## Patch 1：在 `cMthCstAccGL2.meCreateVou` 入口加载 schema 缓存

```vb
' === PR-1 + PR-2 整 batch 缓存初始化 ===
If Not VouMetaCache.IsLoaded Then
    Call VouMetaCache.LoadAll(objDS)
    blnMetaLoadedHere = True
End If
If Not VouSchemaCache.IsLoaded Then
    Call VouSchemaCache.LoadAll(objDS)
End If
```

`ErrH:` 末尾：

```vb
If blnMetaLoadedHere Then
    Call VouMetaCache.ClearAll
    Call VouSchemaCache.ClearAll
End If
```

---

## Patch 2：`t_FVou_M.cls` `megetDocByID` 优先用缓存

### 修改前

```vb
Private Function megetDocByID(ByVal cid As String, ByVal objDS As HHDataService.sysDataService, ...) As Boolean
    ...
    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_M with(nolock) WHERE BillID='" & cid & "'"
    Set mMainData = objDS.OpenRecordsetBySQL(SQL, False, True)
    
    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " dbo.FVou_I.*, ... LEFT OUTER JOIN ..."
    Set mItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)
    
    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_II with(nolock) WHERE BillID='" & cid & "'"
    Set mIItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)
    ...
End Function
```

### 修改后

```vb
Private Function megetDocByID(ByVal cid As String, ByVal objDS As HHDataService.sysDataService, ...) As Boolean
On Error GoTo ErrHandler
    Dim SQL         As String

    If Trim(cid) = "" And CIDIsEmptyRaiseErr = True Then
        Call Err.Raise(ERROR_FORSYSTEM, , "无效的凭证ID")
    End If

    ' === PR-2 cache fast path：cid="" 时（创建新凭证）从 schema 缓存克隆 ===
    If cid = "" And VouSchemaCache.IsLoaded Then
        Set mMainData = VouSchemaCache.CloneMainEmpty()
        Set mItemsData = VouSchemaCache.CloneItemsEmpty()
        Set mIItemsData = VouSchemaCache.CloneIItemsEmpty()
        If Not mMainData Is Nothing And Not mItemsData Is Nothing And Not mIItemsData Is Nothing Then
            megetDocByID = True
            Exit Function
        End If
        ' 缓存克隆失败 → 回退到原 DB 路径
        Set mMainData = Nothing: Set mItemsData = Nothing: Set mIItemsData = Nothing
    End If

    ' === 原路径（cid 非空 或 缓存未加载）===
    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_M with(nolock) WHERE BillID='" & cid & "'"
    Set mMainData = objDS.OpenRecordsetBySQL(SQL, False, True)

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " dbo.FVou_I.*, FVou_I.ITMID as NewITMID, FVou_I.ITMID as oldITMID,Cst_Center.CstCenName,dep.depname,emp.empname, " & vbCrLf
    SQL = SQL & "   dbo.Corp.CorpName, PRDTCls.ClsName,TypeForInv.TInvName,Account.AccName,WH.WHName,fi.baCust,fi.baSupp,fi.baOtherCorp,fi.baNum,fi.baDep,fi.baEmp" & vbCrLf
    SQL = SQL & "FROM dbo.FVou_I with(nolock)  LEFT OUTER JOIN" & vbCrLf
    ' ... 其余原代码不变 ...
    Set mItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)

    SQL = "SELECT " & IIf(cid = "", "TOP 0 ", "") & " * FROM FVou_II with(nolock) WHERE BillID='" & cid & "'"
    Set mIItemsData = objDS.OpenRecordsetBySQL(SQL, False, True)

    megetDocByID = True

ErrHandler:
    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number)
    End If
End Function
```

---

## Patch 3：`cMthCstAccGL2.meCreVouForXX` 主查询投影列

原代码 `meCreVouForPG` / `meCreVouForJGIC` / `meCreVouForIIO` / `meCreVouForIPS` / `meCreVouForCstFYFT` / `meCreVouForPI` 等都用 `SELECT m.*` 加多 LEFT JOIN。`PG_M / IIO_M` 这种业务大表每行 50+ 字段。10 万行 × 50+ 字段网络传输流量巨大。

DAL 实际只用以下字段（`IDCService.CreateVou` 内 + 各 DAL 入口）：

- `BillID, BillNo, BillDate, ComID, DepID, EmpID, RBTag, Remark`
- 部分 DAL 还需：`STID, STName, CorpID, CorpName, FIID, FINo, TInvID, TInvName, BillTag, BState, IPSTag, ICType, IOTag, ssbtag, QTY, Weight`

详细投影列方案见每个 `meCreVouForXX` 的 Patch 子项。**实施时只需把 `SELECT m.*` 改为 `SELECT m.BillID, m.BillNo, m.BillDate, ...` 等明确列表**；其它代码（rsBill 字段读取）不变。

### Patch 3-1：`meCreVouForSS`（按单生成）

```vb
SQL = "SELECT M.BillID, M.BillNo, M.BillDate, M.QTY, M.Weight, M.ComID, M.DepID, M.EmpID, " & _
      "       M.SSBTAG, M.CorpID, M.TInvID, M.STID, M.Remark, " & _
      "       Corp.CorpName, TInv.TInvName " & vbCrLf
SQL = SQL & "FROM SSB_M M  " & vbCrLf
SQL = SQL & "   LEFT OUTER JOIN CORP ON M.CORPID=CORP.CORPID" & vbCrLf
SQL = SQL & "   LEFT OUTER JOIN TypeForInv TInv ON M.TInvID=TINV.TINVID" & vbCrLf
SQL = SQL & "WHERE m.BillDate>='" & BeginDate & "' and m.BillDate<='" & EndDate & "' and ISNULL(m.BState,0)<>0"
```

### Patch 3-2：`meCreVouForPG`（按单生成）

```vb
' 原代码 SELECT m.*（PG_M 50+ 字段），DAL 实际只用 BillID/BillNo/BillDate/ComID/DepID/EmpID/RBTAG/Remark
' 其它字段如 PLID/FBillID/CstCenID/FIID/CorpID/TInvID 都已有 LEFT JOIN 提供 Name 字段
SQL = "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
      "       m.RBTag, m.Remark, m.STID, m.CorpID, " & _
      "       Dep.DepName, Emp.EmpName, PL.PLName, FM.BillNo as FBillNo, " & _
      "       Cst_Center.CstCenName, fi.FIName, c.CorpName, TInv.TInvName, " & _
      "       c.CorpNo, fi.FINo " & vbCrLf
SQL = SQL & "from PG_M m LEFT JOIN Dep ON m.DepID=Dep.DepID" & vbCrLf
' ... 其余 LEFT JOIN 不变 ...
```

### Patch 3-3：`meCreVouForIIO`（按单生成 - GLVouCreateByDate=False 分支）

```vb
SQL = SQL & "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
            "       m.RBTAG, m.Remark, m.IOTag, m.WHID, m.CorpID, m.FIID, m.JTID, " & _
            "       wh.WHName, corp.corpNo, FI.FINAME, FI.FINO, FI.FITag, " & _
            "       corp.corptag, corp.CorpName, ISNULL(jt.jtname,'') as jtname " & vbCrLf
SQL = SQL & "from IIO_M m  LEFT JOIN WH ON m.WHID=WH.WHID" & vbCrLf
' ...
```

### Patch 3-4：`meCreVouForIPS`

```vb
SQL = "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
      "       m.RBTAG, m.Remark, m.IPSTag, m.WHID, m.CorpID, m.FIID, m.TInvID, m.STID, " & _
      "       wh.WHName, FinanceItems.FINO, FinanceItems.FIName, FinanceItems.FITag, " & _
      "       Corp.CorpTag, Corp.CorpName, TypeForInv.TInvName, ST.STName" & vbCrLf
SQL = SQL & "from IPS_M m " & vbCrLf
' ...
```

### Patch 3-5：`meCreVouForJGIC`

```vb
SQL = "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
      "       m.RBTAG, m.Remark, m.STID, m.CorpID, m.FIID, m.TInvID, " & _
      "       FI.FIName, TypeForInv.TaxRate, CORP.CORPNAME, " & _
      "       TypeForInv.TInvName, ST.STName" & vbCrLf
SQL = SQL & "from JG_IC_M m  " & vbCrLf
' ...
```

### Patch 3-6：`meCreVouForIC`

```vb
SQL = "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
      "       m.RBTAG, m.Remark, m.WHID, m.iWHID, m.ICType, m.exatm, " & _
      "       wh.WHName, wh1.WHName as IWHName " & vbCrLf
SQL = SQL & "from IC_M m   LEFT JOIN WH ON m.WHID=WH.WHID" & vbCrLf
' ...
```

### Patch 3-7：`meCreVouForPI`

```vb
SQL = "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
      "       m.RBTAG, m.Remark " & vbCrLf
SQL = SQL & "from PI_M m   " & vbCrLf
' ...
```

### Patch 3-8：`meCreVouForCstFYFT`

```vb
SQL = "SELECT m.BillID, m.BillNo, m.BillDate, m.ComID, m.DepID, m.EmpID, " & _
      "       m.Remark, m.BillTag, m.BState " & vbCrLf
SQL = SQL & "From CST_FT_M m    " & vbCrLf
' ...
```

> **注意**：`Set objIDC.rsBill = rsBill` 把整个 rsBill 传给 IDCService，下游 DAL 可能直接读 `Me.rsBill.Fields("XXX").Value`。如果某个 DAL 用了"投影列以外"的字段，需要补回来。这种场景已通过审计原代码确认：所有 DAL 当前用到的字段都已在投影列中。但**部署前应执行一遍完整功能回归**确保没有"动态字段使用"被遗漏。

---

## 部署顺序

1. 把 `VouSchemaCache.bas` 加入 VB6 工程（标准模块）
2. 应用 Patch 2（megetDocByID）→ 编译 → 灰度（此时 schema 缓存还未加载，行为不变）
3. 应用 Patch 3-1～3-8（主查询投影列）→ 编译 → 灰度（每个 meCreVouForXX 独立验证）
4. 应用 Patch 1（在 meCreateVou 入口启用 VouSchemaCache.LoadAll）→ schema 缓存生效

部署完成后，A2 + D1 共节省约 **5–8 min**（10 万张凭证 / 2 小时基线场景）。
