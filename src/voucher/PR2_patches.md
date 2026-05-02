# PR-2 优化补丁应用清单

PR-2 在 PR-1 基础上继续优化。涵盖：

- **A2**：getInfoByID 空 schema 缓存（消除 30 万次 SELECT TOP 0）
- ~~**D1**：主查询投影列~~ **已撤销**（见 Patch 3）

## 前置依赖

- PR-1 已合入（VouMetaCache + VouCacheHelpers）
- 新增 `VouSchemaCache.bas` —— **只加到工程 C（POPBus3GL2Service）**
  原因：`megetDocByID` / `getInfoByID` 是 t_FVou_M 类的方法，属于工程 C；
  schema 是 SQL Server 表的字段定义，只在 C 工程内创建空 Recordset 用于
  `sysDS.SaveData` 写入。其他工程不需要此缓存。

## Patch 1：~~`cMthCstAccGL2.meCreateVou` 入口加载 schema 缓存~~（懒加载替代）

⚠️ **跨工程架构下不需要**。`VouSchemaCache` 只在工程 C 使用，由 `megetDocByID` 内部首次调用时自动 EnsureLoaded（懒加载）。

工程 A 的 `meCreateVou` 不需要任何 PR-2 改动。

---

## Patch 2：`t_FVou_M.cls` `megetDocByID` 优先用缓存（工程 C）

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

### 修改后（懒加载版）

```vb
Private Function megetDocByID(ByVal cid As String, ByVal objDS As HHDataService.sysDataService, ...) As Boolean
On Error GoTo ErrHandler
    Dim SQL         As String

    If Trim(cid) = "" And CIDIsEmptyRaiseErr = True Then
        Call Err.Raise(ERROR_FORSYSTEM, , "无效的凭证ID")
    End If

    ' === PR-2 cache fast path：cid="" 时从 schema 缓存克隆（懒加载）===
    If cid = "" Then
        Set mMainData = VouSchemaCache.CloneMainEmptyOrLoad(objDS)
        Set mItemsData = VouSchemaCache.CloneItemsEmptyOrLoad(objDS)
        Set mIItemsData = VouSchemaCache.CloneIItemsEmptyOrLoad(objDS)
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

## Patch 3：~~主查询投影列~~（**已撤销**）

⚠️ **审计发现**：D1 主查询投影列的预期收益只有约 -1.2s（占总耗时 < 0.1%），但**风险高**：
- `Set objIDC.rsBill = rsBill` 把 rs 传给下游 DAL，DAL 内部可能动态访问 `rsBill.Fields("XXX")`
- 投影列方案中遗漏任一字段会导致运行时 `项不在该集合中` 错误
- 我们无法 100% 验证每个 DAL 内部所有字段使用路径

**结论**：保留原代码 `SELECT m.*`，PR-2 只做 A2（schema 缓存），D1 撤销。

如果未来确实需要降低网络流量，应通过**逐字段分析**每个 DAL 类（包括所有 if 分支）后再做投影列改造，并配套完整的 UAT 验证。

<!-- Patch 3 子节已撤销

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

> Patch 3 全部撤销 -->

---

## 部署顺序

1. 把 `VouSchemaCache.bas` 加入 VB6 工程（标准模块）
2. 应用 Patch 2（megetDocByID）→ 编译 → 灰度（此时 schema 缓存还未加载，行为不变）
3. 应用 Patch 1（在 meCreateVou 入口启用 VouSchemaCache.LoadAll）→ schema 缓存生效

部署完成后，A2 节省约 **5 min**（10 万张凭证 / 2 小时基线场景）。
原计划的 D1 投影列已撤销（见 Patch 3 节）。
