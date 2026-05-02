# 三工程架构审计与重构方案

## 用户后告知的关键架构信息（2026-05-02）

```
┌────────────────────────────────────────────────────────────────┐
│  POPBus3FileService.dll  (工程 A)                              │
│  ├─ CMthCstAccService.meCreateVou             ← 顶层入口        │
│  └─ cMthCstAccGL2.meCreVouFor*                                 │
│     └─ ↓ COM call                                              │
└────────────────────────────────────────────────────────────────┘
┌────────────────────────────────────────────────────────────────┐
│  POPBus3GL2IDC.dll       (工程 B)                              │
│  ├─ IDCService.CreateVou                                       │
│  ├─ SSBillByDateDAL  / PGBillDAL  / ...   ← 各 DAL              │
│  └─ ↓ COM call                                                 │
└────────────────────────────────────────────────────────────────┘
┌────────────────────────────────────────────────────────────────┐
│  POPBus3GL2Service.dll   (工程 C)                              │
│  ├─ t_FVou_M / t_FVou_I / t_FVou_II / t_FVouIDatas / ...       │
│  ├─ CVouService.SaveDoc            ← PR-1/PR-2/PR-4 改造点      │
│  ├─ CVouService.BeforeAction       ← PR-1 改造点                │
│  ├─ CVouService.CheckDateValidate  ← PR-1 改造点                │
│  ├─ t_FVou_M.Init / GetFIIDByPrdt  ← PR-1 改造点                │
│  ├─ t_FVou_M.megetDocByID          ← PR-2 改造点                │
│  └─ t_FVou_M.CreateVou.meAddRowI   ← PR-1 A3 改造点             │
└────────────────────────────────────────────────────────────────┘
```

## VB6 跨工程的关键约束

1. **每个 ActiveX DLL 工程的标准模块（.bas）状态独立**——三个工程即使各自引用了"同一份" `VouMetaCache.bas`，它们看到的是各自工程的 `m_dctFI`，**完全不共享**。
2. **类模块（.cls）的对象实例可以跨工程传递**，但每次 `New` 都是独立实例。
3. **同一个 Connection / spid** 内的 `#temp` 表跨调用、跨工程**可见**（这是关键）。
4. **`@@TRANCOUNT` 在同一连接里全局可读**。

## 上一版 PR-1～PR-4 在跨工程下的失效情况

| PR | 改动 | 失效原因 | 实际效果 |
|---|---|---|---|
| PR-1 A1 | A 工程 `meCreateVou` 调 `VouMetaCache.LoadAll` | 加载到 A 工程的 `m_dctFI`；C 工程 BeforeAction 看到 C 自己的 `m_dctFI=Nothing` | **回退原 DB 查询，缓存白搭** |
| PR-1 A4 | C 工程 `t_FVou_M.Init` 检查 `VouMetaCache.IsLoaded` | C 工程的 `m_blnLoaded=False` | **走原 mrsDocFIRel 路径** |
| PR-1 A5 | C 工程 `BeforeAction` Dictionary 查找 | 同上 | **走原 IN 列表 SELECT** |
| PR-1 A6 | C 工程 `CheckDateValidate` Dictionary 查找 | 同上 | **走原 AccPeriod SELECT** |
| PR-2 A2 | C 工程 `megetDocByID` 调 `VouSchemaCache.CloneXxxEmpty` | C 工程的 `m_blnLoaded=False`，Clone 返回 Nothing | **回退 SELECT TOP 0** |
| PR-3 B2/B3 | C 工程 SaveDoc 调 `VouBatchOps.RecordSavedBill` | C 工程的 `m_colSavedBills_Direct=Nothing` 且 `m_blnInBatchMode=False` | **走原 per-bill stored proc，批量化失效** |
| PR-3 B1 | A 工程 `meCreateVou` 包大事务 | C 工程 SaveDoc 内部仍 `BeginTrans`（看不到外层事务） | **嵌套事务 / commit 错位** |

**PR-1 / PR-2 / PR-3 的核心优化基本全部失效**。需要按下面方案重构。

---

## 重构方案（最终版）

### 方案 1：元数据 / Schema 缓存（PR-1 + PR-2 A2）→ 三工程独立懒加载

**原理**：FinanceItems / Corp / Emp / Account / AccPeriod / GL2_AchFIID / F_VouCls_Bills 都是**只读元数据**，整 batch 期间不会变。每个工程**独立持有一份缓存**，数据完全相同。

**实现**：
- `VouMetaCache.bas` 加到三个工程（A/B/C 都引用同一份 .bas 文件）
- 改 `IsLoaded` 检查为：**懒加载**，首次 TryGet 时如果 `m_blnLoaded=False`，自动调 `LoadAll(objDS)`
- **不主动 ClearAll**——让缓存在 DLL 进程生命期内一直存在（DLL 卸载时自然释放）
- 每个工程加载 7 次 SQL，3 个工程一共 21 次（一次性，不在主循环里）

**收益**：
- 10 万张凭证 × 5 行明细 × 6 次 DB 查询 = 300 万次 RPC → **21 次 RPC**
- 跨工程懒加载比 PR-1 原方案多 14 次 SQL（B/C 各加载 7 次），但仍然 < 0.05% 总耗时

### 方案 2：VouSchemaCache（PR-2 A2）

同方案 1，**只在 C 工程加载即可**——`megetDocByID` 是 C 工程的方法。

### 方案 3：大事务（PR-3 B1）→ `@@TRANCOUNT` 判断

**原理**：SQL Server 的 `@@TRANCOUNT` 在同一连接（spid）内全局可读，事务嵌套时 `@@TRANCOUNT > 0`。

**实现**：
- A 工程 `meCreateVou` 入口 `BeginTrans`
- C 工程 `SaveDoc` 内部检查：

```vb
Dim rsTC As ADODB.Recordset
Set rsTC = sysDS.OpenRecordsetBySQL("SELECT @@TRANCOUNT AS tc", True, True)
If rsTC.Fields("tc").Value > 0 Then
    ' 已在外层事务 → 跳过 BeginTrans/CommitTrans
    blnHasOuterTxn = True
End If
```

或更简单，**不查询**：直接判断 `objDS.Connection.State` + 一个调用方传入的"已在 batch 模式"标志。但这又回到了 .bas 跨工程问题。

**最稳妥实现**：**接口扩展**。给 `t_FVou_M` 类增加一个 Public 属性 `IsInOuterBatch As Boolean`，A 工程 `meCreVouFor*` 在创建 `objIDC` 后设置 `objIDC.IsInOuterBatch = True`；C 工程 `SaveDoc` 检查这个标志（属性赋值是 COM 调用，类实例跨工程传递时此值保留）。

### 方案 4：BillID 收集（PR-3 B2/B3）→ SQL Server 临时表

**原理**：同一 Connection 的 `#temp` 表跨调用、跨工程可见。

**实现**：
- A 工程 `meCreateVou` 入口创建：

```sql
IF OBJECT_ID('tempdb..#VouBatchBills_Direct') IS NOT NULL DROP TABLE #VouBatchBills_Direct;
CREATE TABLE #VouBatchBills_Direct (rn INT IDENTITY(1,1), BillID VARCHAR(48));
IF OBJECT_ID('tempdb..#VouBatchBills_Trans') IS NOT NULL DROP TABLE #VouBatchBills_Trans;
CREATE TABLE #VouBatchBills_Trans (rn INT IDENTITY(1,1), BillID VARCHAR(48));
```

- C 工程 `SaveDoc` 末尾：

```vb
' 检测 batch 模式：临时表存在则启用
Dim rsCheck As ADODB.Recordset
Set rsCheck = sysDS.OpenRecordsetBySQL( _
    "SELECT OBJECT_ID('tempdb..#VouBatchBills_Direct') AS tid", True, True)
If Not IsNull(rsCheck.Fields("tid").Value) Then
    ' batch 模式：INSERT 到临时表
    If isSaveToTransitionalTable Then
        Call sysDS.ExecSQL("INSERT INTO #VouBatchBills_Trans(BillID) VALUES('" & strBillID & "')")
    Else
        Call sysDS.ExecSQL("INSERT INTO #VouBatchBills_Direct(BillID) VALUES('" & strBillID & "')")
    End If
Else
    ' 非 batch 模式：原 per-bill stored proc 调用
    If isSaveToTransitionalTable = False Then
        Call sysDS.ExecSQL("exec ss_YWVouUpdateAchTimes '" & strBillID & "',1")
        Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',0")
    Else
        Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',1")
    End If
End If
```

但**还有一个性能问题**：每张凭证 INSERT 到 #temp 是 1 次 RPC = 10 万次 RPC。**不省**！

**优化**：把 OBJECT_ID 检查也消除——直接给 `t_FVou_M` 类一个 `BatchModeBillIDTable` 属性（"#VouBatchBills_Direct" 字符串），A 工程在 IDC 之间设置好；C 工程读这个属性而不是查询 OBJECT_ID。

或者更狠：**给 `IDCService` / `t_FVou_M` 都加一个 `BatchOps` 属性**（接口型），A 工程 BeginBatch 后设置；C 工程 SaveDoc 直接调这个对象的方法（COM 调用，跨工程透明）。

**最终方案**：

```vb
' POPBus3GL2Service.dll 工程 C 内：
' t_FVou_M.cls 增加属性
Public BatchMode_BillIDTable As String   ' 空 = 非 batch 模式
                                          ' 非空（如 "#VouBatchBills_Direct"）= 批量插入到此表

' POPBus3GL2IDC.dll 工程 B 内：
' IDCService.cls 增加属性
Public BatchMode_BillIDTable As String

' POPBus3FileService.dll 工程 A 内：
' meCreateVou 入口：
Call objDS.ExecSQL("CREATE TABLE #VouBatchBills_Direct(...)")
Call objDS.ExecSQL("CREATE TABLE #VouBatchBills_Trans(...)")
objIDC.BatchMode_BillIDTable_Direct = "#VouBatchBills_Direct"
objIDC.BatchMode_BillIDTable_Trans  = "#VouBatchBills_Trans"

' IDCService.CreateVou 内创建 t_FVou_M 时：
Set objVouD = New t_FVou_M
objVouD.BatchMode_BillIDTable_Direct = Me.BatchMode_BillIDTable_Direct
objVouD.BatchMode_BillIDTable_Trans  = Me.BatchMode_BillIDTable_Trans
```

**问题**：每次 SaveDoc 仍要 INSERT 一行到 #temp 表 = 10 万次 RPC。

**真正的解法**：**让 A 工程在循环外攒 BillID**。但 BillID 在 SaveDoc 内才生成（CreateVouNo），A 工程拿不到。

**唯一可行的真批量方案**：**让 C 工程 `t_FVou_M.CreateVou` 返回 BillID 给 IDCService**（已经返回 `rtnVouBillID`），IDCService 返回给 A 工程；A 工程 `meCreVouFor*` 循环里收集到内存集合，循环结束后**一次性 INSERT** 到 #temp + Flush。

让我看 A 工程 meCreVouFor 是否能拿到 BillID：

```vb
Call objIDC.CreateVou(objDS, False)
' objIDC.rtnVouBillID 现在是新生成的凭证 BillID
```

是的，`objIDC.rtnVouBillID` 在 CreateVou 后可用。A 工程可以收集这个 BillID。

**最终最终方案**：

```vb
' 工程 A：meCreVouForXX 循环里：
Do While Not rsBill.EOF
    ...
    Call objIDC.CreateVou(objDS, False)
    
    ' === PR-3 batch 收集 ===
    If g_strBatchMode_BillIDTable_Direct <> "" Then
        ' 直接 INSERT 到 #temp (跨工程已在事务内 spid 共享)
        ' 但这是每张 1 次 RPC，没省
        ' 改为：A 工程内存收集
        g_colBatchBills_Direct.Add objIDC.rtnVouBillID
    End If
    ...
Loop

' meCreateVou 末尾：FlushStoredProcs 时把内存 collection 一次性 INSERT 到 #temp
```

这就是 PR-3 原 VouBatchOps.bas 方案，**只要在工程 A 内使用就行**！工程 A 内部的 `m_colSavedBills_Direct` 在工程 A 自己看是一致的。

**关键洞察**：原 PR-3 方案在工程 A 内部就能完成 BillID 收集和 Flush——只要 **`RecordSavedBill` 的调用方是 A 工程而不是 C 工程**！

只需要修改：**让 A 工程的 `meCreVouForXX` 循环每次调 CreateVou 后，调一次 `VouBatchOps.RecordSavedBill(objIDC.rtnVouBillID, isSaveToTransitionalTable)`**。

这样：
- VouBatchOps.bas 只需加到工程 A
- 工程 C 的 SaveDoc 完全不动（PR-3 Patch 2/3 撤销）
- 工程 A 调 `objDS.BeginTrans` 大事务（PR-3 B1）
- 工程 C SaveDoc 内部 BeginTrans/CommitTrans 仍然嵌套——SQL Server 嵌套事务的内层 commit 是 no-op，**实际生效是 A 工程的最外层 commit**！原 PR-3 担心的"嵌套事务问题"实际上 SQL Server 自动处理，commit 计数到 0 才真 commit ✓

但还有一个语义问题：**A 工程怎么知道 C 工程的 SaveDoc 内部跳没跳 stored proc？**

答：**让 A 工程调用一个新增的 `objIDC.SetBatchMode(True)` 属性，IDC 和 t_FVou_M 都接收这个标志，C 工程 SaveDoc 内部判断这个标志决定是否调 stored proc**。

OK，改造路径清晰了：

### 最终最终方案（真）

| 改造 | 加到哪 | 说明 |
|---|---|---|
| `VouMetaCache.bas` | 三个工程都加 | 懒加载 |
| `VouSchemaCache.bas` | 工程 C 加 | C 工程内部使用 |
| `VouBatchOps.bas` | 工程 A 加 | A 工程内部 BillID 收集 |
| `VouBatchWriter.bas` | 工程 C 加 | C 工程内部使用 |
| 工程 A 修改 | `meCreateVou` 大事务 + 创建 #temp 表 + BeginBatch；`meCreVouForXX` 循环内 `RecordSavedBill(rtnVouBillID)`；末尾 FlushStoredProcs | 见 PR-3 |
| 工程 B 修改 | `IDCService` 增加 `BatchMode As Boolean` 属性，转发给 `t_FVou_M`；`SSBillByDateDAL` 等做 C1 预聚合 | 见 PR-3/PR-4 |
| 工程 C 修改 | `t_FVou_M` 增加 `BatchMode As Boolean` 属性；`CVouService.SaveDoc` 内部根据 BatchMode 决定是否调 per-bill stored proc | 见 PR-3 |

---

## 跨工程通信总结

| 数据 | 通信方式 |
|---|---|
| 元数据缓存内容 | **每工程独立加载**（懒加载，数据相同）|
| schema 缓存内容 | **C 工程独立加载** |
| BatchMode 标志 | **类属性传递**（A → IDC → t_FVou_M）|
| BillID list | **A 工程内存 collection**（A 工程在 meCreVouForXX 循环里收集 `objIDC.rtnVouBillID`）|
| 事务上下文 | **SQL Server 自动嵌套事务**（@@TRANCOUNT 计数）|
| #temp 表 | **同一 connection spid 共享**（用于 stored proc 批量调用）|
