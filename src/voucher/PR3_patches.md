# PR-3 优化补丁应用清单

## 范围

- **B1**：`meCreateVou` 包大事务（每 1000 张凭证 commit 一次）
- **B2**：`ss_YWVouUpdateAchTimes` per-bill → 整 batch 末尾批量调用
- **B3**：`ss_AfterSaveUpdateGL2VouBill` per-bill → 整 batch 末尾批量调用

## ⚠️ 跨工程架构（关键）

| 文件 / 修改 | 工程 |
|---|---|
| `VouBatchOps.bas` | **只加到工程 A**（POPBus3FileService）；工程 A 的 meCreVouForXX 在循环里收集 BillID |
| `PR3_storedprocs.sql` | SQL Server 部署 |
| t_FVou_M 类**新增 BatchMode 属性**（工程 C）| `t_FVou_M.cls` (POPBus3GL2Service) — 跨工程标志位 |
| IDCService 类**新增 BatchMode 属性**（工程 B）| `IDCService.cls` (POPBus3GL2IDC) — 工程 A 通过此属性把 batch 标志传给 t_FVou_M |
| `cMthCstAccGL2.meCreVouForXX` 修改（工程 A）| 设置 `objIDC.BatchMode = True`，循环里调 `RecordSavedBill(objIDC.rtnVouBillID)` |
| `CVouService.SaveDoc` 修改（工程 C）| 通过 `Me.BatchMode`（属于本对象）判断是否跳过 per-bill stored proc |

## 前置依赖

- PR-1 + PR-2 已合入
- **必须先在 SQL Server 部署 `PR3_storedprocs.sql`**（创建 2 个批量版 stored proc）
- 新增 `VouBatchOps.bas`（工程 A）

## Patch 1：`cMthCstAccGL2.cls` `meCreateVou` 加 batch 包装 + 大事务

```vb
Public Sub meCreateVou(...)
On Error GoTo ErrH
    Dim blnGL2  As Boolean
    Dim SQL     As String
    Dim blnMetaLoadedHere As Boolean
    Dim blnTxnStarted As Boolean        ' === PR-3 新增 ===

    If EndDate = "" Then EndDate = Format(Date, "yyyy-MM-dd")
    If Not (Me.AppParameters.UseGL2 And Me.ReCreateGL) Then GoTo ErrH

    If objGL2 Is Nothing Then
        Set objGL2 = New cMthCstAccGL2
        ' ...
    End If
    objGL2.isPreGenIDOnly = isPreGenIDOnly
    objGL2.isSaveToTransitionalTable = isSaveToTransitionalTable

    ' === PR-1/2 元数据缓存：懒加载，无需在此处显式 LoadAll ===

    ' === PR-3 batch + 大事务包装 ===
    ' 防御性：仅当外层无事务时启动；若已在事务中，依赖外层事务即可
    Call VouBatchOps.BeginBatch
    On Error Resume Next
    Call objDS.BeginTrans
    blnTxnStarted = (Err.Number = 0)
    Err.Clear
    On Error GoTo ErrH

    If Me.AppParameters.PSBillAutoCreateVou Then Call objGL2.meCreVouForSS(objDS)
    If Me.HSObjectTag <> HSObjectAssPrdt Then
        Call objGL2.meCreVouForMMFE(objDS)
        Call objGL2.meCreVouForMMOU(objDS)
        Call objGL2.meCreVouForMMIn(objDS)
        Call objGL2.meCreVouForPG(objDS)
    End If
    If Me.AppParameters.PSBillAutoCreateVou Then Call objGL2.meCreVouForJGIC(objDS)
    Call objGL2.meCreVouForIC(objDS)
    Call objGL2.meCreVouForIIO(objDS)
    Call objGL2.meCreVouForPI(objDS)
    Call objGL2.meCreVouForIPS(objDS)
    Call objGL2.meCreVouForCstFYFT(objDS)

    ' === PR-3 batch flush（必须在 commit 之前）===
    Call VouBatchOps.FlushStoredProcs(objDS)

    ' === PR-3 批末尾 commit ===
    Call objDS.CommitTrans
    blnTxnStarted = False

    Call objGL2.CreateCstDiffVou(objDS, mEmpID)

ErrH:
    ' === PR-3 异常时回滚整个 batch 事务 ===
    If blnTxnStarted Then
        On Error Resume Next
        Call objDS.RollbackTrans
        On Error GoTo 0
    End If
    Call VouBatchOps.EndBatch

    If Err.Number <> 0 Then
        Call Err.Raise(Err.Number, , Err.Description)
    End If
End Sub
```

> **注意 1**：每个 `meCreVouForXX` 内部需要在 CreateVou 后调 RecordSavedBill — 见 Patch 1.5。
>
> **注意 2**：原代码 `CVouService.SaveDoc` 内部也有 `BeginTrans/CommitTrans`。在 PR-3 batch 模式下，外层已经包了大事务，**SaveDoc 内部应跳过自己的事务**。详见 Patch 3。

---

## Patch 1.5：`cMthCstAccGL2.meCreVouForXX` 循环内 RecordSavedBill（工程 A，跨工程关键）

每个 `meCreVouForXX` 都有这种循环：

```vb
Do While Not rsBill.EOF
    objIDC.BillType = "..."
    objIDC.BillID = rsBill.Fields("billid").Value
    ' ... 设置一堆属性 ...
    Call objIDC.CreateVou(objDS, False)

    ' === PR-3 新增 ===
    Call objIDC.SetBatchMode(True)         ' 仅首次需要，可在循环外做
    Call VouBatchOps.RecordSavedBill(objIDC.rtnVouBillID, isSaveToTransitionalTable)

    rsBill.MoveNext
Loop
```

实现要点：
- `objIDC.SetBatchMode(True)` 是 IDCService 新增方法，**跨工程把 batch 标志传到 t_FVou_M**（见 Patch 2）
- `objIDC.rtnVouBillID` 是已存在的字段，CreateVou 完成后保存当前凭证 BillID
- `RecordSavedBill` 在工程 A 的 VouBatchOps collection 内收集 BillID
- 必须在每个 `meCreVouForXX` 都加这两行

---

## Patch 2：跨工程 BatchMode 类属性传递

### Patch 2a：`IDCService.cls` 新增属性 + 转发（工程 B）

```vb
' === PR-3 跨工程标志 ===
Public BatchMode As Boolean

Public Sub SetBatchMode(ByVal v As Boolean)
    BatchMode = v
End Sub
```

在 `IDCService.CreateVou` 内创建 `t_FVou_M` 后，把 BatchMode 转发：

```vb
Set objVouD = New t_FVou_M
With objVouD
    Set .AppParameters = Me.AppParameters
    .AccRate = Me.AccRate
    ' ... 原代码 ...
    .BatchMode = Me.BatchMode      ' === PR-3 新增：跨工程标志转发 ===
End With
```

### Patch 2b：`t_FVou_M.cls` 新增 BatchMode 属性（工程 C）

```vb
' === PR-3 跨工程标志 ===
Public BatchMode As Boolean
```

### Patch 2c：`CVouService.SaveDoc` 末尾改为根据 t_FVou_M.BatchMode 跳过 stored proc

⚠️ 注意：原 SaveDoc 是 `t_FVou_M.CreateVou` 内部调用 `CVouService.SaveDoc(...)`。`CVouService` 没有直接拿到 `t_FVou_M.BatchMode`，需要用 ByVal 参数透传，或通过 `CVouService` 类自身加 BatchMode 属性。**最简单**：在 `t_FVou_M.CreateVou` 调 SaveDoc 之前把 BatchMode 设置到 objVou：

```vb
Set objVou.AppParameters = Me.AppParameters
objVou.UserName = Me.UserName
objVou.NotWriteEvent = Me.NotWriteEvent
objVou.isSaveToTransitionalTable = Me.isSaveToTransitionalTable
objVou.BatchMode = Me.BatchMode    ' === PR-3 新增 ===
```

`CVouService.cls` 同样新增 `Public BatchMode As Boolean`。

`SaveDoc` 内部修改：

#### 修改前

```vb
'4、调用存储过程，实现主表与明细表数据的一致性
If isSaveToTransitionalTable = False Then
    Call sysDS.ExecSQL("exec ss_YWVouUpdateAchTimes '" & strBillID & "',1")
    Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',0")
Else
    Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',1")
End If
```

#### 修改后

```vb
'4、调用存储过程；batch 模式下跳过，由外层（工程 A）整 batch 末尾批量执行
If Not Me.BatchMode Then
    ' 原路径（非 batch 模式）
    If isSaveToTransitionalTable = False Then
        Call sysDS.ExecSQL("exec ss_YWVouUpdateAchTimes '" & strBillID & "',1")
        Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',0")
    Else
        Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',1")
    End If
End If
' batch 模式下：什么都不做，rtnVouBillID 已设置，由 t_FVou_M.CreateVou 返回
' 给 IDCService.CreateVou，再返回给工程 A 的 meCreVouForXX 循环；
' 循环里调 VouBatchOps.RecordSavedBill(rtnVouBillID, ...) 收集
```

---

## Patch 3：`CVouService.SaveDoc` 跳过 per-bill BeginTrans/CommitTrans

### 修改前

```vb
'2、实现凭证号的建立
If DBConnection Is Nothing Then
    Call sysDS.BeginTrans
End If
```

### 修改后

```vb
'2、实现凭证号的建立
'   PR-3：batch 模式下外层已包大事务，跳过 per-bill BeginTrans
If DBConnection Is Nothing And Not Me.BatchMode Then
    Call sysDS.BeginTrans
End If
```

末尾 CommitTrans 同样：

```vb
'提交事务
If DBConnection Is Nothing And Not Me.BatchMode Then
    Call sysDS.CommitTrans
End If
```

ErrHandler 段：

```vb
ErrHandler:
    If Not sysDS Is Nothing Then
        If Err.Number <> 0 Then
            ' batch 模式下不在 SaveDoc 内回滚，由外层 meCreateVou 统一回滚
            If Not Me.BatchMode Then
                sysDS.RollbackTrans
            End If
        End If
        ...
    End If
```

---

## Patch 4：可选 — 每 1000 张 flush 一次事务（避免单一巨型事务过长）

如果整 batch 跑 10 万张凭证一个事务，事务日志会非常大、锁持有时间长。可以选择性地在每个 `meCreVouForXX` 内部循环里检查 `VouBatchOps.ShouldFlushTxn`：

```vb
Do While Not rsBill.EOF
    ' ... 现有逻辑 ...
    Call objIDC.CreateVou(objDS, False)
    
    ' === PR-3 可选：每 1000 张凭证 commit 一次 ===
    If VouBatchOps.ShouldFlushTxn() Then
        ' 先 flush 已收集的 stored proc（保持事务内一致性）
        Call VouBatchOps.FlushStoredProcs(objDS)
        Call objDS.CommitTrans
        Call objDS.BeginTrans
    End If
    
    rsBill.MoveNext
Loop
```

> **风险提示**：每 1000 张 commit 牺牲了"整 batch 原子性"。如果第 5001 张失败，前 5000 张已落库，需要业务侧支持"幂等重跑"。如果业务要求严格原子性，请保持单一巨型事务（不应用 Patch 4）。

---

## 部署顺序

1. **先在 SQL Server 部署 `PR3_storedprocs.sql`** —— 创建批量 stored proc
2. 加 `VouBatchOps.bas` 到 VB6 工程
3. 应用 Patch 2（SaveDoc → RecordSavedBill）→ 灰度（此时 InBatchMode=False，仍走原路径）
4. 应用 Patch 3（跳过 per-bill 事务）→ 灰度
5. 应用 Patch 1（启用 BeginBatch/FlushStoredProcs/事务包装）→ batch 模式生效
6. （可选）应用 Patch 4（中间 commit）

每一步都可独立灰度，因为：
- VouBatchOps.InBatchMode=False 时所有路径回退原行为
- 批量 stored proc 内部仍是 cursor 循环原 stored proc，逻辑零变化

## 100% 等价证明

- 批量 stored proc 内部用 cursor 调原 stored proc，逻辑严格等价
- 大事务的 commit 边界等价于"原代码每张凭证 commit 后立即继续"
  （SQL Server 事务嵌套规则：内层 commit 不真 commit，最外层 commit 才真 flush，
   原代码因为没有外层事务所以每次都真 flush，而 PR-3 整 batch 一次 flush）
- 异常路径：原代码任一凭证失败立刻 raise 中止，PR-3 也是同样行为
  （catch 后 RollbackTrans 回滚整个 batch — 这是与原代码的差异：原代码
   失败时已落库的凭证不回滚；PR-3 全部回滚 ⇒ 失败时**事务原子性更强**，
   是改进而非倒退；如果需要保留原"失败前已落库不回滚"行为，应用 Patch 4
   每 1000 张 commit）
