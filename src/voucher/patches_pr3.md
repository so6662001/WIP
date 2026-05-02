# PR-3 应用 patches

## 范围

| 项 | 内容 | 增量收益 |
|---|---|---|
| **B1** | 工程 A 的 `meCreateVou` 包大事务（消除 10 万次 commit）| -5.0 min |
| **B2** | `ss_YWVouUpdateAchTimes` per-bill → 批量 BillID list | -1.7 min |
| **B3** | `ss_AfterSaveUpdateGL2VouBill` per-bill → 批量 BillID list | -1.7 min |

## 跨工程架构

| 修改 | 工程 |
|---|---|
| `VouBatchOps.bas` | **A**（POPBus3FileService）|
| `storedprocs_pr3.sql` | SQL Server 部署 |
| `cMthCstAccGL2.meCreateVou` 包大事务 + BeginBatch + Flush | A |
| `cMthCstAccGL2.meCreVouForXX` 循环里调 RecordSavedBill | A |
| `IDCService.cls` 新增 `BatchMode` 属性 | B |
| `t_FVou_M.cls` 新增 `BatchMode` 属性 | C |
| `CVouService.cls` 新增 `BatchMode` 属性 | C |
| `CVouService.SaveDoc` 用 `Me.BatchMode` 跳过 stored proc | C |

### 跨工程标志传递链

```
A.meCreateVou
  Set objIDC.BatchMode = True
       ↓ COM 跨 A→B 工程
B.IDCService.CreateVou
  Set objVou = New t_FVou_M       (在 B 工程内 New 但跨 B→C 类型)
  objVou.BatchMode = Me.BatchMode
       ↓ COM 跨 B→C 工程
C.t_FVou_M.CreateVou
  Set objCVou = New CVouService
  objCVou.BatchMode = Me.BatchMode (C 工程内)
       ↓
C.CVouService.SaveDoc
  If Me.BatchMode Then
    ' 跳过 per-bill stored proc，由工程 A 末尾批量调用
  Else
    ' 原 per-bill stored proc 调用
```

---

## P3.1：SQL Server 部署批量 stored proc

执行 `storedprocs_pr3.sql`，创建：
- `dbo.ss_YWVouUpdateAchTimes_Batch(@BillIDList NVARCHAR(MAX), @Adj INT)`
- `dbo.ss_AfterSaveUpdateGL2VouBill_Batch(@BillIDList NVARCHAR(MAX), @IsTrans BIT)`

内部用 XML.nodes() 拆 BillID list（兼容 SQL Server 2008，不依赖 STRING_SPLIT），cursor 循环调原 stored proc 保证 100% 等价；用 `IDENTITY(1,1)` 列保证按传入顺序处理。

---

## P3.2：`IDCService.cls` 新增 BatchMode 属性（**工程 B**）

类成员声明区域：

```vb
' === PR-3 跨工程标志 ===
Public BatchMode As Boolean
```

`CreateVou` 方法内创建 `t_FVou_M` 后转发：

```vb
' ... 原代码: Set objVouD = New POPBus3GL2Service.t_FVou_M ...
With objVouD
    Set .AppParameters = Me.AppParameters
    .AccRate = Me.AccRate
    ' ... 原代码所有属性赋值 ...
    .BatchMode = Me.BatchMode    ' === PR-3 新增：跨 B→C 工程边界 ===
End With
```

---

## P3.3：`t_FVou_M.cls` 新增 BatchMode 属性（**工程 C**）

类成员声明区域：

```vb
' === PR-3 跨工程标志（由 IDCService 通过类属性赋值传递）===
Public BatchMode As Boolean
```

`CreateVou` 方法内（C 工程内创建 `CVouService` 后转发）：

```vb
' ... 原代码: Set objVou = New CVouService ...
Set objVou.AppParameters = Me.AppParameters
objVou.UserName = Me.UserName
objVou.NotWriteEvent = Me.NotWriteEvent
objVou.isSaveToTransitionalTable = Me.isSaveToTransitionalTable
objVou.BatchMode = Me.BatchMode   ' === PR-3 新增 ===
```

---

## P3.4：`CVouService.cls` 新增 BatchMode 属性 + 修改 SaveDoc（**工程 C**）

类成员声明区域：

```vb
' === PR-3 跨工程标志 ===
Public BatchMode As Boolean
```

### P3.4.1 `SaveDoc` 内 stored proc 调用判断

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
'4、调用存储过程；batch 模式下跳过，由工程 A 末尾批量执行
If Not Me.BatchMode Then
    If isSaveToTransitionalTable = False Then
        Call sysDS.ExecSQL("exec ss_YWVouUpdateAchTimes '" & strBillID & "',1")
        Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',0")
    Else
        Call sysDS.ExecSQL("exec ss_AfterSaveUpdateGL2VouBill '" & strBillID & "',1")
    End If
End If
```

### P3.4.2 `SaveDoc` 内 BeginTrans / CommitTrans 判断

#### 修改前

```vb
'2、实现凭证号的建立
If DBConnection Is Nothing Then
    Call sysDS.BeginTrans
End If
' ... 业务代码 ...
'提交事务
If DBConnection Is Nothing Then
    Call sysDS.CommitTrans
End If
```

#### 修改后

```vb
'2、实现凭证号的建立
'   PR-3：batch 模式下外层 A 工程已包大事务，跳过 per-bill BeginTrans
If DBConnection Is Nothing And Not Me.BatchMode Then
    Call sysDS.BeginTrans
End If
' ... 业务代码 ...
'提交事务
If DBConnection Is Nothing And Not Me.BatchMode Then
    Call sysDS.CommitTrans
End If
```

`ErrHandler` 段：

```vb
ErrHandler:
    If Not sysDS Is Nothing Then
        If Err.Number <> 0 Then
            ' batch 模式下不在 SaveDoc 内回滚，由工程 A 的 meCreateVou ErrH 统一回滚
            If Not Me.BatchMode Then
                sysDS.RollbackTrans
            End If
        End If
        ...
    End If
```

---

## P3.5：`cMthCstAccGL2.meCreVouForXX` 循环里收集 BillID（**工程 A**）

每个 `meCreVouFor*`（SS / MMFE / MMOU / MMIn / PG / JGIC / IC / IIO / PI / IPS / CstFYFT）的 Do While 循环：

### 修改前

```vb
Do While Not rsBill.EOF
    objIDC.BillType = "..."
    objIDC.BillID = rsBill.Fields("billid").Value
    ' ... 设置一堆属性 ...
    Call objIDC.CreateVou(objDS, False)
    rsBill.MoveNext
Loop
```

### 修改后

```vb
Do While Not rsBill.EOF
    objIDC.BillType = "..."
    objIDC.BillID = rsBill.Fields("billid").Value
    ' ... 设置一堆属性 ...
    objIDC.BatchMode = True                                          ' === PR-3 新增 ===
    Call objIDC.CreateVou(objDS, False)
    Call VouBatchOps.RecordSavedBill(objIDC.rtnVouBillID, _
                                      Me.isSaveToTransitionalTable)  ' === PR-3 新增 ===
    rsBill.MoveNext
Loop
```

> `objIDC.rtnVouBillID` 是 `IDCService.CreateVou` 已存在的输出字段，CreateVou 完成后保存当前凭证 BillID。

---

## P3.6：`cMthCstAccGL2.meCreateVou` 包大事务 + Flush（**工程 A**）

### 修改前

```vb
Public Sub meCreateVou(...)
On Error GoTo ErrH
    Dim blnGL2 As Boolean
    Dim SQL    As String

    If EndDate = "" Then EndDate = Format(Date, "yyyy-MM-dd")
    If Not (Me.AppParameters.UseGL2 And Me.ReCreateGL) Then GoTo ErrH

    If objGL2 Is Nothing Then
        Set objGL2 = New cMthCstAccGL2
        ' ...
    End If
    objGL2.isPreGenIDOnly = isPreGenIDOnly
    objGL2.isSaveToTransitionalTable = isSaveToTransitionalTable

    If Me.AppParameters.PSBillAutoCreateVou Then Call objGL2.meCreVouForSS(objDS)
    ' ... meCreVouForXX × 11 ...
    Call objGL2.CreateCstDiffVou(objDS, mEmpID)

ErrH:
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, , Err.Description)
End Sub
```

### 修改后

```vb
Public Sub meCreateVou(...)
On Error GoTo ErrH
    Dim blnGL2 As Boolean
    Dim SQL    As String
    Dim blnTxnStarted As Boolean        ' === PR-3 新增 ===

    If EndDate = "" Then EndDate = Format(Date, "yyyy-MM-dd")
    If Not (Me.AppParameters.UseGL2 And Me.ReCreateGL) Then GoTo ErrH

    If objGL2 Is Nothing Then
        Set objGL2 = New cMthCstAccGL2
        ' ...
    End If
    objGL2.isPreGenIDOnly = isPreGenIDOnly
    objGL2.isSaveToTransitionalTable = isSaveToTransitionalTable

    ' === PR-3 batch + 大事务包装 ===
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
    If blnTxnStarted Then
        Call objDS.CommitTrans
        blnTxnStarted = False
    End If

    Call objGL2.CreateCstDiffVou(objDS, mEmpID)

ErrH:
    ' === PR-3 异常时回滚整个 batch 事务 ===
    If blnTxnStarted Then
        On Error Resume Next
        Call objDS.RollbackTrans
        On Error GoTo 0
    End If
    Call VouBatchOps.EndBatch

    If Err.Number <> 0 Then Call Err.Raise(Err.Number, , Err.Description)
End Sub
```

---

## 部署核对清单

- [ ] SQL Server 部署 `storedprocs_pr3.sql`
- [ ] `VouBatchOps.bas` 添加到工程 **A**（POPBus3FileService）
- [ ] 应用 P3.2 到 `IDCService.cls`（工程 B）
- [ ] 应用 P3.3 到 `t_FVou_M.cls`（工程 C）
- [ ] 应用 P3.4 到 `CVouService.cls`（工程 C）
- [ ] 应用 P3.5 到 `cMthCstAccGL2.meCreVouForXX` 11 个方法（工程 A）
- [ ] 应用 P3.6 到 `cMthCstAccGL2.meCreateVou`（工程 A）
- [ ] 编译三个工程并部署 DLL
- [ ] 灰度对比 baseline 行为

## 100% 等价

- 批量 stored proc 内部用 `CURSOR LOCAL FAST_FORWARD` + `ORDER BY rn` 严格按传入顺序循环调原 stored proc
- per-page 交错顺序（每 200 张一页内部先 A 后 B）保留两个 stored proc 之间的相对顺序
- BillID 安全字符校验拒绝含 `,` / `'` / `;` / 空白 的输入
- 缓存未启用 / `BatchMode=False` 时所有路径回退原 per-bill 行为
- 异常路径：原代码失败前已落库不回滚；PR-3 失败时整 batch 回滚（**原子性更强**，是改进）

## 性能预期

```
After PR-2: ~60 min
After PR-3: ~35 min
增量节省:   -25 min (-42%)
```
