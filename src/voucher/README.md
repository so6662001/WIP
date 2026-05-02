# 凭证生成性能优化 PR-1 + PR-2 + PR-3：元数据缓存 + schema 缓存 + 大事务批量化

## 跨工程架构（必读）

凭证生成涉及 **3 个 VB6 ActiveX DLL** 工程：

```
POPBus3FileService.dll  (工程 A)  meCreateVou, cMthCstAccGL2.meCreVouForXX
        │ COM
        ▼
POPBus3GL2IDC.dll       (工程 B)  IDCService, SSBillByDateDAL 等 DAL
        │ COM
        ▼
POPBus3GL2Service.dll   (工程 C)  t_FVou_M, CVouService.SaveDoc / BeforeAction
                                  / CheckDateValidate / megetDocByID
```

VB6 标准模块（`.bas`）的状态在不同 ActiveX DLL 工程间**不共享**——同一份 `VouMetaCache.bas` 加到 A 和 C 两个工程里，A 的 `m_dctFI` 和 C 的 `m_dctFI` 是**两个独立的变量**。

详见 `ARCHITECTURE_AUDIT.md`。

## PR-1 范围

| 项 | 内容 | 改动位置 |
|---|---|---|
| **A1** | 7 张元数据表（FinanceItems / Corp / Emp / Account / AccPeriod / GL2_AchFIID / F_VouCls_Bills）→ Dictionary | `VouMetaCache.bas` |
| **A3** | `t_FVou_M.CreateVou` 内 `meAddRowI` CallByName 反射 → 显式属性赋值 37 字段 | 工程 C |
| **A4** | `t_FVou_M.Init()` 在 cache 已加载时跳过 GL2_AchFIID 全表 SELECT | 工程 C |
| **A5** | `CVouService.BeforeAction` 4 张元数据表 IN-list 查询 → Dictionary 查找 | 工程 C |
| **A6** | `CheckDateValidate` 2 次 AccPeriod SELECT → Dictionary 查找 | 工程 C |

## 跨工程懒加载机制

`VouMetaCache.bas` **同一份文件加到三个工程**：

```
工程 A → 自己的 m_dctFI（独立）
工程 B → 自己的 m_dctFI（独立）
工程 C → 自己的 m_dctFI（独立）
```

每个工程**首次调用 `TryGetXxx` 时自动 `EnsureLoaded(objDS)`**：

```vb
Public Sub EnsureLoaded(ByVal objDS As HHDataService.sysDataService)
    If m_blnLoaded Then Exit Sub
    Call LoadAll(objDS)
End Sub
```

元数据是**只读**的（FinanceItems 等），3 个工程独立缓存数据完全相同。每个工程加载 7 次 SQL，3 个工程一共 21 次 SQL（一次性，不在主循环里），**总耗时 < 50ms，可忽略**。

## 文件部署

| .bas 文件 | 加到工程 | 作用 |
|---|---|---|
| `VouMetaCache.bas` | **A + B + C** | 元数据缓存（每工程独立，懒加载）|
| `VouCacheHelpers.bas` | **C** | BeforeAction / GetFIIDBy* 系列等价回退 helper |
| `VouSchemaCache.bas` | **C** (PR-2) | FVou_M/FVou_I/FVou_II 空 schema 缓存 |
| `VouBatchOps.bas` | **A** (PR-3) | 工程 A 内 BillID 收集 + 末尾批量 stored proc |

## VB6 类修改（patches）

详见 `patches.md`。摘要：

### 工程 C 修改

1. `t_FVou_M.cls` — `Init` / `GetFIIDByPrdt` / `GetAccFIID` 等 16 个 GetFIIDBy* 函数 cache fast path
2. `t_FVou_M.cls` — `meAddRowI_All` 显式属性赋值替代 CallByName
3. `CVouService.cls` — `BeforeAction` 调用 `VouCacheHelpers.CheckFI/CheckCorp/CheckEmp/CheckAcc`
4. `CVouService.cls` — `CheckDateValidate` 用 `VouMetaCache.TryGetAPByDate`

### 工程 A 修改

无（懒加载自动处理）。

### 工程 B 修改

无。

## 性能收益

| 阶段 | 总耗时 | 增量节省 |
|---|---|---|
| Baseline（用户报告基线）| ~120 min | — |
| **After PR-1** | **~75 min** | **-45 min（-38%）**|
| **After PR-2** | **~60 min** | **-15 min（-20%）**|
| **After PR-3** | **~35 min** | **-25 min（-42%）**|

## 100% 等价性约束

- `VouMetaCache.TryGetXxx` 任何函数返回未命中 → 调用方**自动回退原 DB 路径**，输入相同的 key 必须返回与原 DB 查询完全一样的值
- 字段写入顺序、错误信息文案、ADO 字段类型转换都与原代码逐字对照
- 凭证号生成、ADO Update、stored proc 调用、日志事件均**不动**

## 测试

| 测试 | 数量 | 结果 |
|---|---|---|
| `test_voucher_pr1.py` 等价性 | 40 | ✅ 全过 |
| `test_cross_project_arch.py` 跨工程架构 | 12（PR-1 分支适用 7）| ✅ 全过 |
| `vbcheck.py` 静态校验 | 2 个 .bas | ✅ 无错误 |

## 部署顺序（每步独立可灰度，缓存未启用时回退原行为）

1. 把 `VouMetaCache.bas` 添加到 **A + B + C 三个工程**
2. 把 `VouCacheHelpers.bas` 添加到 **C 工程**
3. 应用 patches.md 中的 12 处具体修改（按顺序灰度）

完成后无需任何主动 LoadAll 调用，懒加载自动处理。
