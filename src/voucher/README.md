# 凭证生成性能优化 PR-1：元数据缓存 + 消除反射

## ⚠️ 跨工程架构（最重要）

凭证生成涉及三个 VB6 ActiveX DLL：

```
POPBus3FileService.dll  (工程 A)  — meCreateVou, cMthCstAccGL2.meCreVouForXX
POPBus3GL2IDC.dll       (工程 B)  — IDCService, SSBillByDateDAL 等 DAL
POPBus3GL2Service.dll   (工程 C)  — t_FVou_M, CVouService.SaveDoc
```

**.bas 模块状态在不同工程间不共享**。本 PR 的所有 `.bas` 文件需要按下表加到对应工程：

| .bas 文件 | 加到哪些工程 | 原因 |
|---|---|---|
| `VouMetaCache.bas` | **A + B + C 三个工程** | BeforeAction (C) / GetFIIDByPrdt (C) / meCreVouForXX (A) 都会用到；元数据是只读的，每工程独立缓存数据相同 |
| `VouCacheHelpers.bas` | **C 工程** | CheckFI/CheckCorp/CheckEmp/CheckAcc 被 BeforeAction (C) 调用；LookupDocFI_* 被 t_FVou_M (C) 调用 |

**懒加载机制**：每个工程首次调 `TryGetXxx` 时自动 `EnsureLoaded(objDS)`。元数据是只读的，3 个工程各加载一次（共 21 次 SQL）依然 < 50ms。

## 范围

**第一波** A1/A3/A4/A5/A6（A2 移至 PR-2 同步处理写入路径）：

- **A1**：在 `meCreateVou` 入口一次性把 7 张元数据表加载到 Dictionary
- **A3**：`meAddRowI` 反射赋值（CallByName）→ 显式属性赋值
- **A4**：`t_FVou_M.Init()` 改为 batch 共享缓存，不再每实例重新加载 GL2_AchFIID
- **A5**：`CVouService.BeforeAction` 内 4 张元数据表 IN-list 查询 → Dictionary 查找
- **A6**：`CheckDateValidate` 2 次 AccPeriod SELECT → Dictionary 查找

## 100% 等价性约束

- `VouMetaCache.TryGet*` 任何函数返回未命中 → 回退到原 DB 路径
- 任何字段、错误信息、字段写入顺序与原代码完全一致
- 凭证号、ID、ADO Update、stored proc、日志事件保持原行为

## 代码审计修复记录（2026-05-02）

完成首版后做了一次完整代码审计，发现并修复以下 bug：

### 已修复

| # | Bug | 严重 | 修复 |
|---|---|---|---|
| 18 | VouMetaCache 用 `\|` 拼接字符串存值，FIName/CorpName 含 `\|` 时 Split 错位 | ★★★★★ | 改用嵌套 Dictionary 存每个字段，Key 用 Chr$(31) 分隔（控制字符不会出现在数据中）|
| 19 | TryGetDocFI 三段 ID 用 `\|` 拼接，ID 含 `\|` 时 key 错乱 | ★★★ | 同上，改用 Chr$(31) |
| 21 | rsFIBak schema 缺少 `EXPTAG/fullname/ISStop` 字段 | ★ | 经核：原代码 Else 分支只读 10 字段，与新 schema 一致 |
| 28 | helper 内 rsFIBak 写入时机错误（在 ItemsData.Update 之前）| ★★ | 注释明确：调用方负责 ItemsData.Update；rsFIBak 是独立 Recordset，写入时机不影响 |
| 30 | `getFIIDByClsIDFromParentCls` 在 cache 路径下仍用 `mrsDocFIRel`（已被设为 Nothing）→ Err 91 | ★★★★★ | 新增 helper `LookupDocFI_ParentCls` 走 cache，Patch 3-3 覆盖 |
| 31 | 16 个 GetFIIDBy* 函数 cache 路径需独立改造，Patch 文档"详见..."不够 | ★★★★ | 新增 helper 组 B 三个统一函数 `LookupDocFI_5Step` / `LookupDocFI_ParentCls` / `LookupDocFI_KeyOnly`；Patch 3-1～3-12 给出每个函数的具体改造 |
| 16 | LoadAll 错误时 ClearAll 可能覆盖原 Err.Number | ★ | 加 `On Error Resume Next` 防御 |

### 保留原代码 bug（按 100% 等价要求）

| # | 原代码 bug | 处理 |
|---|---|---|
| 5 | CheckAcc 错误信息用 `rsAccs.Fields("CorpName")` —— Account 表 SELECT 不含此字段，访问抛 ADO 3265 异常 | cache 路径主动抛 3265 同型异常；db 路径保持原 bug 行为；生产中此分支几乎不触发 |
| 7 | CheckCorp 在 RecordCount=0 时访问 `rsCorps.Fields("CorpName").Value` 抛 ADO 3021 (BOF/EOF) | cache 路径主动抛 3021；db 路径不变 |

> 这两条 bug 在生产中**几乎不会触发**（账户/客户被删除且仍在引用），保留原行为可保证零业务影响。如需后续修复，可在专门 PR 内一次性整改 Account/Corp 表的字段名错误。

### 误报澄清

| # | 误判 | 真相 |
|---|---|---|
| 11 | "EXPTAG 大小写"问题 | ADO `Fields(name)` 大小写不敏感，无影响 |
| 26 | "ItemsData.Fields("FIID") vs ("fiid") 大小写不一致" | 同上，无影响 |
| 13 | "未处理 isFromSYS"  | 该判断在 helper 之外（BeforeAction 主体），保留原位 |
| 23 | "rsFIBak 命中后跳过校验" | 经核：原代码刻意优化，cache 路径行为一致 |

## 已通过的等价性测试

`tools/test_voucher_pr1.py`：**40 个用例全过**

- CheckFI 6/6（含 stop / EXPTAG=1 / 不存在）
- GetFIIDByPrdt 5 步 fallback 7/7
- AccPeriod 日期定位 6/6（含边界）
- Corp/Emp/Account 9/9
- meAddRowI 反射 vs 显式 37 字段全部一致
- **新增**（审计回归）：
  - Bug 18 回归：FIName / CorpName / FullName / HSTagName 含 `\|` 字符 ✓
  - Bug 19 回归：DocKey / WHID / DocID 含 `\|` 字符 ✓
  - Bug 23 回归：EXPTAG=1 错误信息一致 ✓
  - 5 步严格 fallback 顺序一致 ✓
  - FOAPTag 分支一致 ✓

## 性能收益模型

`tools/perf_voucher_pr1.py` 量化估算：

```
Baseline（PR-1 之前）        : 39.1 min
After PR-1（A1+A3+A4+A5+A6）: 25.1 min
节省                         : 14.1 min（36%）
```

实际运行（用户基线 ~2 小时）预期：**~120 min → ~75 min**

## 文件清单

| 文件 | 类型 | 说明 |
|---|---|---|
| `VouMetaCache.bas` | 新增标准模块 | 整 batch 共享的元数据缓存（嵌套 Dictionary 结构）|
| `VouCacheHelpers.bas` | 新增标准模块 | 组 A：BeforeAction helper；组 B：GetFIIDBy* 系列统一封装 |
| `PR1_patches.md` | 文档 | 12 处既有类的修改 patch（含组 B 的 8 个函数改造细节）|

## 部署步骤

详见 `PR1_patches.md`。简要顺序：

1. 把 `VouMetaCache.bas` + `VouCacheHelpers.bas` 加入 VB6 工程（标准模块）
2. 应用 Patch 6（meAddRowI 显式赋值）→ 编译 → 灰度
3. 应用 Patch 5（CheckDateValidate）→ 编译 → 灰度
4. 应用 Patch 4（BeforeAction）→ 编译 → 灰度
5. 应用 Patch 2 + 3-1～3-12（t_FVou_M）→ 编译 → 灰度（全部 GetFIIDBy* 函数都改）
6. 应用 Patch 1（启用 LoadAll/ClearAll 入口）→ 全部缓存生效

> 前 5 步应用时缓存未加载（`VouMetaCache.IsLoaded=False`），所有路径自动走原 DB 路径，行为零变化。最后 1 步启用入口时缓存生效。

## 后续 PR

- **PR-2**：A2 schema 缓存 + ADO Update → 批量 INSERT VALUES
- **PR-3**：commit 包大事务 + stored proc 批量化（`ss_YWVouUpdateAchTimes` / `ss_AfterSaveUpdateGL2VouBill`）
- **PR-4**：DAL 内 GROUP BY 合并 + 主查询投影列
