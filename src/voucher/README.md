# 凭证生成性能优化 PR-1：元数据缓存 + 消除反射

## 范围

**第一波** A1/A3/A4/A5/A6（A2 移至 PR-2 同步处理写入路径）：

- **A1**：在 `meCreateVou` 入口一次性把 7 张元数据表加载到 Dictionary
- **A3**：`meAddRowI` 反射赋值（CallByName）→ 显式属性赋值
- **A4**：`t_FVou_M.Init()` 改为 batch 共享缓存，不再每实例重新加载 GL2_AchFIID
- **A5**：`CVouService.BeforeAction` 内 4 张元数据表 IN-list 查询 → Dictionary 查找
- **A6**：`CheckDateValidate` 2 次 AccPeriod SELECT → Dictionary 查找

## 等价性约束（100%）

- `VouMetaCache.TryGet*` 任何函数返回未命中 → 回退到原 DB 路径
- 任何字段、错误信息、字段写入顺序与原代码完全一致
- 凭证号、ID、ADO Update、stored proc、日志事件保持原行为

## 已通过的等价性测试

- `tools/test_voucher_pr1.py`：35 个用例
  - CheckFI 6/6（含 stop / EXPTAG=1 / 不存在）
  - GetFIIDByPrdt 5 步 fallback 7/7
  - AccPeriod 日期定位 6/6（含边界）
  - Corp/Emp/Account 9/9
  - meAddRowI 反射 vs 显式 37 字段全部一致

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
| `VouMetaCache.bas` | 新增标准模块 | 整 batch 共享的元数据缓存 |
| `VouCacheHelpers.bas` | 新增标准模块 | BeforeAction 等价回退 helper |
| `PR1_patches.md` | 文档 | 6 处既有 .cls 文件的修改 patch |

## 部署步骤

详见 `PR1_patches.md`。简要顺序：

1. 把 `VouMetaCache.bas` + `VouCacheHelpers.bas` 加入 VB6 工程（标准模块）
2. 应用 Patch 6（meAddRowI 显式赋值）→ 编译 → 灰度测试
3. 应用 Patch 5（CheckDateValidate）→ 编译 → 灰度
4. 应用 Patch 4（BeforeAction）→ 编译 → 灰度
5. 应用 Patch 2 + 3（t_FVou_M）→ 编译 → 灰度
6. 应用 Patch 1（启用 LoadAll/ClearAll 入口）→ 全部缓存生效

> 之所以最后才应用 Patch 1：前几个 Patch 在缓存未加载时都自动回退到原 DB 路径，逐步部署时每一步都安全。

## 后续 PR

- **PR-2**：A2 schema 缓存 + ADO Update → 批量 INSERT VALUES
- **PR-3**：commit 包大事务 + stored proc 批量化（`ss_YWVouUpdateAchTimes` / `ss_AfterSaveUpdateGL2VouBill`）
- **PR-4**：DAL 内 GROUP BY 合并 + 主查询投影列
