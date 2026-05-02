"""
PR-1 性能模型：量化预期收益。

模型参数（基于用户描述："每条 SQL 都很快、调用次数太多、每张凭证 ~10 次 DB 访问"）:
  - 单次 SQL round-trip = 1.0 ms
  - 单次 ADO Update    = 1.0 ms
  - 单次 CallByName    = 8 us (VB6 IDispatch::Invoke)
  - 单次 Dictionary 查找 = 0.5 us
  - 单次内存属性赋值   = 0.1 us

凭证规模：10 万张，平均每张 5 行明细
"""
from __future__ import annotations


SQL_RPC_MS         = 1.0        # 每次 DB round-trip
ADO_UPDATE_MS      = 1.0        # 每行 ADO Update
CALLBYNAME_US      = 8.0        # 反射调用
DICT_LOOKUP_US     = 0.5        # 缓存查找
EXPLICIT_ASSIGN_US = 0.1        # 显式属性赋值
COMMIT_MS          = 3.0        # 单次 commit log flush


def model_baseline(n_vouchers=100_000, items_per_voucher=5, fields_per_item=37):
    """原版 baseline：每张凭证 ~10 次 DB 访问 + 反射赋值 + per-bill commit"""
    breakdown = {}

    # 1) getInfoByID("") 3 次 SELECT TOP 0 / 凭证（PR-2 才优化，先列出当前耗时）
    breakdown["getInfoByID 3 SELECT TOP 0"] = n_vouchers * 3 * SQL_RPC_MS

    # 2) BeforeAction: 4 张元数据表 IN list 查询 / 凭证
    breakdown["BeforeAction FI/Corp/Emp/Acc x4"] = n_vouchers * 4 * SQL_RPC_MS

    # 3) CheckDateValidate: 2 次 AccPeriod / 凭证
    breakdown["CheckDateValidate AccPeriod x2"] = n_vouchers * 2 * SQL_RPC_MS

    # 4) t_FVou_M.Init: GL2_AchFIID 全表 SELECT / 凭证（每个 New 实例）
    breakdown["t_FVou_M.Init GL2_AchFIID"] = n_vouchers * 1 * SQL_RPC_MS

    # 5) t_FVou_M.GetFIIDByPrdt: 平均每凭证 2 次（成本 + 收入）每次 1-3 步 Filter（不算 SQL，只是 ADO Filter）
    #    这里不计入 baseline 因为已经在 mrsDocFIRel 中（除了 getPrdtClsTree 上溯，跳过）

    # 6) CreateVouNo stored proc / 凭证
    breakdown["CreateVouNo stored proc"] = n_vouchers * 1 * SQL_RPC_MS

    # 7) ss_YWVouUpdateAchTimes / 凭证（PR-3 优化，先列出）
    breakdown["ss_YWVouUpdateAchTimes"] = n_vouchers * 1 * SQL_RPC_MS

    # 8) ss_AfterSaveUpdateGL2VouBill / 凭证（PR-3 优化，先列出）
    breakdown["ss_AfterSaveUpdateGL2VouBill"] = n_vouchers * 1 * SQL_RPC_MS

    # 9) ADO Update 主表 + 每行明细
    rows = n_vouchers * (1 + items_per_voucher)
    breakdown["ADO Update FVou_M + FVou_I rows"] = rows * ADO_UPDATE_MS

    # 10) meAddRowI CallByName 反射
    cb_calls = n_vouchers * items_per_voucher * fields_per_item
    breakdown["meAddRowI CallByName 反射"] = cb_calls * CALLBYNAME_US / 1000  # us -> ms

    # 11) per-bill commit
    breakdown["per-bill COMMIT"] = n_vouchers * COMMIT_MS

    return breakdown


def model_after_pr1(n_vouchers=100_000, items_per_voucher=5, fields_per_item=37):
    """PR-1 之后：A1+A3+A4+A5+A6 生效。A2/B/C/D 未变。"""
    breakdown = {}

    # PR-1 一次性的元数据加载（整 batch 摊销）
    META_LOAD_MS = (
        SQL_RPC_MS              # FinanceItems
        + SQL_RPC_MS            # Corp
        + SQL_RPC_MS            # Emp
        + SQL_RPC_MS            # Account
        + SQL_RPC_MS            # AccPeriod
        + SQL_RPC_MS            # GL2_AchFIID + FinanceItems join
        + SQL_RPC_MS            # F_VouCls_Bills
    )
    breakdown["[新增] VouMetaCache.LoadAll (1 次)"] = META_LOAD_MS

    # 1) getInfoByID("") 不变（A2 在 PR-2）
    breakdown["getInfoByID 3 SELECT TOP 0"] = n_vouchers * 3 * SQL_RPC_MS

    # 2) BeforeAction：消除 → 只剩 Dictionary 查找
    bf_lookups = n_vouchers * items_per_voucher * 4   # 4 张表
    breakdown["[消除] BeforeAction FI/Corp/Emp/Acc"] = bf_lookups * DICT_LOOKUP_US / 1000

    # 3) CheckDateValidate：消除 → Dictionary 查找
    breakdown["[消除] CheckDateValidate AccPeriod"] = n_vouchers * 2 * DICT_LOOKUP_US / 1000

    # 4) t_FVou_M.Init：消除（VouMetaCache 已加载）
    breakdown["[消除] t_FVou_M.Init"] = 0

    # 5) GetFIIDByPrdt：5 步 Dictionary 查找 / 凭证（平均 2 次调用）
    breakdown["GetFIIDByPrdt 5-step lookup"] = n_vouchers * 2 * 5 * DICT_LOOKUP_US / 1000

    # 6) CreateVouNo 不变（PR-3 优化）
    breakdown["CreateVouNo stored proc"] = n_vouchers * 1 * SQL_RPC_MS

    # 7) ss_YWVouUpdateAchTimes 不变（PR-3）
    breakdown["ss_YWVouUpdateAchTimes"] = n_vouchers * 1 * SQL_RPC_MS

    # 8) ss_AfterSaveUpdateGL2VouBill 不变（PR-3）
    breakdown["ss_AfterSaveUpdateGL2VouBill"] = n_vouchers * 1 * SQL_RPC_MS

    # 9) ADO Update 不变（PR-2）
    rows = n_vouchers * (1 + items_per_voucher)
    breakdown["ADO Update FVou_M + FVou_I rows"] = rows * ADO_UPDATE_MS

    # 10) meAddRowI 显式赋值（消除反射）
    explicit = n_vouchers * items_per_voucher * fields_per_item
    breakdown["[优化] meAddRowI 显式赋值"] = explicit * EXPLICIT_ASSIGN_US / 1000

    # 11) per-bill commit 不变（PR-3）
    breakdown["per-bill COMMIT"] = n_vouchers * COMMIT_MS

    return breakdown


def fmt_ms(ms):
    if ms < 1000:
        return f"{ms:.0f} ms"
    s = ms / 1000
    if s < 60:
        return f"{s:.1f} s"
    m = s / 60
    return f"{m:.1f} min"


def print_breakdown(name, b):
    total = sum(b.values())
    print(f"\n{'='*70}\n{name}\n{'='*70}")
    for k, v in b.items():
        pct = v / total * 100 if total > 0 else 0
        print(f"  {fmt_ms(v):>10s}  ({pct:5.1f}%)  {k}")
    print(f"  {'-'*60}")
    print(f"  {fmt_ms(total):>10s}            合计")
    return total


def main():
    print("=" * 70)
    print("PR-1 性能收益模型")
    print(f"  规模：10 万张凭证，每张 5 行明细，每行 37 字段")
    print(f"  参数：SQL RPC=1ms, ADO Update=1ms, CallByName=8us, DictLookup=0.5us, COMMIT=3ms")

    base = model_baseline()
    after = model_after_pr1()

    t_base = print_breakdown("Baseline（PR-1 之前）", base)
    t_after = print_breakdown("After PR-1（A1+A3+A4+A5+A6）", after)

    saved = t_base - t_after
    pct = saved / t_base * 100
    print(f"\n{'='*70}")
    print(f"PR-1 节省：{fmt_ms(saved)}（{pct:.0f}%）")
    print(f"  {fmt_ms(t_base)} → {fmt_ms(t_after)}")
    print(f"{'='*70}")

    print("\n剩余瓶颈（待 PR-2 / PR-3 / PR-4 处理）：")
    remaining = sorted(after.items(), key=lambda x: -x[1])[:5]
    for k, v in remaining:
        if v > 100:  # 只显示 > 0.1s 的
            print(f"  - {k}: {fmt_ms(v)}")


if __name__ == "__main__":
    main()
