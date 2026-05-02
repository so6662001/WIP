"""
PR-2 性能模型：A2 schema 缓存 + D1 主查询投影列。
"""
SQL_RPC_MS = 1.0
ADO_UPDATE_MS = 1.0
COMMIT_MS = 3.0
DICT_LOOKUP_US = 0.5
EXPLICIT_ASSIGN_US = 0.1


def fmt_ms(ms):
    if ms < 1000: return f"{ms:.0f} ms"
    s = ms / 1000
    if s < 60: return f"{s:.1f} s"
    return f"{s/60:.1f} min"


def model(after_pr2: bool, n=100_000, items=5, fields=37, batch_unique_keys=10_000):
    b = {}
    META = 7 * SQL_RPC_MS
    SCHEMA = 3 * SQL_RPC_MS
    if after_pr2:
        b["[新增] VouSchemaCache.LoadAll"] = SCHEMA
        b["[消除] getInfoByID 3 SELECT TOP 0"] = 0
    else:
        b["getInfoByID 3 SELECT TOP 0"] = n * 3 * SQL_RPC_MS

    # PR-1 已经做的（保留以便看到全图）
    b["VouMetaCache.LoadAll"] = META
    b["BeforeAction (Dict)"] = n * items * 4 * DICT_LOOKUP_US / 1000
    b["CheckDateValidate (Dict)"] = n * 2 * DICT_LOOKUP_US / 1000
    b["t_FVou_M.Init (cache)"] = 0
    b["GetFIIDByPrdt (cache)"] = n * 2 * 5 * DICT_LOOKUP_US / 1000

    # PR-3 + PR-4 还没做的
    b["CreateVouNo stored proc"] = n * SQL_RPC_MS
    b["ss_YWVouUpdateAchTimes"] = n * SQL_RPC_MS
    b["ss_AfterSaveUpdateGL2VouBill"] = n * SQL_RPC_MS
    b["ADO Update FVou_M + FVou_I"] = n * (1 + items) * ADO_UPDATE_MS
    b["meAddRowI 显式赋值"] = n * items * fields * EXPLICIT_ASSIGN_US / 1000
    b["per-bill COMMIT"] = n * COMMIT_MS

    # PR-2 D1 收益：减少 SELECT 拉取的网络流量
    # 假设原 SELECT m.* 比投影列多带回 30 列 × 4 字节 / 行 = 120 字节/行
    # 网络 100Mbps 实际有效 ~10MB/s
    # 10万行 × 120字节 = 12MB → 1.2s
    # 投影列后只多 4 字节/行 → 0.04s
    if not after_pr2:
        b["[D1] meCreVouForXX SELECT m.* 网络传输"] = 1200
    else:
        b["[D1] meCreVouForXX 投影列"] = 100

    return b


def diff(a, b):
    print(f"\n{'='*70}")
    title = "Baseline (after PR-1)" if not (b is a) else "After PR-2"
    rows = sorted(a.items(), key=lambda x: -x[1])
    total = sum(a.values())
    for k, v in rows:
        pct = v / total * 100 if total > 0 else 0
        print(f"  {fmt_ms(v):>10s}  ({pct:5.1f}%)  {k}")
    print(f"  {'-'*60}\n  {fmt_ms(total):>10s}            合计")
    return total


def main():
    base = model(after_pr2=False)
    after = model(after_pr2=True)

    print("=" * 70)
    print("PR-2 性能模型：A2 schema 缓存 + D1 主查询投影列")
    print(f"  规模：10 万张凭证、5 行/张、37 字段")

    print(f"\n{'='*70}\nBaseline (after PR-1)\n{'='*70}")
    t_b = sum(base.values())
    for k, v in sorted(base.items(), key=lambda x: -x[1]):
        if v >= 1: print(f"  {fmt_ms(v):>10s}  {k}")
    print(f"  合计: {fmt_ms(t_b)}")

    print(f"\n{'='*70}\nAfter PR-2\n{'='*70}")
    t_a = sum(after.values())
    for k, v in sorted(after.items(), key=lambda x: -x[1]):
        if v >= 1: print(f"  {fmt_ms(v):>10s}  {k}")
    print(f"  合计: {fmt_ms(t_a)}")

    saved = t_b - t_a
    print(f"\n{'='*70}")
    print(f"PR-2 增量节省：{fmt_ms(saved)}（{saved/t_b*100:.0f}%）")
    print(f"  {fmt_ms(t_b)} → {fmt_ms(t_a)}")

    print(f"\n{'='*70}")
    print("剩余瓶颈（PR-3 + PR-4 处理）：")
    for k, v in sorted(after.items(), key=lambda x: -x[1])[:5]:
        if v >= 1: print(f"  - {k}: {fmt_ms(v)}")


if __name__ == "__main__":
    main()
