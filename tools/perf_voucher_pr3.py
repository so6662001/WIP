"""PR-3 性能模型：大事务 + stored proc 批量化"""
SQL_RPC_MS = 1.0
ADO_UPDATE_MS = 1.0
COMMIT_MS = 3.0
DICT_LOOKUP_US = 0.5
EXPLICIT_ASSIGN_US = 0.1


def fmt(ms):
    if ms < 1000: return f"{ms:.0f} ms"
    if ms < 60000: return f"{ms/1000:.1f} s"
    return f"{ms/60000:.1f} min"


def model(stage, n=100_000, items=5):
    """stage='after_pr2' or 'after_pr3'"""
    b = {}
    META, SCHEMA = 7 * SQL_RPC_MS, 3 * SQL_RPC_MS
    PAGE = 200

    if stage == "after_pr2":
        b["CreateVouNo stored proc"] = n * SQL_RPC_MS
        b["ss_YWVouUpdateAchTimes (per-bill)"] = n * SQL_RPC_MS
        b["ss_AfterSaveUpdateGL2VouBill (per-bill)"] = n * SQL_RPC_MS
        b["per-bill COMMIT"] = n * COMMIT_MS
    else:
        # PR-3 改造：批量调用 + 大事务
        b["CreateVouNo stored proc"] = n * SQL_RPC_MS  # PR-3 不动这个
        # 批量后：n / 200 次 RPC × 2 procs（直写路径）
        # cursor 循环时间不变（服务端逻辑等价），但服务端 cursor 不算客户端 RTT
        # 假设 stored proc 内部本身的执行时间没变（因为 cursor 调原 proc）
        # 收益 = N × RPC - (N/200) × 2 × RPC ≈ N × RPC × 99%
        n_pages = (n + PAGE - 1) // PAGE
        b["[消除] ss_YWVouUpdateAchTimes per-bill"] = 0
        b["[消除] ss_AfterSaveUpdateGL2VouBill per-bill"] = 0
        b["[新增] FlushStoredProcs 批量"] = n_pages * 2 * SQL_RPC_MS
        # 大事务：从 N × COMMIT 降到 ~1 次大 commit
        # 但事务越大日志越多，commit 时间也大（约 5x）
        b["[消除] per-bill COMMIT"] = 0
        b["[新增] 大事务 COMMIT"] = COMMIT_MS * 5  # 单次大 commit ≈ 15ms

    b["ADO Update FVou_M + FVou_I"] = n * (1 + items) * ADO_UPDATE_MS  # PR-4 处理
    b["meAddRowI 显式赋值"] = n * items * 37 * EXPLICIT_ASSIGN_US / 1000
    b["VouMetaCache.LoadAll"] = META
    b["VouSchemaCache.LoadAll"] = SCHEMA
    b["BeforeAction (Dict)"] = n * items * 4 * DICT_LOOKUP_US / 1000
    b["GetFIIDByPrdt (cache)"] = n * 2 * 5 * DICT_LOOKUP_US / 1000
    b["meCreVouForXX 投影列"] = 100
    return b


def main():
    base = model("after_pr2")
    after = model("after_pr3")

    print("=" * 70)
    print("PR-3 性能模型：大事务 + stored proc 批量化")

    print(f"\n--- Baseline (after PR-2) ---")
    t_b = sum(base.values())
    for k, v in sorted(base.items(), key=lambda x: -x[1])[:10]:
        if v >= 1: print(f"  {fmt(v):>10s}  {k}")
    print(f"  合计: {fmt(t_b)}")

    print(f"\n--- After PR-3 ---")
    t_a = sum(after.values())
    for k, v in sorted(after.items(), key=lambda x: -x[1])[:10]:
        if v >= 1: print(f"  {fmt(v):>10s}  {k}")
    print(f"  合计: {fmt(t_a)}")

    saved = t_b - t_a
    print(f"\n增量节省：{fmt(saved)}（{saved/t_b*100:.0f}%）")
    print(f"  {fmt(t_b)} → {fmt(t_a)}")

    print(f"\n剩余瓶颈（PR-4 处理）：")
    for k, v in sorted(after.items(), key=lambda x: -x[1])[:3]:
        if v >= 1: print(f"  - {k}: {fmt(v)}")


if __name__ == "__main__":
    main()
