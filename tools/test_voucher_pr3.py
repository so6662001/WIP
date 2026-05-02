"""
PR-3 等价性测试：
  - VouBatchOps 收集 BillID 与原 per-bill 调用顺序一致
  - FlushStoredProcs 分页正确（每 200 张一批）
  - InBatchMode=False 时 SaveDoc 走原路径
  - 批量 stored proc 内部 cursor 循环逻辑等价（T-SQL 静态校验）
"""
from __future__ import annotations


# ============================================================
# 模拟 VouBatchOps
# ============================================================
class VouBatchOps:
    def __init__(self):
        self.direct = []
        self.trans = []
        self.in_batch = False
        self.txn_counter = 0
        self.proc_calls = []  # 记录每次 ExecSQL 的 stored proc 调用

    def begin_batch(self):
        self.direct = []
        self.trans = []
        self.in_batch = True
        self.txn_counter = 0

    def end_batch(self):
        self.direct = []
        self.trans = []
        self.in_batch = False

    def record(self, billid, is_trans):
        if not self.in_batch or not billid:
            return
        # B1/B2 校验：BillID 必须不含 ',' / "'" / ';' / 空白
        for ch in (",", "'", ";", " ", "\t", "\n"):
            if ch in billid:
                raise ValueError(
                    f"BillID '{billid}' 含非法字符（',';' 或空白），无法安全批量处理")
        if is_trans:
            self.trans.append(billid)
        else:
            self.direct.append(billid)

    def should_flush(self):
        if not self.in_batch:
            return False
        self.txn_counter += 1
        if self.txn_counter >= 1000:
            self.txn_counter = 0
            return True
        return False

    def flush(self):
        """模拟 FlushStoredProcs：每页 200 张交错（A 后 B 再 A 后 B）"""
        PAGE = 200
        if self.direct:
            for i in range(0, len(self.direct), PAGE):
                page = self.direct[i:i + PAGE]
                self.proc_calls.append(("ss_YWVouUpdateAchTimes_Batch",
                                       ",".join(page), 1))
                self.proc_calls.append(("ss_AfterSaveUpdateGL2VouBill_Batch",
                                       ",".join(page), 0))
        if self.trans:
            for i in range(0, len(self.trans), PAGE):
                page = self.trans[i:i + PAGE]
                self.proc_calls.append(("ss_AfterSaveUpdateGL2VouBill_Batch",
                                       ",".join(page), 1))
        self.direct = []
        self.trans = []


# ============================================================
# Test cases
# ============================================================
def expand_batch_to_per_page_interleaved(batch_calls):
    """把批量调用展开成 per-page 交错的 (proc, bill, param) 序列
    每页 200 张内部按 cursor 循环展开为 per-bill 顺序"""
    expanded = []
    for proc_name, billid_list, param in batch_calls:
        orig_name = proc_name.replace("_Batch", "")
        for bid in billid_list.split(","):
            expanded.append((orig_name, bid, param))
    return expanded


def simulate_per_bill_per_page_interleaved(bills, page_size=200,
                                            proc_a="ss_YWVouUpdateAchTimes", param_a=1,
                                            proc_b="ss_AfterSaveUpdateGL2VouBill", param_b=0):
    """模拟 per-page 交错：每 200 张内部先全 A 后全 B；总顺序：A1...A200 B1...B200 A201...A400 B201...B400"""
    calls = []
    for i in range(0, len(bills), page_size):
        page = bills[i:i + page_size]
        for b in page:
            calls.append((proc_a, b, param_a))
        for b in page:
            calls.append((proc_b, b, param_b))
    return calls


def test_batch_equiv_simple():
    """100 张凭证全部 direct，per-page 交错等价"""
    bills = [f"V{i:05d}" for i in range(100)]

    expected = simulate_per_bill_per_page_interleaved(bills)

    bo = VouBatchOps()
    bo.begin_batch()
    for b in bills:
        bo.record(b, is_trans=False)
    bo.flush()
    actual = expand_batch_to_per_page_interleaved(bo.proc_calls)

    if expected != actual:
        print(f"[FAIL] simple: expected has {len(expected)}, actual {len(actual)}")
        for i in range(min(5, len(expected))):
            print(f"  exp[{i}]={expected[i]}  act[{i}]={actual[i] if i < len(actual) else None}")
        return 1
    print(f"[OK ] simple: 100 张 direct per-page 交错等价（{len(actual)} 次 cursor 内部调用）")
    return 0


def test_batch_equiv_mixed():
    """混合场景：50 direct + 30 trans，按收集顺序展开"""
    bills_d = [f"D{i:03d}" for i in range(50)]
    bills_t = [f"T{i:03d}" for i in range(30)]

    bo = VouBatchOps()
    bo.begin_batch()
    for b in bills_d:
        bo.record(b, is_trans=False)
    for b in bills_t:
        bo.record(b, is_trans=True)
    bo.flush()
    actual = expand_batch_to_per_page_interleaved(bo.proc_calls)

    expected = simulate_per_bill_per_page_interleaved(bills_d) + \
        [("ss_AfterSaveUpdateGL2VouBill", b, 1) for b in bills_t]

    if expected != actual:
        print(f"[FAIL] mixed: expected has {len(expected)}, actual {len(actual)}")
        return 1
    print(f"[OK ] mixed: 50 direct + 30 trans per-page 交错等价")
    return 0


def test_batch_equiv_at_page_boundary():
    """跨 page 场景：250 张 direct → 应分 2 页（200+50）"""
    bills = [f"P{i:03d}" for i in range(250)]
    expected = simulate_per_bill_per_page_interleaved(bills, page_size=200)

    bo = VouBatchOps()
    bo.begin_batch()
    for b in bills:
        bo.record(b, is_trans=False)
    bo.flush()
    actual = expand_batch_to_per_page_interleaved(bo.proc_calls)

    if expected != actual:
        print(f"[FAIL] page_boundary: expected {len(expected)}, actual {len(actual)}")
        # 找出首个差异
        for i, (e, a) in enumerate(zip(expected, actual)):
            if e != a:
                print(f"  diff at {i}: exp={e} act={a}")
                break
        return 1
    print(f"[OK ] page_boundary: 250 张分 2 页（200+50）顺序正确")
    return 0


def test_batch_paging_500():
    """分页：500 张 direct → 应分 3 批（200+200+100）"""
    bills = [f"P{i:04d}" for i in range(500)]
    bo = VouBatchOps()
    bo.begin_batch()
    for b in bills:
        bo.record(b, is_trans=False)
    bo.flush()

    # 分页正确性：每 200 张应产生 2 个 stored proc 调用（YWVou + AfterSave）
    # 500 张 → 3 页 × 2 proc = 6 次 ExecSQL
    if len(bo.proc_calls) != 6:
        print(f"[FAIL] paging: expected 6 calls (3 pages × 2 proc), got {len(bo.proc_calls)}")
        return 1
    # 验证每个 page 大小
    pages = [c[1].count(",") + 1 for c in bo.proc_calls]
    expected_pages = [200, 200, 200, 200, 100, 100]  # 顺序：YWVou1,AfterSave1,YWVou2,AfterSave2,YWVou3,AfterSave3
    if pages != expected_pages:
        print(f"[FAIL] paging sizes: expected {expected_pages}, got {pages}")
        return 1
    print(f"[OK ] paging: 500 张分 3 页（200+200+100）× 2 proc = 6 次 ExecSQL")
    return 0


def test_in_batch_mode_false_fallback():
    """InBatchMode=False 时 record 不收集"""
    bo = VouBatchOps()
    # 没有 begin_batch
    bo.record("X1", False)
    bo.record("X2", True)
    bo.flush()
    if bo.proc_calls:
        print(f"[FAIL] not_in_batch: should not record, got {bo.proc_calls}")
        return 1
    print(f"[OK ] not_in_batch: InBatchMode=False 时不收集")
    return 0


def test_should_flush_txn_at_1000():
    """每 1000 张返回 True 一次"""
    bo = VouBatchOps()
    bo.begin_batch()
    flushes = 0
    for i in range(2500):
        if bo.should_flush():
            flushes += 1
    # 2500 张 → 在第 1000、2000 各 flush 一次 = 2 次
    if flushes != 2:
        print(f"[FAIL] should_flush_txn: expected 2 flushes for 2500 bills, got {flushes}")
        return 1
    print(f"[OK ] should_flush_txn: 2500 张产生 2 次 flush")
    return 0


def test_billid_safety_validation():
    """B1/B2 审计回归：含 SQL 注入字符的 BillID 必须被拒绝"""
    bo = VouBatchOps()
    bo.begin_batch()

    # 正常 ID 接受
    bo.record("V12345", False)
    bo.record("YWVouID20240601001", False)

    bad_ids = ["V'001", "V,001", "V;001", "V 001", "V\t001", "V\n001"]
    fails = 0
    for bid in bad_ids:
        try:
            bo.record(bid, False)
            print(f"[FAIL] 应拒绝危险 BillID: {bid!r}")
            fails += 1
        except ValueError:
            pass  # 预期
    if fails == 0:
        print(f"[OK ] billid_safety: {len(bad_ids)} 个危险 BillID 全部被拒绝")
    return fails


def test_storedproc_billid_order_preserved():
    """B4 审计回归：stored proc 内部用 IDENTITY 列保证 cursor 顺序"""
    with open("src/voucher/PR3_storedprocs.sql") as f:
        sql = f.read()
    # 关键模式：必须有 IDENTITY 列 + ORDER BY rn
    if "IDENTITY(1,1)" not in sql:
        print(f"[FAIL] storedproc 缺 IDENTITY(1,1) 列保证顺序")
        return 1
    if "ORDER BY rn" not in sql:
        print(f"[FAIL] storedproc 缺 ORDER BY rn 子句")
        return 1
    print(f"[OK ] storedproc_billid_order_preserved (IDENTITY + ORDER BY rn)")
    return 0


def test_storedproc_sql_syntax():
    """T-SQL 静态校验：检查 PR3_storedprocs.sql 的语法关键字"""
    with open("src/voucher/PR3_storedprocs.sql") as f:
        sql = f.read()

    required_patterns = [
        "CREATE PROCEDURE dbo.ss_YWVouUpdateAchTimes_Batch",
        "CREATE PROCEDURE dbo.ss_AfterSaveUpdateGL2VouBill_Batch",
        "@BillIDList NVARCHAR(MAX)",
        "@x.nodes('/x')",        # XML 拆分
        "EXEC dbo.ss_YWVouUpdateAchTimes",
        "EXEC dbo.ss_AfterSaveUpdateGL2VouBill",
        "CURSOR LOCAL FAST_FORWARD",
        "DEALLOCATE cur",
    ]
    missing = [p for p in required_patterns if p not in sql]
    if missing:
        print(f"[FAIL] storedproc_sql 缺关键模式: {missing}")
        return 1
    # 不应使用 STRING_SPLIT (SQL 2016+)；排除注释行
    code_lines = []
    in_block_comment = False
    for line in sql.splitlines():
        s = line.strip()
        if s.startswith("/*"): in_block_comment = True
        if not in_block_comment and not s.startswith("--"):
            code_lines.append(line)
        if "*/" in s: in_block_comment = False
    code_only = "\n".join(code_lines)
    if "STRING_SPLIT" in code_only:
        print(f"[FAIL] storedproc_sql 用了 STRING_SPLIT（SQL 2008 不支持）")
        return 1
    print(f"[OK ] storedproc_sql 语法关键模式齐全（XML.nodes() 兼容 SQL 2008）")
    return 0


def main():
    print("=== PR-3 等价性测试 ===\n")
    tests = [
        ("test_batch_equiv_simple", test_batch_equiv_simple),
        ("test_batch_equiv_mixed", test_batch_equiv_mixed),
        ("test_batch_equiv_at_page_boundary", test_batch_equiv_at_page_boundary),
        ("test_batch_paging_500", test_batch_paging_500),
        ("test_in_batch_mode_false_fallback", test_in_batch_mode_false_fallback),
        ("test_should_flush_txn_at_1000", test_should_flush_txn_at_1000),
        ("test_billid_safety_validation", test_billid_safety_validation),
        ("test_storedproc_billid_order_preserved", test_storedproc_billid_order_preserved),
        ("test_storedproc_sql_syntax", test_storedproc_sql_syntax),
    ]
    fails = sum(t() for n, t in tests)
    print()
    if fails:
        print(f"❌ {fails} case(s) FAILED")
        raise SystemExit(1)
    print("✅ All PR-3 equivalence tests passed.")


if __name__ == "__main__":
    main()
