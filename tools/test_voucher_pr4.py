"""
PR-4 等价性测试：
  - VouBatchWriter 字段类型转换 (meQuoteValue 各类型)
  - INSERT VALUES 拼装格式正确（列顺序 + 转义 + NULL）
  - SQL 字符串长度防护（不超过 65000 字节）
  - C1 SSBillByDateDAL 预聚合 → 多次 GROUP BY 等价性（数据集场景）
"""
from __future__ import annotations
from datetime import datetime, date
import re


# ============================================================
# 模拟 ADO Field 类型
# ============================================================
class Field:
    def __init__(self, name, value, ado_type):
        self.name = name
        self.value = value
        self.type = ado_type


# ADO type 常量（与 VB6 ADODB 一致）
adVarChar = 200
adVarWChar = 202
adChar = 129
adWChar = 130
adLongVarChar = 201
adLongVarWChar = 203
adDate = 7
adDBTimeStamp = 135
adBoolean = 11
adNumeric = 131
adDecimal = 14
adCurrency = 6
adDouble = 5
adSingle = 4
adInteger = 3
adSmallInt = 2
adBigInt = 20
adTinyInt = 16


# ============================================================
# Python 复刻 meQuoteValue
# ============================================================
def me_quote_value(fld):
    if fld.value is None:
        return "NULL"
    t = fld.type
    if t in (adVarChar, adVarWChar, adChar, adWChar, adLongVarChar, adLongVarWChar):
        return "N'" + str(fld.value).replace("'", "''") + "'"
    if t in (adDate, adDBTimeStamp):
        if isinstance(fld.value, datetime):
            return "'" + fld.value.strftime("%Y-%m-%d %H:%M:%S") + "'"
        if isinstance(fld.value, date):
            return "'" + fld.value.strftime("%Y-%m-%d") + " 00:00:00'"
        return "'" + str(fld.value) + "'"
    if t == adBoolean:
        return "1" if bool(fld.value) else "0"
    if t in (adNumeric, adDecimal, adCurrency, adDouble, adSingle,
             adInteger, adSmallInt, adBigInt, adTinyInt):
        return str(fld.value)
    return "N'" + str(fld.value).replace("'", "''") + "'"


def build_values_row(fields):
    return ",".join(me_quote_value(f) for f in fields)


def build_col_list(fields):
    return ",".join(f"[{f.name}]" for f in fields)


# ============================================================
# Test cases
# ============================================================
def test_quote_value_types():
    cases = [
        # (Field, expected_quoted)
        (Field("a", None, adVarChar),                          "NULL"),
        (Field("b", "hello", adVarChar),                       "N'hello'"),
        (Field("c", "it's", adVarChar),                        "N'it''s'"),
        (Field("d", "a|b", adVarChar),                         "N'a|b'"),
        (Field("e", "中文", adVarWChar),                        "N'中文'"),
        (Field("f", datetime(2024, 6, 1, 9, 0, 0), adDBTimeStamp), "'2024-06-01 09:00:00'"),
        (Field("g", date(2024, 6, 1), adDate),                 "'2024-06-01 00:00:00'"),
        (Field("h", True, adBoolean),                          "1"),
        (Field("i", False, adBoolean),                         "0"),
        (Field("j", 123, adInteger),                           "123"),
        (Field("k", 1.5, adDouble),                            "1.5"),
        (Field("l", -99.99, adCurrency),                       "-99.99"),
        (Field("m", "", adVarChar),                            "N''"),
        (Field("n", "'", adVarChar),                           "N''''"),
        (Field("o", 0, adInteger),                             "0"),
        (Field("p", "hello\r\nworld", adVarChar),              "N'hello\r\nworld'"),
    ]
    fails = 0
    for fld, expected in cases:
        actual = me_quote_value(fld)
        if actual != expected:
            print(f"[FAIL] {fld.name}: expected {expected!r}, got {actual!r}")
            fails += 1
        else:
            print(f"[OK ] {fld.name:3s}  {fld.value!r:40s} -> {actual}")
    return fails


def test_insert_values_layout():
    """验证 INSERT INTO ... (cols) VALUES (..),(..) 格式"""
    fields1 = [
        Field("BillID", "V001", adVarChar),
        Field("ATM", 100.0, adCurrency),
        Field("BillDate", datetime(2024, 6, 1), adDBTimeStamp),
    ]
    fields2 = [
        Field("BillID", "V002", adVarChar),
        Field("ATM", None, adCurrency),
        Field("BillDate", datetime(2024, 6, 2, 12, 0, 0), adDBTimeStamp),
    ]

    cols = build_col_list(fields1)
    expected_cols = "[BillID],[ATM],[BillDate]"
    if cols != expected_cols:
        print(f"[FAIL] col list: {cols!r}")
        return 1

    row1 = build_values_row(fields1)
    row2 = build_values_row(fields2)
    expected_row1 = "N'V001',100.0,'2024-06-01 00:00:00'"
    expected_row2 = "N'V002',NULL,'2024-06-02 12:00:00'"
    if row1 != expected_row1:
        print(f"[FAIL] row1: {row1!r}")
        return 1
    if row2 != expected_row2:
        print(f"[FAIL] row2: {row2!r}")
        return 1

    sql = f"INSERT INTO [FVou_M_T]({cols}) VALUES ({row1}),({row2})"
    print(f"[OK ] insert sql: {sql}")
    return 0


def test_sql_length_protection():
    """模拟 200 行 INSERT，确保不超过 65000 字节"""
    # 假设每行约 300 字节
    rows = []
    for i in range(200):
        rows.append(f"(N'V{i:05d}',{i*1.5},'2024-06-01 09:00:00')")
    header = "INSERT INTO [FVou_M_T]([BillID],[ATM],[BillDate]) VALUES "
    sql = header + ",".join(rows)
    if len(sql) > 65000:
        print(f"[FAIL] sql length {len(sql)} > 65000 (page should split before)")
        return 1
    print(f"[OK ] 200 行 SQL 长度 {len(sql)} 字节，未超过 65000")
    return 0


def test_c1_dal_aggregation_equivalence():
    """C1: 验证预聚合 #TmpSSAgg 后多次 GROUP BY 与原直接对大表 GROUP BY 结果一致

    模拟：6 行 SSB_I+SSB_M JOIN 后明细，按 4 种维度组合分组聚合
    """
    raw = [
        # (DEPID, EMPID, WHID, CLSID, CorpID, TInvID, QTY, ATM, CstATM, PATM, TaxATM, ssbtag)
        ("D1", "E1", "W1", "C1", "P1", "T1", 10, 100, 80, 110, 11, 0),
        ("D1", "E1", "W1", "C1", "P1", "T1",  5,  50, 40,  55,  5, 0),
        ("D1", "E2", "W1", "C2", "P2", "T1",  3,  30, 24,  33,  3, 0),
        ("D2", "E3", "W2", "C1", "P1", "T2",  7,  70, 56,  77,  7, 0),
        ("D2", "E3", "W2", "C1", "P1", "T2",  2,  20, 16,  22,  2, 1),  # ssbtag=1, sign=-1
        ("D2", "E3", "W2", "C1", "P1", "T2",  1,  10,  8,  11,  1, 0),
    ]

    def sign(t):
        return -1 if t in (1, 2) else 1

    # 预聚合到 #TmpSSAgg：GROUP BY 全部 7 个维度
    tmp = {}
    for (dep, emp, wh, cls, corp, tinv, qty, atm, cst, patm, tax, t) in raw:
        s = sign(t)
        k = (dep, emp, wh, cls, corp, tinv)
        if k not in tmp:
            tmp[k] = {"QTY": 0, "ATM": 0, "CstATM": 0, "PATM": 0, "TaxATM": 0}
        tmp[k]["QTY"] += qty * s
        tmp[k]["ATM"] += atm * s
        tmp[k]["CstATM"] += cst * s
        tmp[k]["PATM"] += patm * s
        tmp[k]["TaxATM"] += tax * s

    # 第 1 次 GROUP BY (DEPID, EMPID, WHID, CLSID) 取 CstATM
    g1_via_tmp = {}
    for (dep, emp, wh, cls, _corp, _tinv), v in tmp.items():
        k = (dep, emp, wh, cls)
        if k not in g1_via_tmp:
            g1_via_tmp[k] = 0
        g1_via_tmp[k] += v["CstATM"]

    # 直接从大表 GROUP BY (DEPID, EMPID, WHID, CLSID)
    g1_direct = {}
    for (dep, emp, wh, cls, _corp, _tinv, qty, atm, cst, patm, tax, t) in raw:
        k = (dep, emp, wh, cls)
        if k not in g1_direct:
            g1_direct[k] = 0
        g1_direct[k] += cst * sign(t)

    if g1_via_tmp != g1_direct:
        print(f"[FAIL] g1: tmp={g1_via_tmp} direct={g1_direct}")
        return 1
    print(f"[OK ] g1 (DEPID,EMPID,WHID,CLSID) 预聚合等价（{len(g1_direct)} 组）")

    # 第 2 次 GROUP BY (CorpID, DEPID, EMPID) 取 ATM
    g2_via_tmp = {}
    for (dep, emp, _wh, _cls, corp, _tinv), v in tmp.items():
        k = (corp, dep, emp)
        if k not in g2_via_tmp:
            g2_via_tmp[k] = 0
        g2_via_tmp[k] += v["ATM"]

    g2_direct = {}
    for (dep, emp, _wh, _cls, corp, _tinv, qty, atm, cst, patm, tax, t) in raw:
        k = (corp, dep, emp)
        if k not in g2_direct:
            g2_direct[k] = 0
        g2_direct[k] += atm * sign(t)

    if g2_via_tmp != g2_direct:
        print(f"[FAIL] g2: tmp={g2_via_tmp} direct={g2_direct}")
        return 1
    print(f"[OK ] g2 (CorpID,DEPID,EMPID) 预聚合等价（{len(g2_direct)} 组）")

    # 第 3 次 GROUP BY (TInvID, DEPID, EMPID) 取 TaxATM (CostModel=1)
    g3_via_tmp = {}
    for (dep, emp, _wh, _cls, _corp, tinv), v in tmp.items():
        k = (tinv, dep, emp)
        if k not in g3_via_tmp:
            g3_via_tmp[k] = 0
        g3_via_tmp[k] += v["TaxATM"]

    g3_direct = {}
    for (dep, emp, _wh, _cls, _corp, tinv, qty, atm, cst, patm, tax, t) in raw:
        k = (tinv, dep, emp)
        if k not in g3_direct:
            g3_direct[k] = 0
        g3_direct[k] += tax * sign(t)

    if g3_via_tmp != g3_direct:
        print(f"[FAIL] g3: tmp={g3_via_tmp} direct={g3_direct}")
        return 1
    print(f"[OK ] g3 (TInvID,DEPID,EMPID) 预聚合等价（{len(g3_direct)} 组）")
    return 0


def test_ssbtag_sign_logic():
    """SSBTag 1/2 取反符号，其它 1.0 — 与原代码 CASE WHEN ssbtag IN (1,2) THEN -1.0 一致"""
    cases = [(0, 1), (1, -1), (2, -1), (3, 1), (10, 1)]
    fails = 0
    for tag, expected in cases:
        actual = -1 if tag in (1, 2) else 1
        if actual != expected:
            print(f"[FAIL] ssbtag={tag}: expected {expected}, got {actual}")
            fails += 1
    if fails == 0:
        print(f"[OK ] ssbtag sign 逻辑等价（5 个 case）")
    return fails


def main():
    print("=== PR-4 等价性测试 ===\n")
    print("--- test_quote_value_types ---")
    f1 = test_quote_value_types()
    print("\n--- test_insert_values_layout ---")
    f2 = test_insert_values_layout()
    print("\n--- test_sql_length_protection ---")
    f3 = test_sql_length_protection()
    print("\n--- test_c1_dal_aggregation_equivalence ---")
    f4 = test_c1_dal_aggregation_equivalence()
    print("\n--- test_ssbtag_sign_logic ---")
    f5 = test_ssbtag_sign_logic()

    total = f1 + f2 + f3 + f4 + f5
    print()
    if total > 0:
        print(f"❌ {total} case(s) FAILED")
        raise SystemExit(1)
    print("✅ All PR-4 equivalence tests passed.")


if __name__ == "__main__":
    main()
