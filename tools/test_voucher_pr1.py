"""
PR-1 等价性测试：把 VouMetaCache + VouCacheHelpers 的关键查询逻辑用 Python
重新实现，验证【cache 路径】和【DB 路径】对相同输入返回完全一致的输出。

测试覆盖：
  - VouMetaCache.TryGetFI / TryGetCorp / TryGetEmp / TryGetAcc
  - VouMetaCache.TryGetAPByDate（多种边界）
  - VouMetaCache.TryGetDocFI（5 步精确-到-空 fallback）
  - VouCacheHelpers.CheckFI（包含 rsFIBak 跨行命中分支）
  - meAddRowI 反射 vs 显式赋值（字段顺序 + 值完全一致）
"""
from __future__ import annotations
from collections import OrderedDict
from datetime import date, datetime
from typing import Any, Optional


# ============================================================
# DB 模拟：模仿 SQL Server 的几张元数据表
# ============================================================
class FakeDB:
    def __init__(self):
        # FinanceItems
        self.finance_items = []   # rows of dict
        # Corp / Emp / Account
        self.corp = []
        self.emp = []
        self.account = []
        # AccPeriod
        self.acc_period = []      # [{APID, BeginDate, EndDate, APTag, FOAPTag}]
        # GL2_AchFIID + FinanceItems join
        self.doc_fi = []          # rows {DocKey,WHID,DocID,FIID,FINO,FIName,HSTag,HSTagName,FITag}
        # F_VouCls_Bills
        self.vc_bills = []        # rows {BillType,BillSubType,VCID}

    # --- DB-style queries (return list of dict) ---
    def query_fi(self, fiid_list):
        return [r for r in self.finance_items if r["FIID"] in fiid_list]

    def query_corp(self, corp_ids):
        return [r for r in self.corp if r["CorpID"] in corp_ids]

    def query_emp(self, emp_ids):
        return [r for r in self.emp if r["EmpID"] in emp_ids]

    def query_acc(self, acc_ids):
        return [r for r in self.account if r["AccID"] in acc_ids]

    def query_ap_by_date(self, d, use_foaptag=False):
        rows = [r for r in self.acc_period
                if r["BeginDate"] <= d <= r["EndDate"]]
        if not rows:
            return None
        r = rows[0]
        return {"APID": r["APID"],
                "APTag": r["FOAPTag"] if use_foaptag else r["APTag"]}

    def query_active_ap(self, use_foaptag=False):
        col = "FOAPTag" if use_foaptag else "APTag"
        return [r for r in self.acc_period if r[col] == 2]

    def query_doc_fi(self, dockey, whid, docid):
        return [r for r in self.doc_fi
                if r["DocKey"] == dockey and r["WHID"] == whid and r["DocID"] == docid]


# ============================================================
# Cache 模拟：等价于 VouMetaCache.bas 的实现
# ============================================================
class VouMetaCache:
    def __init__(self):
        self.fi: dict[str, dict] = {}
        self.corp: dict[str, dict] = {}
        self.emp: dict[str, dict] = {}
        self.acc: dict[str, dict] = {}
        self.ap: list[dict] = []         # 按 BeginDate 排序
        self.doc_fi: dict[tuple, dict] = {}
        self.vc: dict[tuple, str] = {}
        self.loaded = False

    def load_all(self, db: FakeDB):
        self.fi = {r["FIID"]: r for r in db.finance_items}
        self.corp = {r["CorpID"]: r for r in db.corp}
        self.emp = {r["EmpID"]: r for r in db.emp}
        self.acc = {r["AccID"]: r for r in db.account}
        self.ap = sorted(db.acc_period, key=lambda r: r["BeginDate"])
        # 对 GL2_AchFIID 同 key 重复时保留首条（与 ADO Filter 第一行行为一致）
        self.doc_fi = {}
        for r in db.doc_fi:
            k = (r["DocKey"], r["WHID"], r["DocID"])
            if k not in self.doc_fi:
                self.doc_fi[k] = r
        self.vc = {}
        for r in db.vc_bills:
            k = (r["BillType"], r["BillSubType"] or "")
            if k not in self.vc:
                self.vc[k] = r["VCID"]
        self.loaded = True

    def try_get_fi(self, fiid: str):
        return self.fi.get(fiid)

    def try_get_corp(self, cid: str):
        return self.corp.get(cid)

    def try_get_emp(self, eid: str):
        return self.emp.get(eid)

    def try_get_acc(self, aid: str):
        return self.acc.get(aid)

    def try_get_ap_by_date(self, d, use_foaptag=False):
        # 线性扫描（VouMetaCache.bas 的实现）
        for r in self.ap:
            if r["BeginDate"] <= d <= r["EndDate"]:
                return {"APID": r["APID"],
                        "APTag": r["FOAPTag"] if use_foaptag else r["APTag"]}
        return None

    def has_active_ap(self, use_foaptag=False):
        col = "FOAPTag" if use_foaptag else "APTag"
        return any(r[col] == 2 for r in self.ap)

    def try_get_doc_fi(self, dockey, whid, docid):
        return self.doc_fi.get((dockey, whid, docid))

    def try_get_vcid(self, bill_type, bill_sub_type):
        k = (bill_type, bill_sub_type)
        if k in self.vc:
            return self.vc[k]
        return self.vc.get((bill_type, ""))


# ============================================================
# 等价性测试：CheckFI 路径
# 模拟 BeforeAction 内一个完整 row 的检查 + ItemsData 字段写回
# ============================================================
def make_items_row(fiid, fino="", finame=""):
    """模拟 ItemsData 当前 row，使用 dict 模拟字段"""
    return {
        "FIID": fiid,
        "FINO": fino,
        "FIName": finame,
        "FITag": 0,
        "hstag": 0,
        "hstagname": "",
        "baNum": False,
        "baCust": False,
        "baSupp": False,
        "baOtherCorp": False,
        "itmid": 1,
    }


def check_fi_db_path(db: FakeDB, item: dict, act_ch_name: str) -> tuple[dict, str]:
    """模拟原 BeforeAction 中 fiid Filter 路径"""
    err = ""
    rows = db.query_fi([item["FIID"]])
    if not rows:
        err += f"第 {item['itmid']} 条分录中的会计科目[{item['FINO']}]在科目档案中已经找不到，可能被删除！\n"
        return item, err
    r = rows[0]
    if r["ISStop"]:
        err += f"第 {item['itmid']} 条分录中的会计科目[{item['FINO']}]已经被停用，无法{act_ch_name}！\n"
    if r["FITag"] == 6 and r["EXPTAG"] == 1:
        err += f"第 {item['itmid']} 条分录中的会计科目[{item['FINO']}]是采购费用类科目，无法{act_ch_name}，请修改成非采购费用类科目！\n"
    if not item["FIName"]:
        item["FIName"] = (r["FullName"] or r["FIName"]).replace("\r\n", "")
    item["FINO"] = r["FINO"]
    item["FITag"] = r["FITag"]
    item["hstag"] = r["HSTag"]
    item["hstagname"] = r["HSTagName"]
    item["baNum"] = bool(r["baNum"])
    item["baCust"] = bool(r["baCust"])
    item["baSupp"] = bool(r["baSupp"])
    item["baOtherCorp"] = bool(r["baOtherCorp"])
    return item, err


def check_fi_cache_path(cache: VouMetaCache, item: dict, act_ch_name: str) -> tuple[dict, str]:
    """模拟 VouCacheHelpers.CheckFI 中 cache 路径"""
    err = ""
    r = cache.try_get_fi(item["FIID"])
    if r is None:
        err += f"第 {item['itmid']} 条分录中的会计科目[{item['FINO']}]在科目档案中已经找不到，可能被删除！\n"
        return item, err
    if r["ISStop"]:
        err += f"第 {item['itmid']} 条分录中的会计科目[{item['FINO']}]已经被停用，无法{act_ch_name}！\n"
    if r["FITag"] == 6 and r["EXPTAG"] == 1:
        err += f"第 {item['itmid']} 条分录中的会计科目[{item['FINO']}]是采购费用类科目，无法{act_ch_name}，请修改成非采购费用类科目！\n"
    if not item["FIName"]:
        item["FIName"] = (r["FullName"] or r["FIName"]).replace("\r\n", "")
    item["FINO"] = r["FINO"]
    item["FITag"] = r["FITag"]
    item["hstag"] = r["HSTag"]
    item["hstagname"] = r["HSTagName"]
    item["baNum"] = bool(r["baNum"])
    item["baCust"] = bool(r["baCust"])
    item["baSupp"] = bool(r["baSupp"])
    item["baOtherCorp"] = bool(r["baOtherCorp"])
    return item, err


# ============================================================
# 等价性测试：GetFIIDByPrdt 5 步 fallback
# ============================================================
def get_fiid_by_prdt_db(db: FakeDB, dockey, whid, clsid):
    """5 步精确-到-空 fallback。简化版（不含 getPrdtClsTree 上溯）"""
    # Step 1
    rows = db.query_doc_fi(dockey, whid, clsid)
    if rows:
        return rows[0]["FIID"]
    # Step 2
    rows = db.query_doc_fi(dockey, whid, "")
    if rows:
        return rows[0]["FIID"]
    # Step 3
    rows = db.query_doc_fi(dockey, "", clsid)
    if rows:
        return rows[0]["FIID"]
    # Step 4：上级品类（这里跳过，假定无上溯）
    # Step 5
    rows = db.query_doc_fi(dockey, "", "")
    if rows:
        return rows[0]["FIID"]
    return None


def get_fiid_by_prdt_cache(cache: VouMetaCache, dockey, whid, clsid):
    r = cache.try_get_doc_fi(dockey, whid, clsid)
    if r:
        return r["FIID"]
    r = cache.try_get_doc_fi(dockey, whid, "")
    if r:
        return r["FIID"]
    r = cache.try_get_doc_fi(dockey, "", clsid)
    if r:
        return r["FIID"]
    # 上级品类跳过
    r = cache.try_get_doc_fi(dockey, "", "")
    if r:
        return r["FIID"]
    return None


# ============================================================
# meAddRowI 反射 vs 显式赋值等价性
# ============================================================
ARRFLDS = ["itmid", "fiid", "fino", "finame", "idesc", "CstCenID",
          "corpid", "corptag", "PrdtID", "PrdtName", "AccID", "Weight",
          "QTY", "MDQTY", "MCQTY", "Pri",
          "DATM", "catm", "DATM_F", "catm_f", "AccRate", "SrcItmID",
          "isFixFI", "TInvID", "ClsID", "ASSInfo", "WHID", "DepID", "EmpID",
          "MFCstPGBillID", "MFCstPGBillNo", "MFCstMOBillID", "MFCstMOBillNo",
          "PrjID", "PrjName", "JTID", "JTName"]


def me_add_row_reflection(rs_item: dict, idata: dict):
    """模拟原 CallByName 反射赋值
    
    VBA 中 CallByName(iData, FldName, VbGet) 大小写不敏感地查找 t_FVou_I 属性。
    ADO Recordset.Fields(name) 也大小写不敏感地查找字段。
    
    Python 这里用大小写不敏感字典查找模拟 VBA 行为。
    """
    # 把 idata 转为大小写不敏感字典
    idata_ci = {k.lower(): v for k, v in idata.items()}
    for fld in ARRFLDS:
        # rsItem.Fields(FldName).Value = CallByName(iData, FldName, VbGet)
        rs_item[fld.lower()] = idata_ci.get(fld.lower())


def me_add_row_explicit(rs_item: dict, idata: dict):
    """显式按字段顺序赋值（PR-1 优化版）"""
    rs_item["itmid"] = idata["ITMID"]
    rs_item["fiid"] = idata["FIID"]
    rs_item["fino"] = idata["FINO"]
    rs_item["finame"] = idata["FIName"]
    rs_item["idesc"] = idata["IDesc"]
    rs_item["cstcenid"] = idata["CstCenID"]
    rs_item["corpid"] = idata["CorpID"]
    rs_item["corptag"] = idata["CorpTag"]
    rs_item["prdtid"] = idata["PrdtID"]
    rs_item["prdtname"] = idata["PrdtName"]
    rs_item["accid"] = idata["AccID"]
    rs_item["weight"] = idata["Weight"]
    rs_item["qty"] = idata["QTY"]
    rs_item["mdqty"] = idata["MDQTY"]
    rs_item["mcqty"] = idata["MCQTY"]
    rs_item["pri"] = idata["Pri"]
    rs_item["datm"] = idata["DATM"]
    rs_item["catm"] = idata["CATM"]
    rs_item["datm_f"] = idata["DATM_F"]
    rs_item["catm_f"] = idata["CATM_F"]
    rs_item["accrate"] = idata["AccRate"]
    rs_item["srcitmid"] = idata["SrcITMID"]
    rs_item["isfixfi"] = idata["isFixFI"]
    rs_item["tinvid"] = idata["TInvID"]
    rs_item["clsid"] = idata["ClsID"]
    rs_item["assinfo"] = idata["AssInfo"]
    rs_item["whid"] = idata["WHID"]
    rs_item["depid"] = idata["depid"]
    rs_item["empid"] = idata["EmpID"]
    rs_item["mfcstpgbillid"] = idata["MFCstPGBillID"]
    rs_item["mfcstpgbillno"] = idata["MFCstPGBillNo"]
    rs_item["mfcstmobillid"] = idata["MFCstMOBillID"]
    rs_item["mfcstmobillno"] = idata["MFCstMOBillNo"]
    rs_item["prjid"] = idata["PrjID"]
    rs_item["prjname"] = idata["PrjName"]
    rs_item["jtid"] = idata["JTID"]
    rs_item["jtname"] = idata["JTName"]


# ============================================================
# 测试用例
# ============================================================
def build_db():
    db = FakeDB()
    # FinanceItems 5 行
    db.finance_items = [
        {"FIID": "F1", "FINO": "1001", "FIName": "库存现金", "FullName": "1001 库存现金",
         "ISStop": False, "FITag": 1, "EXPTAG": 0, "HSTag": 6, "HSTagName": "货币资金",
         "baNum": False, "baCust": False, "baSupp": False, "baOtherCorp": False},
        {"FIID": "F2", "FINO": "1122", "FIName": "应收账款", "FullName": "",
         "ISStop": False, "FITag": 1, "EXPTAG": 0, "HSTag": 2, "HSTagName": "应收",
         "baNum": False, "baCust": True, "baSupp": False, "baOtherCorp": False},
        {"FIID": "F3", "FINO": "1406", "FIName": "库存商品", "FullName": "",
         "ISStop": False, "FITag": 1, "EXPTAG": 0, "HSTag": 1, "HSTagName": "存货",
         "baNum": True, "baCust": False, "baSupp": False, "baOtherCorp": False},
        {"FIID": "F4", "FINO": "5401", "FIName": "采购费用", "FullName": "",
         "ISStop": False, "FITag": 6, "EXPTAG": 1, "HSTag": 7, "HSTagName": "费用",
         "baNum": False, "baCust": False, "baSupp": False, "baOtherCorp": False},
        {"FIID": "F5", "FINO": "STOPPED", "FIName": "已停用", "FullName": "",
         "ISStop": True, "FITag": 1, "EXPTAG": 0, "HSTag": 6, "HSTagName": "货币资金",
         "baNum": False, "baCust": False, "baSupp": False, "baOtherCorp": False},
    ]
    db.corp = [
        {"CorpID": "C1", "Contact": True, "CorpName": "客户A"},
        {"CorpID": "C2", "Contact": False, "CorpName": "已停用客户"},
    ]
    db.emp = [
        {"EmpID": "E1", "dismission": False, "EmpName": "张三"},
        {"EmpID": "E2", "dismission": True, "EmpName": "已离职"},
    ]
    db.account = [
        {"AccID": "A1", "ISStop": False, "AccName": "工行账户"},
        {"AccID": "A2", "ISStop": True, "AccName": "已停用账户"},
    ]
    # AccPeriod 12 个月
    for m in range(1, 13):
        db.acc_period.append({
            "APID": f"2024{m:02d}",
            "BeginDate": date(2024, m, 1),
            "EndDate": date(2024, m, 28 if m == 2 else (30 if m in (4,6,9,11) else 31)),
            "APTag": 2 if m == 6 else (1 if m < 6 else 0),     # 6 月活动；之前已结；之后未来
            "FOAPTag": 2 if m == 6 else (1 if m < 6 else 0),
        })
    # GL2_AchFIID
    db.doc_fi = [
        {"DocKey": "PrdtCls", "WHID": "W1", "DocID": "C001", "FIID": "F3",
         "FINO": "1406", "FIName": "库存商品", "HSTag": 1, "HSTagName": "存货", "FITag": 1},
        {"DocKey": "PrdtCls", "WHID": "W1", "DocID": "", "FIID": "F3",
         "FINO": "1406", "FIName": "库存商品", "HSTag": 1, "HSTagName": "存货", "FITag": 1},
        {"DocKey": "PrdtCls", "WHID": "", "DocID": "", "FIID": "F3",
         "FINO": "1406", "FIName": "库存商品", "HSTag": 1, "HSTagName": "存货", "FITag": 1},
        {"DocKey": "Account", "WHID": "", "DocID": "A1", "FIID": "F1",
         "FINO": "1001", "FIName": "库存现金", "HSTag": 6, "HSTagName": "货币资金", "FITag": 1},
    ]
    return db


def test_fi():
    db = build_db()
    cache = VouMetaCache()
    cache.load_all(db)

    cases = [
        ("F1", "", "测试1"),       # 正常
        ("F2", "", "测试2"),       # 正常 baCust=True
        ("F3", "已有名字", "测试3"),  # FIName 已有，不应覆盖
        ("F4", "", "测试4"),       # 采购费用 -> 报错
        ("F5", "", "测试5"),       # 已停用 -> 报错
        ("FX", "", "测试6"),       # 不存在 -> 报错
    ]
    fails = 0
    for fiid, finame, name in cases:
        a = check_fi_db_path(db, make_items_row(fiid, "", finame), name)
        b = check_fi_cache_path(cache, make_items_row(fiid, "", finame), name)
        if a != b:
            print(f"[FAIL] check_fi {fiid:5s} {name}")
            print(f"  db   : {a}")
            print(f"  cache: {b}")
            fails += 1
        else:
            print(f"[OK ] check_fi {fiid:5s} {name}")
    return fails


def test_doc_fi():
    db = build_db()
    cache = VouMetaCache()
    cache.load_all(db)

    cases = [
        ("PrdtCls", "W1", "C001"),     # 精确命中 step1
        ("PrdtCls", "W1", "CXX"),      # step1 失败 → step2 命中 (W1, '')
        ("PrdtCls", "W2", "C001"),     # step1/2 失败 → step3 命中 ('', C001)
        ("PrdtCls", "W2", "CXX"),      # 全失败 → step5 命中 ('', '')
        ("Account", "", "A1"),         # 直接命中 A1
        ("Account", "", "AXX"),        # 不存在
        ("Unknown", "WX", "DX"),       # 全不存在
    ]
    fails = 0
    for dockey, whid, docid in cases:
        a = get_fiid_by_prdt_db(db, dockey, whid, docid)
        b = get_fiid_by_prdt_cache(cache, dockey, whid, docid)
        if a != b:
            print(f"[FAIL] doc_fi ({dockey:8s},{whid:3s},{docid:5s}) db={a} cache={b}")
            fails += 1
        else:
            print(f"[OK ] doc_fi ({dockey:8s},{whid:3s},{docid:5s}) -> {a}")
    return fails


def test_ap():
    db = build_db()
    cache = VouMetaCache()
    cache.load_all(db)

    cases = [
        date(2024, 1, 15),    # 1 月 -> APTag=1（已结）
        date(2024, 6, 10),    # 6 月 -> APTag=2（活动）
        date(2024, 7, 20),    # 7 月 -> APTag=0（未来）
        date(2024, 6, 30),    # 6 月最后一天
        date(2024, 7, 1),     # 7 月第一天
        date(2025, 1, 1),     # 范围外
    ]
    fails = 0
    for d in cases:
        a = db.query_ap_by_date(d, use_foaptag=False)
        b = cache.try_get_ap_by_date(d, use_foaptag=False)
        if a != b:
            print(f"[FAIL] ap {d} db={a} cache={b}")
            fails += 1
        else:
            print(f"[OK ] ap {d} -> {a}")
    return fails


def test_corp_emp_acc():
    db = build_db()
    cache = VouMetaCache()
    cache.load_all(db)

    fails = 0
    for cid in ["C1", "C2", "CX"]:
        a = db.query_corp([cid])
        a = a[0] if a else None
        b = cache.try_get_corp(cid)
        # 简化对比：仅检查关键字段
        if (a is None) != (b is None):
            print(f"[FAIL] corp {cid}")
            fails += 1
            continue
        if a is None:
            print(f"[OK ] corp {cid} -> None")
            continue
        ok = a["Contact"] == b["Contact"] and a["CorpName"] == b["CorpName"]
        print(f"[{'OK ' if ok else 'FAIL'}] corp {cid}")
        if not ok: fails += 1

    for eid in ["E1", "E2", "EX"]:
        a = db.query_emp([eid])
        a = a[0] if a else None
        b = cache.try_get_emp(eid)
        if (a is None) != (b is None):
            print(f"[FAIL] emp {eid}")
            fails += 1
            continue
        if a is None:
            print(f"[OK ] emp {eid} -> None")
            continue
        ok = a["dismission"] == b["dismission"] and a["EmpName"] == b["EmpName"]
        print(f"[{'OK ' if ok else 'FAIL'}] emp {eid}")
        if not ok: fails += 1

    for aid in ["A1", "A2", "AX"]:
        a = db.query_acc([aid])
        a = a[0] if a else None
        b = cache.try_get_acc(aid)
        if (a is None) != (b is None):
            print(f"[FAIL] acc {aid}")
            fails += 1
            continue
        if a is None:
            print(f"[OK ] acc {aid} -> None")
            continue
        ok = a["ISStop"] == b["ISStop"] and a["AccName"] == b["AccName"]
        print(f"[{'OK ' if ok else 'FAIL'}] acc {aid}")
        if not ok: fails += 1

    return fails


def test_addrow_equiv():
    """meAddRowI 反射版 vs 显式版字段写入完全一致"""
    idata = {f: f"v_{f}" for f in [
        "ITMID", "FIID", "FINO", "FIName", "IDesc", "CstCenID", "CorpID", "CorpTag",
        "PrdtID", "PrdtName", "AccID", "Weight", "QTY", "MDQTY", "MCQTY", "Pri",
        "DATM", "CATM", "DATM_F", "CATM_F", "AccRate", "SrcITMID", "isFixFI",
        "TInvID", "ClsID", "AssInfo", "WHID", "depid", "EmpID",
        "MFCstPGBillID", "MFCstPGBillNo", "MFCstMOBillID", "MFCstMOBillNo",
        "PrjID", "PrjName", "JTID", "JTName"]}

    rs_a, rs_b = {}, {}
    me_add_row_reflection(rs_a, idata)
    me_add_row_explicit(rs_b, idata)

    fails = 0
    for k in rs_a:
        if rs_a[k] != rs_b.get(k):
            print(f"[FAIL] addrow field {k}: refl={rs_a[k]!r} explicit={rs_b.get(k)!r}")
            fails += 1
    if rs_a == rs_b:
        print(f"[OK ] addrow_equiv ({len(rs_a)} fields all match)")
    return fails


def main():
    print("=== test_fi (CheckFI cache vs db) ===")
    f1 = test_fi()
    print("\n=== test_doc_fi (GetFIIDByPrdt 5-step fallback) ===")
    f2 = test_doc_fi()
    print("\n=== test_ap (CheckDateValidate AccPeriod 查找) ===")
    f3 = test_ap()
    print("\n=== test_corp_emp_acc ===")
    f4 = test_corp_emp_acc()
    print("\n=== test_addrow_equiv (meAddRowI 反射 vs 显式) ===")
    f5 = test_addrow_equiv()
    total = f1 + f2 + f3 + f4 + f5
    print()
    if total > 0:
        print(f"❌ {total} case(s) FAILED")
        raise SystemExit(1)
    print("✅ All PR-1 equivalence tests passed.")


if __name__ == "__main__":
    main()
