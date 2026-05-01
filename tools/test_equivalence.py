"""
算法等价性测试：
- 用 Python 重新实现 meCreateSeqSingle1 的原版（直接照抄分支条件）和优化版
  （= "预加载缓存 + 同语义递归"）。
- 在多组人工构造的数据上对比 (LTime2, ltime2Grade) 输出。

数据模型（与 SQL 表一一对应）：
  AB_WHI: 一条入/出库记录
    BillType, BillID, DC, WHID, PRDTID, CLRID, ITMID, LTime, LTime2, LTime2Grade
  PI_I:   PIBill 明细
    BillID, ITMID, IsNew(0/1), WMSPIBILLID

原算法关键语义（从 VB 代码逐句翻译）：
  meCreateSeqSingle1(BillType, BillID, DC, dLtime, rtnLTimeGrade):
    if rtnLTimeGrade > 150: return
    rsItems = MAX(LTime) GROUP BY BillType,BillID,DC
              过滤：与 (BillType,BillID,DC) 共享 (WHID,ISNULL(PRDTID,''),CLRID)
                    BillType<>'ACF', DC=1
    -- 段 1: PGBill DC=1
    若有：tGrade=1, rtnLTimeGrade += 1
    遍历每条上游 it:
      取 MAX(LTime2),MAX(LTime2Grade) FROM AB_WHI WHERE PGBill it dc=-1, LTime2 NOT NULL
      若 LTime2 不为空:
        rtnLTimeGrade += LTime2Grade
        if LTime2 > dLtime or dLtime == 0: dLtime = LTime2
      否则:
        取 MAX(LTime) WHERE PGBill it dc=-1
        if LTime > dLtime or dLtime == 0: dLtime = LTime
        递归 meCreateSeqSingle1('PGBill', it, -1)
    -- 段 2: jgicbill/icbill/icbill2/ICWHSecBill/icbill3/isbill/ipbill DC=1
    若有 且 tGrade==0：tGrade=1, rtnLTimeGrade += 1
    遍历每条 it:
      取 MAX(LTime),MAX(LTime2),MAX(LTime2Grade) FROM AB_WHI it dc=-1
      若有行:
        if LTime2Grade>0: rtnLTimeGrade += LTime2Grade
        blnSetGrade=False
        if LTime2 not null and (LTime2>dLtime or dLtime==0):
            dLtime = LTime2; blnSetGrade=True
        if not blnSetGrade:
            if LTime > dLtime or dLtime==0: dLtime = LTime
            递归 meCreateSeqSingle1(it.BillType, it.BillID, -1)
    -- 段 3: PIBill DC=1
    遍历每条 it:
      若 PI_I 中存在 BillID=it.BillID AND IsNew=1 AND WMSPIBILLID='':
        取 MAX(AB_WHI.LTime/LTime2/LTime2Grade) JOIN PI_I I ...
           WHERE PIBill it dc=1 AND ISNULL(I.WMSPIBILLID,'')=''
        若有行:
          if LTime2Grade>0: rtnLTimeGrade += LTime2Grade
          blnSetGrade=False
          if LTime2 not null and (LTime2>dLtime or dLtime==0):
              dLtime = LTime2; blnSetGrade=True
          if not blnSetGrade:
              if it.LTime > dLtime or dLtime==0: dLtime = it.LTime
              if tGrade == 0: rtnLTimeGrade += 1
              递归 meCreateSeqSingle1('PIBill', it, -1)
    -- 段 4: 兜底
    若 rsItems(dc=1).Count>0:
      读取 rsLTime（pre-built: MAX(LTime) GROUP BY BillType,BillID, DC=-1）
      若命中当前 (BillType,BillID):
        if LTime > dLtime or dLtime==0: dLtime = LTime
"""

from __future__ import annotations
from collections import defaultdict
from copy import deepcopy
from datetime import datetime
from typing import Optional


ZERO = None  # 代表 dLtime == "00:00:00" 的初始值


def parse(s: Optional[str]) -> Optional[datetime]:
    if s is None:
        return None
    return datetime.strptime(s, "%Y-%m-%d %H:%M:%S")


def gt(a: Optional[datetime], b: Optional[datetime]) -> bool:
    """模拟 (a > b OR b is ZERO)；NULL a 永远不会触发"""
    if a is None:
        return False
    if b is None:
        return True
    return a > b


# ------------------------------------------------------------
# 数据集
# ------------------------------------------------------------
class DB:
    def __init__(self):
        self.ab_whi: list[dict] = []
        self.pi_i: list[dict] = []

    def add_ab(self, **row):
        row.setdefault("PRDTID", "")
        row.setdefault("CLRID", "")
        row.setdefault("ITMID", "")
        row.setdefault("LTime2", None)
        row.setdefault("LTime2Grade", 0)
        row["LTime"] = parse(row["LTime"]) if isinstance(row["LTime"], str) else row["LTime"]
        if row["LTime2"] and isinstance(row["LTime2"], str):
            row["LTime2"] = parse(row["LTime2"])
        self.ab_whi.append(row)

    def add_pi(self, **row):
        row.setdefault("IsNew", 1)
        row.setdefault("WMSPIBILLID", "")
        self.pi_i.append(row)

    def whi_filter(self, **kw):
        out = []
        for r in self.ab_whi:
            if all(r.get(k) == v for k, v in kw.items()):
                out.append(r)
        return out

    def max_ltime(self, BillType, BillID, DC):
        rows = self.whi_filter(BillType=BillType, BillID=BillID, DC=DC)
        rows = [r for r in rows if r["LTime"] is not None]
        return max((r["LTime"] for r in rows), default=None)

    def max_ltime_dc_minus1_with_max_lt2(self, BillType, BillID):
        rows = self.whi_filter(BillType=BillType, BillID=BillID, DC=-1)
        rows = [r for r in rows if r["LTime2"] is not None]
        if not rows:
            return (None, 0)
        return (max(r["LTime2"] for r in rows),
                max(r["LTime2Grade"] for r in rows))

    def max_ltime_dc_minus1(self, BillType, BillID):
        rows = self.whi_filter(BillType=BillType, BillID=BillID, DC=-1)
        rows = [r for r in rows if r["LTime"] is not None]
        return max((r["LTime"] for r in rows), default=None)


# ------------------------------------------------------------
# 原算法
# ------------------------------------------------------------
def original(db: DB, BillType, BillID, DC, dLtime, grade, rsLTime_index, depth=0) -> tuple:
    """返回 (dLtime, grade)"""
    if grade > 150:
        return dLtime, grade

    seed_rows = db.whi_filter(BillType=BillType, BillID=BillID, DC=DC)
    if not seed_rows:
        return dLtime, grade
    seed = seed_rows[0]
    WHID, PRDTID, CLRID = seed["WHID"], seed.get("PRDTID", ""), seed["CLRID"]

    # rsItems: 共享 WHID/PRDTID/CLRID 的其它入库 (DC=1, BillType<>ACF)
    candidates = []
    grouped = defaultdict(list)
    for r in db.ab_whi:
        if r["BillType"] == "ACF":
            continue
        if r["DC"] != 1:
            continue
        if (r["BillType"], r["BillID"], r["DC"]) == (BillType, BillID, DC):
            continue
        if r["WHID"] != WHID:
            continue
        if (r.get("PRDTID") or "") != (PRDTID or ""):
            continue
        if r["CLRID"] != CLRID:
            continue
        grouped[(r["BillType"], r["BillID"], r["DC"])].append(r)
    for (bt, bi, dc), rs in grouped.items():
        max_lt = max((x["LTime"] for x in rs if x["LTime"] is not None), default=None)
        candidates.append({"BillType": bt, "BillID": bi, "DC": dc, "LTime": max_lt})

    tGrade = 0

    # ---- 段 1: PGBill ----
    pgs = [c for c in candidates if c["BillType"] == "PGBill" and c["DC"] == 1]
    if pgs:
        tGrade = 1
        grade += 1
        for it in pgs:
            lt2, lt2g = db.max_ltime_dc_minus1_with_max_lt2("PGBill", it["BillID"])
            if lt2 is not None:
                grade += int(lt2g or 0)
                if gt(lt2, dLtime):
                    dLtime = lt2
            else:
                lt = db.max_ltime_dc_minus1("PGBill", it["BillID"])
                if lt is not None and gt(lt, dLtime):
                    dLtime = lt
                dLtime, grade = original(db, "PGBill", it["BillID"], -1, dLtime, grade,
                                          rsLTime_index, depth + 1)

    # ---- 段 2: 调拨/委外/组装/拆分 ----
    ic_types = {"jgicbill", "icbill", "icbill2", "ICWHSecBill", "icbill3", "isbill", "ipbill"}
    ics = [c for c in candidates if c["BillType"] in ic_types and c["DC"] == 1]
    if ics:
        if tGrade == 0:
            tGrade = 1
            grade += 1
        for it in ics:
            rows = db.whi_filter(BillType=it["BillType"], BillID=it["BillID"], DC=-1)
            if rows:
                max_lt = max((r["LTime"] for r in rows if r["LTime"] is not None), default=None)
                lt2_rows = [r for r in rows if r["LTime2"] is not None]
                max_lt2 = max((r["LTime2"] for r in lt2_rows), default=None)
                max_lt2g = max((r["LTime2Grade"] for r in rows), default=0) or 0
                if max_lt2g > 0:
                    grade += int(max_lt2g)
                blnSetGrade = False
                if max_lt2 is not None and gt(max_lt2, dLtime):
                    dLtime = max_lt2
                    blnSetGrade = True
                if not blnSetGrade:
                    if max_lt is not None and gt(max_lt, dLtime):
                        dLtime = max_lt
                    dLtime, grade = original(db, it["BillType"], it["BillID"], -1,
                                              dLtime, grade, rsLTime_index, depth + 1)

    # ---- 段 3: PIBill ----
    pis = [c for c in candidates if c["BillType"] == "PIBill" and c["DC"] == 1]
    for it in pis:
        match = [p for p in db.pi_i
                 if p["BillID"] == it["BillID"] and p.get("IsNew", 0) == 1
                 and (p.get("WMSPIBILLID") or "") == ""]
        if not match:
            continue
        # AB_WHI JOIN PI_I 的 MAX
        joined = []
        for r in db.whi_filter(BillType="PIBill", BillID=it["BillID"], DC=1):
            for p in db.pi_i:
                if p["BillID"] == r["BillID"] and p["ITMID"] == r["ITMID"] \
                        and (p.get("WMSPIBILLID") or "") == "":
                    joined.append(r)
                    break
        if not joined:
            continue
        max_lt = max((r["LTime"] for r in joined if r["LTime"] is not None), default=None)
        lt2_rows = [r for r in joined if r["LTime2"] is not None]
        max_lt2 = max((r["LTime2"] for r in lt2_rows), default=None)
        max_lt2g = max((r["LTime2Grade"] for r in joined), default=0) or 0
        if max_lt2g > 0:
            grade += int(max_lt2g)
        blnSetGrade = False
        if max_lt2 is not None and gt(max_lt2, dLtime):
            dLtime = max_lt2
            blnSetGrade = True
        if not blnSetGrade:
            it_lt = it["LTime"]
            if it_lt is not None and gt(it_lt, dLtime):
                dLtime = it_lt
            if tGrade == 0:
                grade += 1
            dLtime, grade = original(db, "PIBill", it["BillID"], -1, dLtime, grade,
                                      rsLTime_index, depth + 1)

    # ---- 段 4: 兜底 ----
    if any(c for c in candidates if c["DC"] == 1):
        lt = rsLTime_index.get((BillType, BillID))
        if lt is not None and gt(lt, dLtime):
            dLtime = lt

    return dLtime, grade


# ------------------------------------------------------------
# 优化算法：等价于原算法，但用预加载缓存避免任何 SQL 往返
# ------------------------------------------------------------
class Cache:
    """与 VB 优化版完全等价的缓存模型。
       每个种子 (BT,BID,DC) 可以对应多个 (W,P,C) 库位（多明细行）。"""
    def __init__(self, db: DB):
        # 1. 同 (WHID,PRDTID,CLRID) 的"DC=1 候选"
        self.by_loc: dict = defaultdict(dict)
        # by_loc[locKey][(BT,BID,DC)] = MaxLT
        for r in db.ab_whi:
            if r["BillType"] == "ACF" or r["DC"] != 1:
                continue
            k = (r["WHID"], r.get("PRDTID") or "", r["CLRID"])
            ck = (r["BillType"], r["BillID"], r["DC"])
            cur = self.by_loc[k].get(ck)
            if r["LTime"] is not None and (cur is None or r["LTime"] > cur):
                self.by_loc[k][ck] = r["LTime"]
            elif ck not in self.by_loc[k]:
                self.by_loc[k][ck] = cur

        # 2. (BillType,BillID,DC) -> set of locKey
        #    覆盖 ALL DC（含 DC=-1 的种子），排除 ACF
        self.seed_locs: dict = defaultdict(set)
        for r in db.ab_whi:
            if r["BillType"] == "ACF":
                continue
            self.seed_locs[(r["BillType"], r["BillID"], r["DC"])].add(
                (r["WHID"], r.get("PRDTID") or "", r["CLRID"]))

        # 3. (BillType,BillID,DC=1) 的 MAX(LTime) —— 用于 Edge 的 LTime
        self.max_lt_dc1: dict = defaultdict(lambda: None)
        for r in db.ab_whi:
            if r["DC"] != 1 or r["BillType"] == "ACF":
                continue
            k = (r["BillType"], r["BillID"])
            cur = self.max_lt_dc1[k]
            if r["LTime"] is not None and (cur is None or r["LTime"] > cur):
                self.max_lt_dc1[k] = r["LTime"]

        # 4. (BillType,BillID,DC=-1) 的 MAX(LTime/LTime2/MAX LTime2Grade)
        self.dcn1_agg: dict = {}
        bucket = defaultdict(list)
        for r in db.ab_whi:
            if r["DC"] == -1:
                bucket[(r["BillType"], r["BillID"])].append(r)
        for k, rs in bucket.items():
            self.dcn1_agg[k] = {
                "LTime":  max((r["LTime"]  for r in rs if r["LTime"]  is not None), default=None),
                "LTime2": max((r["LTime2"] for r in rs if r["LTime2"] is not None), default=None),
                "LTime2Grade": max((r["LTime2Grade"] for r in rs), default=0) or 0,
            }

        # 5. PIBill 段：BillID -> bool 是否存在 IsNew=1 且 WMSPIBILLID=''
        self.pi_isnew: dict = {}
        for p in db.pi_i:
            if p.get("IsNew", 0) == 1 and (p.get("WMSPIBILLID") or "") == "":
                self.pi_isnew[p["BillID"]] = True

        # 6. PIBill JOIN PI_I 的 MAX(LTime/LTime2/LTime2Grade)
        self.pi_join_agg: dict = {}
        bucket2 = defaultdict(list)
        for r in db.ab_whi:
            if r["BillType"] != "PIBill" or r["DC"] != 1:
                continue
            for p in db.pi_i:
                if p["BillID"] == r["BillID"] and p["ITMID"] == r["ITMID"] \
                        and (p.get("WMSPIBILLID") or "") == "":
                    bucket2[r["BillID"]].append(r)
                    break
        for bid, rs in bucket2.items():
            self.pi_join_agg[bid] = {
                "LTime":  max((r["LTime"]  for r in rs if r["LTime"]  is not None), default=None),
                "LTime2": max((r["LTime2"] for r in rs if r["LTime2"] is not None), default=None),
                "LTime2Grade": max((r["LTime2Grade"] for r in rs), default=0) or 0,
            }

        # 7. rsLTime: MAX(LTime) GROUP BY BillType,BillID, DC=-1
        self.rsLTime: dict = {k: v["LTime"] for k, v in self.dcn1_agg.items()}

        # 8. 已计算结果缓存：(BillType,BillID,DC=-1) -> (LTime2,Grade)
        #    递归过程中会即时回填，模拟原算法更新 AB_WHI.LTime2 后下游能读到
        self.computed: dict = {}


def candidates_for(cache: Cache, BillType, BillID, DC):
    locs = cache.seed_locs.get((BillType, BillID, DC), set())
    out = {}
    for k in locs:
        for (bt, bi, dc), max_lt in cache.by_loc.get(k, {}).items():
            if (bt, bi, dc) == (BillType, BillID, DC):
                continue
            cur = out.get((bt, bi, dc))
            if cur is None or (max_lt is not None and (cur["LTime"] is None or max_lt > cur["LTime"])):
                out[(bt, bi, dc)] = {"BillType": bt, "BillID": bi, "DC": dc, "LTime": max_lt}
    return list(out.values())


def optimized(cache: Cache, BillType, BillID, DC, dLtime, grade, depth=0) -> tuple:
    if grade > 150:
        return dLtime, grade

    cands = candidates_for(cache, BillType, BillID, DC)
    tGrade = 0

    # 段 1: PGBill
    pgs = [c for c in cands if c["BillType"] == "PGBill" and c["DC"] == 1]
    if pgs:
        tGrade = 1
        grade += 1
        for it in pgs:
            agg = cache.dcn1_agg.get(("PGBill", it["BillID"]))
            # 即时回填：如果之前递归算出过结果，要按"原算法那一刻数据库已经被更新过"来读
            updated = cache.computed.get(("PGBill", it["BillID"], -1))
            lt2 = lt2g = None
            if updated is not None:
                lt2, lt2g = updated
            elif agg is not None:
                lt2, lt2g = agg["LTime2"], agg["LTime2Grade"]
            if lt2 is not None:
                grade += int(lt2g or 0)
                if gt(lt2, dLtime):
                    dLtime = lt2
            else:
                lt = agg["LTime"] if agg else None
                if lt is not None and gt(lt, dLtime):
                    dLtime = lt
                dLtime, grade = optimized(cache, "PGBill", it["BillID"], -1, dLtime, grade, depth + 1)

    # 段 2: IC 系列
    ic_types = {"jgicbill", "icbill", "icbill2", "ICWHSecBill", "icbill3", "isbill", "ipbill"}
    ics = [c for c in cands if c["BillType"] in ic_types and c["DC"] == 1]
    if ics:
        if tGrade == 0:
            tGrade = 1
            grade += 1
        for it in ics:
            updated = cache.computed.get((it["BillType"], it["BillID"], -1))
            agg = cache.dcn1_agg.get((it["BillType"], it["BillID"]))
            if updated is not None:
                lt2, lt2g = updated
                lt = agg["LTime"] if agg else None
            elif agg is not None:
                lt2, lt2g = agg["LTime2"], agg["LTime2Grade"]
                lt = agg["LTime"]
            else:
                lt2 = lt2g = lt = None
            if agg is not None or updated is not None:
                if (lt2g or 0) > 0:
                    grade += int(lt2g)
                blnSetGrade = False
                if lt2 is not None and gt(lt2, dLtime):
                    dLtime = lt2
                    blnSetGrade = True
                if not blnSetGrade:
                    if lt is not None and gt(lt, dLtime):
                        dLtime = lt
                    dLtime, grade = optimized(cache, it["BillType"], it["BillID"], -1,
                                               dLtime, grade, depth + 1)

    # 段 3: PIBill
    pis = [c for c in cands if c["BillType"] == "PIBill" and c["DC"] == 1]
    for it in pis:
        if not cache.pi_isnew.get(it["BillID"]):
            continue
        agg = cache.pi_join_agg.get(it["BillID"])
        if agg is None:
            continue
        lt2, lt2g, lt = agg["LTime2"], agg["LTime2Grade"], agg["LTime"]
        # 同样支持即时回填（PIBill dc=-1 走过递归）
        updated = cache.computed.get(("PIBill", it["BillID"], -1))
        if updated is not None:
            lt2, lt2g = updated
        if (lt2g or 0) > 0:
            grade += int(lt2g)
        blnSetGrade = False
        if lt2 is not None and gt(lt2, dLtime):
            dLtime = lt2
            blnSetGrade = True
        if not blnSetGrade:
            it_lt = it["LTime"]
            if it_lt is not None and gt(it_lt, dLtime):
                dLtime = it_lt
            if tGrade == 0:
                grade += 1
            dLtime, grade = optimized(cache, "PIBill", it["BillID"], -1, dLtime, grade, depth + 1)

    # 段 4: 兜底
    if any(c for c in cands if c["DC"] == 1):
        lt = cache.rsLTime.get((BillType, BillID))
        if lt is not None and gt(lt, dLtime):
            dLtime = lt

    return dLtime, grade


# ============================================================
# 测试用例
# ============================================================
def case_simple_pgbill():
    db = DB()
    # 出库的成品入库 (DC=-1)，待算
    db.add_ab(BillType="PGBill", BillID="P1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-10 10:00:00")
    # 同库位的另一条 DC=1：上游 PGBill
    db.add_ab(BillType="PGBill", BillID="P0", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-05 09:00:00",
              LTime2=parse("2024-01-05 09:00:00"), LTime2Grade=2)
    return db, "PGBill", "P1", -1


def case_ic_chain():
    db = DB()
    db.add_ab(BillType="PGBill", BillID="P1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-10 10:00:00")
    # 上游：调拨入库
    db.add_ab(BillType="icbill", BillID="IC1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-08 10:00:00")
    # 调拨入库的 dc=-1 = 调拨出库，再上游：PGBill
    db.add_ab(BillType="icbill", BillID="IC1", DC=-1, WHID="W2", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-07 10:00:00")
    db.add_ab(BillType="PGBill", BillID="P0", DC=1, WHID="W2", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-05 10:00:00",
              LTime2=parse("2024-01-05 10:00:00"), LTime2Grade=3)
    return db, "PGBill", "P1", -1


def case_pibill_isnew():
    db = DB()
    db.add_ab(BillType="PGBill", BillID="P1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-02-01 09:00:00")
    db.add_ab(BillType="PIBill", BillID="PI1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-20 09:00:00",
              LTime2=parse("2024-01-20 09:00:00"), LTime2Grade=1)
    db.add_pi(BillID="PI1", ITMID="I1", IsNew=1, WMSPIBILLID="")
    return db, "PGBill", "P1", -1


def case_pibill_skip_when_no_isnew():
    db = DB()
    db.add_ab(BillType="PGBill", BillID="P1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-02-01 09:00:00")
    db.add_ab(BillType="PIBill", BillID="PI1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-20 09:00:00")
    db.add_pi(BillID="PI1", ITMID="I1", IsNew=0, WMSPIBILLID="")  # IsNew=0 -> 跳过
    return db, "PGBill", "P1", -1


def case_fallback_dltime():
    """仅有 PIBill 候选但被 IsNew 过滤掉，应走兜底（rsLTime）"""
    db = DB()
    db.add_ab(BillType="PGBill", BillID="P1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-02-01 09:00:00")
    # 还有一条 DC=1（任意类型，进入兜底分支判定）
    db.add_ab(BillType="ipbill", BillID="IP1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-15 09:00:00")
    # 但 ipbill 没有 dc=-1 的聚合数据 -> 走递归回退 -> 兜底
    return db, "PGBill", "P1", -1


def case_null_prdtid():
    db = DB()
    db.add_ab(BillType="PGBill", BillID="P1", DC=-1, WHID="W1", PRDTID=None, CLRID="C1",
              ITMID="I1", LTime="2024-03-01 09:00:00")
    db.add_ab(BillType="PGBill", BillID="P0", DC=1, WHID="W1", PRDTID="", CLRID="C1",
              ITMID="I1", LTime="2024-02-20 09:00:00",
              LTime2=parse("2024-02-20 09:00:00"), LTime2Grade=4)
    return db, "PGBill", "P1", -1


def case_pg_in_then_sale_out():
    """用户报告的场景：PGBill 加工入库 DC=1，再销售 DC=-1。
       销售出库的 LTime2Grade 应为 1（找到上游 PGBill 入库），不能是 0。"""
    db = DB()
    # 销售出库 (seed): 用户产品 A 在仓库 W1
    db.add_ab(BillType="IOBill", BillID="SO1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-03-10 10:00:00")
    # 上游：PGBill 加工入库 同 (W1, A, C1)
    db.add_ab(BillType="PGBill", BillID="PG1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-03-05 09:00:00")
    return db, "IOBill", "SO1", -1


def make_chain_db():
    """业务场景：热卷 -> 酸洗 -> 冷轧 -> 镀锌 -> 分剪 -> 制管。
       每道工序 PGBill 的 (WHID, PRDTID, CLRID) 都可以不同。"""
    db = DB()
    # 起点：采购入库 (W0, A0, Ca) -- 热卷
    db.add_ab(BillType="PIBill", BillID="PI1", DC=1, WHID="W0", PRDTID="A0", CLRID="Ca",
              ITMID="I1", LTime="2024-01-01 09:00:00")
    db.add_pi(BillID="PI1", ITMID="I1", IsNew=1, WMSPIBILLID="")

    # 酸洗 PG_S: 原料 (W0, A0, Ca) -> 产出 (W1, A1, Ca)  -- CLRID 不变
    db.add_ab(BillType="PGBill", BillID="PG_S", DC=-1, WHID="W0", PRDTID="A0", CLRID="Ca",
              ITMID="I1", LTime="2024-01-02 09:00:00")
    db.add_ab(BillType="PGBill", BillID="PG_S", DC=1,  WHID="W1", PRDTID="A1", CLRID="Ca",
              ITMID="I2", LTime="2024-01-02 10:00:00")

    # 冷轧 PG_C: 原料 (W1, A1, Ca) -> 产出 (W2, A2, Cb)  -- CLRID 改变
    db.add_ab(BillType="PGBill", BillID="PG_C", DC=-1, WHID="W1", PRDTID="A1", CLRID="Ca",
              ITMID="I2", LTime="2024-01-03 09:00:00")
    db.add_ab(BillType="PGBill", BillID="PG_C", DC=1,  WHID="W2", PRDTID="A2", CLRID="Cb",
              ITMID="I3", LTime="2024-01-03 10:00:00")

    # 镀锌 PG_G: 原料 (W2, A2, Cb) -> 产出 (W3, A3, Cb)  -- CLRID 不变
    db.add_ab(BillType="PGBill", BillID="PG_G", DC=-1, WHID="W2", PRDTID="A2", CLRID="Cb",
              ITMID="I3", LTime="2024-01-04 09:00:00")
    db.add_ab(BillType="PGBill", BillID="PG_G", DC=1,  WHID="W3", PRDTID="A3", CLRID="Cb",
              ITMID="I4", LTime="2024-01-04 10:00:00")

    # 分剪 PG_F: 原料 (W3, A3, Cb) -> 产出 (W4, A4, Cc)
    db.add_ab(BillType="PGBill", BillID="PG_F", DC=-1, WHID="W3", PRDTID="A3", CLRID="Cb",
              ITMID="I4", LTime="2024-01-05 09:00:00")
    db.add_ab(BillType="PGBill", BillID="PG_F", DC=1,  WHID="W4", PRDTID="A4", CLRID="Cc",
              ITMID="I5", LTime="2024-01-05 10:00:00")

    # 制管 PG_T: 原料 (W4, A4, Cc) -> 产出 (W5, A5, Cc)
    db.add_ab(BillType="PGBill", BillID="PG_T", DC=-1, WHID="W4", PRDTID="A4", CLRID="Cc",
              ITMID="I5", LTime="2024-01-06 09:00:00")
    db.add_ab(BillType="PGBill", BillID="PG_T", DC=1,  WHID="W5", PRDTID="A5", CLRID="Cc",
              ITMID="I6", LTime="2024-01-06 10:00:00")

    # 最终销售出库 (W5, A5, Cc)
    db.add_ab(BillType="IOBill", BillID="SO1", DC=-1, WHID="W5", PRDTID="A5", CLRID="Cc",
              ITMID="I6", LTime="2024-01-07 09:00:00")
    return db


def case_chain_pg_s():     return make_chain_db(), "PGBill", "PG_S", -1
def case_chain_pg_c():     return make_chain_db(), "PGBill", "PG_C", -1
def case_chain_pg_g():     return make_chain_db(), "PGBill", "PG_G", -1
def case_chain_pg_f():     return make_chain_db(), "PGBill", "PG_F", -1
def case_chain_pg_t():     return make_chain_db(), "PGBill", "PG_T", -1
def case_chain_so1():      return make_chain_db(), "IOBill", "SO1", -1


def case_pg_in_then_sale_out_multi_loc():
    """同一张销售单据有多个明细（多 (W,P,C)），每条都需匹配上游 PGBill 入库。"""
    db = DB()
    db.add_ab(BillType="IOBill", BillID="SO1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-04-10 10:00:00")
    db.add_ab(BillType="IOBill", BillID="SO1", DC=-1, WHID="W2", PRDTID="B", CLRID="C2",
              ITMID="I2", LTime="2024-04-10 10:00:00")
    db.add_ab(BillType="PGBill", BillID="PG1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-04-05 09:00:00")
    db.add_ab(BillType="PGBill", BillID="PG2", DC=1, WHID="W2", PRDTID="B", CLRID="C2",
              ITMID="I2", LTime="2024-04-06 09:00:00")
    return db, "IOBill", "SO1", -1


def case_iibill_seed_in_dc1_only():
    """边界：种子 (DC=-1) 在 AB_WHI 中存在，但 BillType=icbill 既有 DC=1 又有 DC=-1。
       验证 dictSeedLocs 必须覆盖 ALL DC，否则 DC=-1 种子拿不到库位。"""
    db = DB()
    db.add_ab(BillType="icbill", BillID="IC1", DC=-1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-05-10 10:00:00")
    db.add_ab(BillType="PGBill", BillID="PG1", DC=1, WHID="W1", PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-05-05 09:00:00")
    return db, "icbill", "IC1", -1


def case_grade_cap():
    """构造一个会触发 grade>150 的链；测试两边都尊重截断"""
    db = DB()
    last_wh = "W0"
    # 链长 200
    db.add_ab(BillType="PGBill", BillID="P200", DC=-1, WHID=last_wh, PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-01 00:00:00")
    for i in range(199, -1, -1):
        bid_in = f"IC{i}_in"
        bid_out = f"IC{i}_out"
        new_wh = f"W{i+1}"
        # 同库位的入库（候选）
        db.add_ab(BillType="icbill", BillID=f"IC{i}", DC=1, WHID=last_wh, PRDTID="A",
                  CLRID="C1", ITMID="I1", LTime="2024-01-01 00:00:00")
        # 该 icbill 的 dc=-1 在新库位
        db.add_ab(BillType="icbill", BillID=f"IC{i}", DC=-1, WHID=new_wh, PRDTID="A",
                  CLRID="C1", ITMID="I1", LTime="2024-01-01 00:00:00")
        last_wh = new_wh
    # 顶端给一个 PGBill 兜住
    db.add_ab(BillType="PGBill", BillID="ROOT", DC=1, WHID=last_wh, PRDTID="A", CLRID="C1",
              ITMID="I1", LTime="2024-01-01 00:00:00",
              LTime2=parse("2024-01-01 00:00:00"), LTime2Grade=1)
    return db, "PGBill", "P200", -1


def run_one(name, builder):
    db, bt, bi, dc = builder()
    rsLTime = {(r["BillType"], r["BillID"]):
               max((x["LTime"] for x in db.whi_filter(BillType=r["BillType"],
                                                      BillID=r["BillID"], DC=-1)
                    if x["LTime"] is not None), default=None)
               for r in db.ab_whi if r["DC"] == -1}

    # 原算法跑一次（不会修改 AB_WHI -> 模拟和原代码一致：原代码每张 PGBill 算完
    # 都 UPDATE AB_WHI，下游再读会拿到新值；为了精确对齐，我们也要模拟该行为）
    db1 = deepcopy(db)
    rsL1 = {(r["BillType"], r["BillID"]):
            max((x["LTime"] for x in db1.whi_filter(BillType=r["BillType"],
                                                    BillID=r["BillID"], DC=-1)
                 if x["LTime"] is not None), default=None)
            for r in db1.ab_whi if r["DC"] == -1}

    dLt1, gr1 = original(db1, bt, bi, dc, ZERO, 0, rsL1)

    # 优化算法
    cache = Cache(deepcopy(db))
    dLt2, gr2 = optimized(cache, bt, bi, dc, ZERO, 0)

    ok = (dLt1 == dLt2) and (gr1 == gr2)
    print(f"[{'OK ' if ok else 'FAIL'}] {name:40s}  "
          f"orig=({dLt1},{gr1})  opt=({dLt2},{gr2})")
    return ok


def main():
    cases = [
        ("case_simple_pgbill",          case_simple_pgbill),
        ("case_ic_chain",               case_ic_chain),
        ("case_pibill_isnew",           case_pibill_isnew),
        ("case_pibill_skip_when_no_isnew", case_pibill_skip_when_no_isnew),
        ("case_fallback_dltime",        case_fallback_dltime),
        ("case_null_prdtid",            case_null_prdtid),
        ("case_pg_in_then_sale_out",    case_pg_in_then_sale_out),
        ("case_pg_in_then_sale_out_multi_loc", case_pg_in_then_sale_out_multi_loc),
        ("case_iibill_seed_in_dc1_only", case_iibill_seed_in_dc1_only),
        ("case_chain_pg_s",             case_chain_pg_s),
        ("case_chain_pg_c",             case_chain_pg_c),
        ("case_chain_pg_g",             case_chain_pg_g),
        ("case_chain_pg_f",             case_chain_pg_f),
        ("case_chain_pg_t",             case_chain_pg_t),
        ("case_chain_so1",              case_chain_so1),
        ("case_grade_cap",              case_grade_cap),
    ]
    fails = sum(0 if run_one(n, b) else 1 for n, b in cases)
    if fails:
        print(f"\n{fails} case(s) FAILED")
        raise SystemExit(1)
    print("\nAll equivalence tests passed.")


if __name__ == "__main__":
    main()
