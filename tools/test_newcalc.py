"""
新成本核算分支等价性测试。

被测语义（原 VB 代码 getTreeLevel1 / getTreeLevelSub，按用户要求 curid/upid 都改为 PRDTID）：

  Function getTreeLevel1(W, P, C):
      uppers = SELECT DISTINCT upid FROM #Tmp_LTime2Grade
                WHERE whid=W AND prdtid=P AND ISNULL(clrid,'')=C
                GROUP BY upid
      if not uppers: return 0
      level = 1
      for u in uppers:
          TRUNCATE #Tmp_Tree                        # 兄弟之间 reset visited
          INSERT (W-P-C, 1) into #Tmp_Tree
          tLevel = 1                                 # 本地变量
          getTreeLevelSub(u.W, u.P, u.C, ByRef tLevel)   # 注意：tLevel ByRef
          if tLevel > level: level = tLevel
      return level

  Sub getTreeLevelSub(W, P, C, ByRef Level):
      INSERT (W-P-C, 1) into #Tmp_Tree              # 标记访问
      uppers = SELECT DISTINCT upid FROM #Tmp_LTime2Grade
                WHERE whid=W AND prdtid=P AND ISNULL(clrid,'')=C
                  AND upid NOT IN (SELECT curid FROM #Tmp_Tree)
                GROUP BY upid
      if uppers is empty: return
      Level = Level + 1
      for u in uppers:
          tLevel = Level                            # 注意：先复制当前最新 Level
          getTreeLevelSub(u, ByRef tLevel)
          if tLevel > Level: Level = tLevel         # 兄弟之间通过 Level 累积

注意：path-shared visited + ByRef Level 累积 -> 同 sub 的多个兄弟之间会"假性深度叠加"。
这是原代码固有特征。本测试 1:1 保留该语义。
"""

from __future__ import annotations
from collections import defaultdict
from typing import Iterable
import random
import sys
import time

sys.setrecursionlimit(50000)


# ============================================================
# 原算法（1:1 翻译 VB 代码）
# ============================================================
def orig_get_tree_level1(edges: dict, root: tuple) -> int:
    """getTreeLevel1(WHID, PRDTID, CLRID) - root 是 (W,P,C) 三元组"""
    uppers = sorted(set(edges.get(root, [])))   # GROUP BY upid -> distinct
    if not uppers:
        return 0
    level = [1]
    for u in uppers:
        # 兄弟之间 TRUNCATE：每次 visited 重置
        visited = {root}
        tlevel = [1]
        orig_get_tree_level_sub(edges, u, visited, tlevel)
        if tlevel[0] > level[0]:
            level[0] = tlevel[0]
    return level[0]


def orig_get_tree_level_sub(edges: dict, node: tuple, visited: set, level: list):
    """getTreeLevelSub(W,P,C, ByRef Level)
       visited / level 都按 ByRef 行为（list 包装 + set 共享）。"""
    visited.add(node)                                          # INSERT into #Tmp_Tree
    uppers = sorted(u for u in set(edges.get(node, []))
                    if u not in visited)                        # NOT IN (SELECT curid FROM #Tmp_Tree)
    if not uppers:
        return
    level[0] = level[0] + 1                                     # Level = Level + 1
    for u in uppers:
        tlevel = [level[0]]                                     # tLevel = Level (按值)
        orig_get_tree_level_sub(edges, u, visited, tlevel)
        if tlevel[0] > level[0]:                                # If tLevel > Level Then Level = tLevel
            level[0] = tlevel[0]


# ============================================================
# 优化算法（A3：内存 DFS + Dictionary，1:1 等价于原算法）
# ============================================================
class TreeLevelCalc:
    """A3 实现：把 #Tmp_LTime2Grade 一次性载入内存，所有 DFS 走 dict。
       行为与原算法逐字段对齐。"""

    def __init__(self, edges: dict):
        self.edges = edges

    def get_level(self, root: tuple) -> int:
        """与 VB meTreeLevel 1:1 等价的 Python 实现。"""
        if root not in self.edges:
            return 0
        root_uppers = sorted(set(u for u in set(self.edges.get(root, []))
                                  if u != root))
        if not root_uppers:
            return 0
        best = 1
        for u in root_uppers:
            visited = {root}
            ret = self._sub_iter(u, visited, 1)
            if ret > best:
                best = ret
        return best

    def _sub_iter(self, start_node, visited, init_level):
        """与 VB meSubIter 1:1 等价：显式栈，并行数组结构。"""
        # root frame 入栈
        def distinct_sorted_uppers(node):
            if node not in visited:
                visited.add(node)
            ups = []
            seen = set()
            for u in self.edges.get(node, []):
                if u != node and u not in visited and u not in seen:
                    seen.add(u)
                    ups.append(u)
            ups.sort()
            return ups

        u0 = distinct_sorted_uppers(start_node)
        if not u0:
            return init_level

        stk_uppers = [u0]
        stk_idx = [0]
        stk_level = [init_level + 1]
        stk_child = [0]
        stk_has_ch = [False]
        sp = 0
        final = init_level

        while sp >= 0:
            if stk_has_ch[sp]:
                if stk_child[sp] > stk_level[sp]:
                    stk_level[sp] = stk_child[sp]
                stk_has_ch[sp] = False

            if stk_idx[sp] > len(stk_uppers[sp]) - 1:
                ret = stk_level[sp]
                sp -= 1
                stk_uppers.pop(); stk_idx.pop(); stk_level.pop()
                stk_child.pop(); stk_has_ch.pop()
                if sp >= 0:
                    stk_child[sp] = ret
                    stk_has_ch[sp] = True
                else:
                    final = ret
            else:
                u = stk_uppers[sp][stk_idx[sp]]
                stk_idx[sp] += 1
                sub_u = distinct_sorted_uppers(u)
                if not sub_u:
                    pass
                else:
                    init = stk_level[sp]
                    stk_uppers.append(sub_u)
                    stk_idx.append(0)
                    stk_level.append(init + 1)
                    stk_child.append(0)
                    stk_has_ch.append(False)
                    sp += 1
        return final


# ============================================================
# 测试用例构造
# ============================================================
def build_chain_edges():
    """钢铁链：a -> b -> c -> d -> e -> f, 每步对应一道工序的 (W,P,C)"""
    A = ("W0", "A0", "")
    B = ("W1", "A1", "Ca")
    C = ("W2", "A2", "Cb")
    D = ("W3", "A3", "Cb")
    E = ("W4", "A4", "Cc")
    F = ("W5", "A5", "Cc")
    edges = defaultdict(list)
    edges[F].append(E)
    edges[E].append(D)
    edges[D].append(C)
    edges[C].append(B)
    edges[B].append(A)
    return dict(edges), [A, B, C, D, E, F]


def build_diamond():
    """菱形：a 既被 b 也被 c 引用，b/c 都引用 d。
       原算法因 path-shared 会把 b/c 计算为不同深度。"""
    A = ("W0", "P0", "")
    B = ("W1", "P1", "")
    C = ("W2", "P2", "")
    D = ("W3", "P3", "")
    edges = {
        D: [B, C],
        B: [A],
        C: [A],
    }
    return edges, [A, B, C, D]


def build_shared_upper():
    """两个根共享一个上游链：a 和 b 都到 x -> y -> z"""
    A = ("W0", "PA", "")
    B = ("W0", "PB", "")
    X = ("W1", "PX", "")
    Y = ("W2", "PY", "")
    Z = ("W3", "PZ", "")
    edges = {
        A: [X],
        B: [X],
        X: [Y],
        Y: [Z],
    }
    return edges, [A, B, X, Y, Z]


def build_with_cycle():
    """有环：a -> b -> c -> a，外加 c -> d。
       原算法靠 visited 在递归中防环，不应死循环。"""
    A = ("W0", "PA", "")
    B = ("W1", "PB", "")
    C = ("W2", "PC", "")
    D = ("W3", "PD", "")
    edges = {
        A: [B],
        B: [C],
        C: [A, D],
    }
    return edges, [A, B, C, D]


def build_null_clrid():
    """CLRID 为空字符串 vs 实际值的对比：保证两者被作为不同节点。"""
    A1 = ("W1", "P1", "")    # CLRID 空
    A2 = ("W1", "P1", "C1")  # CLRID 有值
    B = ("W2", "P2", "")
    edges = {
        A1: [B],
    }
    return edges, [A1, A2, B]


def build_random_dag(n_nodes=200, density=2.5, seed=42):
    """随机 DAG，每个节点平均向上指向 density 个上游（取 id 更大的节点）。"""
    rnd = random.Random(seed)
    nodes = [(f"W{i}", f"P{i}", "") for i in range(n_nodes)]
    edges = defaultdict(list)
    for i, n in enumerate(nodes):
        # 越靠后的节点上游越多（叶子在最前）
        candidates = list(range(i + 1, min(i + 10, n_nodes)))
        if not candidates:
            continue
        k = min(len(candidates), max(1, int(rnd.expovariate(1 / density))))
        for j in rnd.sample(candidates, k):
            edges[n].append(nodes[j])
    return dict(edges), nodes


# ============================================================
# 测试运行
# ============================================================
def assert_equiv(name, edges, nodes):
    calc = TreeLevelCalc(edges)
    fails = []
    for n in nodes:
        a = orig_get_tree_level1(edges, n)
        b = calc.get_level(n)
        if a != b:
            fails.append((n, a, b))
    status = "OK " if not fails else "FAIL"
    print(f"[{status}] {name:30s}  nodes={len(nodes):4d}  fails={len(fails)}")
    if fails:
        for n, a, b in fails[:5]:
            print(f"    node={n}  orig={a}  opt={b}")
    return len(fails) == 0


# ============================================================
# 性能模型：模拟 SQL 实现 vs A3 内存实现的耗时
# ------------------------------------------------------------
# 原 VB 实现的瓶颈是"递归内每层 2 个 SQL，且 SELECT 走全表扫"。
# 我们保守按以下参数模拟（线上实测可能更糟）：
#   - 每条 SQL 的固定网络往返 RPC = 1.5 ms
#   - 每条 SELECT 因索引不匹配走表扫，扫描成本 = 0.5 us × 边表行数
#   - INSERT/TRUNCATE 走 0 额外成本
# A3 内存实现的瓶颈是 dict lookup，每次 0.5 us
# ============================================================
SQL_RPC_MS = 1.5
SQL_SCAN_PER_ROW_US = 0.5
MEM_LOOKUP_US = 0.5


def perf_orig_simulated(edges, nodes):
    """模拟原 SQL 实现：每个 (W,P,C) 跑一次 getTreeLevel1，递归内每层执行 2 条 SQL。
    返回估算耗时（毫秒），并附带 SQL 调用次数 / DFS 节点访问数。"""
    edge_count = sum(len(v) for v in edges.values())
    sql_calls = 0
    visits = 0

    def sub(node, visited, level):
        nonlocal sql_calls, visits
        visited.add(node)
        sql_calls += 1                                          # INSERT into #Tmp_Tree
        # SELECT distinct upid (NOT IN visited)
        sql_calls += 1
        visits += edge_count                                    # 全表扫
        uppers = sorted(u for u in set(edges.get(node, []))
                        if u not in visited)
        if not uppers:
            return
        level[0] += 1
        for u in uppers:
            tlevel = [level[0]]
            sub(u, visited, tlevel)
            if tlevel[0] > level[0]:
                level[0] = tlevel[0]

    for root in nodes:
        sql_calls += 1                                          # SELECT distinct upid (initial)
        visits += edge_count
        uppers = sorted(set(edges.get(root, [])))
        if not uppers:
            continue
        for u in uppers:
            sql_calls += 1                                      # TRUNCATE #Tmp_Tree
            sql_calls += 1                                      # INSERT root into #Tmp_Tree
            visited = {root}
            tlevel = [1]
            sub(u, visited, tlevel)

    elapsed = sql_calls * SQL_RPC_MS + visits * SQL_SCAN_PER_ROW_US / 1000
    return elapsed, sql_calls, visits


def perf_opt_simulated(edges, nodes):
    """模拟 A3 内存实现：1 次 SELECT 拉边表 + N 次 in-memory DFS。"""
    edge_count = sum(len(v) for v in edges.values())
    sql_calls = 1
    initial_load_ms = SQL_RPC_MS + edge_count * SQL_SCAN_PER_ROW_US / 1000

    # 估算 visits：DFS 中 visited.add 的次数（与原算法相同）
    visits = 0

    def count_dfs(root):
        nonlocal visits
        if root not in edges:
            return 0
        root_uppers = sorted(set(u for u in edges.get(root, []) if u != root))
        if not root_uppers:
            return 0
        best = 1
        for u in root_uppers:
            visited = {root}
            stack_uppers, stack_idx, stack_level = [], [], []

            def enter(node, init_lv):
                nonlocal visits
                if node not in visited:
                    visited.add(node)
                    visits += 1
                ups = sorted(set(uu for uu in edges.get(node, [])
                                 if uu != node and uu not in visited))
                if not ups:
                    return None
                return (ups, init_lv + 1)

            r = enter(u, 1)
            if r is None:
                continue
            stack_uppers.append(r[0])
            stack_idx.append(0)
            stack_level.append(r[1])
            local_best = 1
            child_pending = None
            while stack_uppers:
                if child_pending is not None:
                    if child_pending > stack_level[-1]:
                        stack_level[-1] = child_pending
                    child_pending = None
                if stack_idx[-1] > len(stack_uppers[-1]) - 1:
                    ret = stack_level[-1]
                    stack_uppers.pop(); stack_idx.pop(); stack_level.pop()
                    if stack_uppers:
                        child_pending = ret
                    else:
                        if ret > local_best:
                            local_best = ret
                else:
                    nx = stack_uppers[-1][stack_idx[-1]]
                    stack_idx[-1] += 1
                    nr = enter(nx, stack_level[-1])
                    if nr is not None:
                        stack_uppers.append(nr[0])
                        stack_idx.append(0)
                        stack_level.append(nr[1])
            if local_best > best:
                best = local_best
        return best

    for root in nodes:
        count_dfs(root)

    elapsed = initial_load_ms + visits * MEM_LOOKUP_US / 1000
    return elapsed, sql_calls, visits


def perf_compare(name, edges, nodes):
    t_o, s_o, v_o = perf_orig_simulated(edges, nodes)
    t_p, s_p, v_p = perf_opt_simulated(edges, nodes)
    speedup = t_o / t_p if t_p > 0 else float("inf")
    print(f"[PERF] {name:30s} nodes={len(nodes):>5d} edges={sum(len(v) for v in edges.values()):>5d}  "
          f"orig={t_o:>10.1f}ms (sql={s_o:>7d})  "
          f"opt={t_p:>8.1f}ms (sql={s_p:>2d})  "
          f"speedup={speedup:>6.0f}x")


def main():
    print("=== Equivalence ===")
    cases = [
        ("chain_6_stages", *build_chain_edges()),
        ("diamond", *build_diamond()),
        ("shared_upper", *build_shared_upper()),
        ("cycle_3node", *build_with_cycle()),
        ("null_clrid", *build_null_clrid()),
    ]
    fail = 0
    for name, edges, nodes in cases:
        if not assert_equiv(name, edges, nodes):
            fail += 1

    # 大规模随机 DAG
    edges, nodes = build_random_dag(n_nodes=500, density=2.0)
    if not assert_equiv("random_dag_500", edges, nodes):
        fail += 1
    edges, nodes = build_random_dag(n_nodes=1000, density=3.0, seed=99)
    if not assert_equiv("random_dag_1000", edges, nodes):
        fail += 1

    if fail:
        print(f"\n{fail} case(s) FAILED")
        raise SystemExit(1)

    print("\nAll new-cost-calc equivalence tests passed.")

    print("\n=== Perf model: simulated SQL impl vs in-memory A3 ===")
    print("(参数: SQL RPC=1.5ms, 表扫=0.5us/行, 内存查询=0.5us/节点)")
    # 注意：原算法在中等规模上就已经指数级展开，这里跑不动大规模是预料之内
    perf_compare("dag_500_d2",  *build_random_dag(n_nodes=500,  density=2.0, seed=99))
    perf_compare("dag_1k_d2",   *build_random_dag(n_nodes=1000, density=2.0, seed=7))
    perf_compare("dag_2k_d2",   *build_random_dag(n_nodes=2000, density=2.0, seed=11))


if __name__ == "__main__":
    main()
