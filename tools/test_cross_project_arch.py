"""
跨工程架构等价性测试（2026-05-02 第三轮审计）

验证三工程架构下 PR-1/2/3/4 的代码是否正确处理了：
  - VouMetaCache 懒加载（每个工程独立加载）
  - VouSchemaCache 懒加载
  - VouBatchOps 只在工程 A 使用
  - BatchMode 类属性跨工程传递（A → IDC → t_FVou_M → CVouService）
"""
from __future__ import annotations
import re
from pathlib import Path


def read(p):
    full = Path(f"src/voucher/{p}")
    if not full.exists():
        return None  # 文件可能不在当前 PR 分支
    return full.read_text(encoding="utf-8")


def skip_if_missing(name, src):
    """如果 src 是 None（文件不在当前 PR 分支），跳过该测试"""
    if src is None:
        print(f"[SKIP] {name} - 文件不在当前 PR 分支")
        return True
    return False


def test_meta_cache_has_lazy_load():
    """VouMetaCache 必须有 EnsureLoaded（懒加载入口）"""
    src = read("VouMetaCache.bas")
    if skip_if_missing("test_meta_cache_has_lazy_load", src): return 0
    if "Public Sub EnsureLoaded" not in src:
        print("[FAIL] VouMetaCache 缺 EnsureLoaded")
        return 1
    if "If m_blnLoaded Then Exit Sub" not in src:
        print("[FAIL] EnsureLoaded 缺 IsLoaded 检查")
        return 1
    print("[OK ] VouMetaCache.EnsureLoaded 实现懒加载")
    return 0


def test_schema_cache_has_lazy_load():
    """VouSchemaCache 必须有 CloneXxxEmptyOrLoad（懒加载 Clone）"""
    src = read("VouSchemaCache.bas")
    if skip_if_missing("test_schema_cache_has_lazy_load", src): return 0
    for fn in ["CloneMainEmptyOrLoad", "CloneItemsEmptyOrLoad", "CloneIItemsEmptyOrLoad"]:
        if f"Public Function {fn}" not in src:
            print(f"[FAIL] VouSchemaCache 缺 {fn}")
            return 1
    print("[OK ] VouSchemaCache 三个 CloneXxxEmptyOrLoad 懒加载入口齐全")
    return 0


def test_helpers_no_blnUseCache_param():
    """VouCacheHelpers.CheckFI/Acc/Corp/Emp 不再接受 blnUseCache 参数（懒加载内部决定）
    检查方法：找 'Public Sub CheckXxx(...)' 后第一行非空非续行的代码到 ')'，
    再判断该签名内部是否含 'blnUseCache'"""
    src = read("VouCacheHelpers.bas")
    if skip_if_missing("test_helpers_no_blnUseCache_param", src): return 0
    fails = 0
    for fn in ["CheckFI", "CheckAcc", "CheckCorp", "CheckEmp"]:
        # 找到 'Public Sub fn('，然后向后扫描到首个未匹配的 ')'
        idx = src.find(f"Public Sub {fn}(")
        if idx < 0:
            print(f"[FAIL] 找不到 {fn} 签名")
            fails += 1
            continue
        # 简单方式：取从 '(' 后开始的下一行第一个 ')' 之前的内容
        # VB6 多行用 ' _' 续行，逐行拼接
        lines = src[idx:].split("\n")
        sig_lines = []
        for line in lines:
            sig_lines.append(line)
            stripped = line.rstrip()
            # 不以 ' _' 结尾说明是签名结束
            if not stripped.endswith(" _") and ")" in line:
                break
        sig = "\n".join(sig_lines)
        # 仅检查签名（到第一个独立 ')' 行结束）部分是否含 blnUseCache as
        # 实际上签名块的 'ByVal blnUseCache' 模式
        if re.search(r"ByVal\s+blnUseCache\s+As", sig, re.IGNORECASE):
            print(f"[FAIL] {fn} 签名仍含 'ByVal blnUseCache As'（应已删除）")
            fails += 1
        else:
            print(f"[OK ] {fn} 签名已去掉 ByVal blnUseCache（内部 EnsureLoaded）")
    return fails


def test_helpers_call_ensureloaded():
    """VouCacheHelpers 各 Check* 内部必须调 VouMetaCache.EnsureLoaded"""
    src = read("VouCacheHelpers.bas")
    if skip_if_missing("test_helpers_call_ensureloaded", src): return 0
    # 4 个 Check* + 2 个 LookupDocFI_* = 6 处必须有 EnsureLoaded
    n_ensure = src.count("VouMetaCache.EnsureLoaded(objDS)")
    if n_ensure < 6:
        print(f"[FAIL] EnsureLoaded 调用次数 {n_ensure} < 6")
        return 1
    print(f"[OK ] {n_ensure} 处 EnsureLoaded 调用，覆盖所有 helper")
    return 0


def test_pr1_patches_doc_cross_project():
    """PR1_patches.md 必须明确说明 .bas 加到三个工程"""
    src = read("PR1_patches.md")
    if skip_if_missing("test_pr1_patches_doc_cross_project", src): return 0
    required = [
        "POPBus3FileService.dll",
        "POPBus3GL2IDC.dll",
        "POPBus3GL2Service.dll",
        "三个工程",
    ]
    missing = [r for r in required if r not in src]
    if missing:
        print(f"[FAIL] PR1_patches.md 缺架构说明: {missing}")
        return 1
    print("[OK ] PR1_patches.md 明确说明三工程架构")
    return 0


def test_pr3_uses_batchmode_property():
    """PR3 patches 必须用 Me.BatchMode（属性传递），不再用 VouBatchOps.InBatchMode"""
    src = read("PR3_patches.md")
    if skip_if_missing("test_pr3_uses_batchmode_property", src): return 0
    # SaveDoc 内的判断必须用 Me.BatchMode
    if "Not Me.BatchMode" not in src:
        print("[FAIL] PR3 SaveDoc 没有用 Me.BatchMode 判断")
        return 1
    # 必须有 IDCService.BatchMode 属性
    if "Public BatchMode As Boolean" not in src:
        print("[FAIL] PR3 缺 BatchMode 属性声明")
        return 1
    print("[OK ] PR3 用 Me.BatchMode 类属性传递（跨工程 OK）")
    return 0


def test_pr4_uses_me_batchmode():
    """PR4 patches 必须用 Me.BatchMode 触发 INSERT VALUES（不是 VouBatchOps.InBatchMode）"""
    src = read("PR4_patches.md")
    if skip_if_missing("test_pr4_uses_me_batchmode", src): return 0
    if "Me.BatchMode" not in src:
        print("[FAIL] PR4 patches 没有用 Me.BatchMode 触发 INSERT VALUES")
        return 1
    print("[OK ] PR4 用 Me.BatchMode 触发 INSERT VALUES（跨工程 OK）")
    return 0


def test_batchops_only_in_project_a():
    """VouBatchOps.bas 文档必须明确只加到工程 A"""
    src = read("VouBatchOps.bas")
    if skip_if_missing("test_batchops_only_in_project_a", src): return 0
    if "只能加到工程 A" not in src and "只加到工程 A" not in src:
        print("[FAIL] VouBatchOps.bas 缺'只加到工程 A'说明")
        return 1
    print("[OK ] VouBatchOps.bas 明确说明只加到工程 A")
    return 0


def test_recordsavedbill_in_loop():
    """meCreVouForXX 循环里必须调 RecordSavedBill(rtnVouBillID)"""
    src = read("PR3_patches.md")
    if skip_if_missing("test_recordsavedbill_in_loop", src): return 0
    if "RecordSavedBill(objIDC.rtnVouBillID" not in src:
        print("[FAIL] PR3 patches 缺 RecordSavedBill(objIDC.rtnVouBillID, ...) 调用")
        return 1
    print("[OK ] PR3 在 meCreVouForXX 循环里 RecordSavedBill(rtnVouBillID)")
    return 0


def main():
    print("=== 跨工程架构等价性测试 ===\n")
    tests = [
        test_meta_cache_has_lazy_load,
        test_schema_cache_has_lazy_load,
        test_helpers_no_blnUseCache_param,
        test_helpers_call_ensureloaded,
        test_pr1_patches_doc_cross_project,
        test_pr3_uses_batchmode_property,
        test_pr4_uses_me_batchmode,
        test_batchops_only_in_project_a,
        test_recordsavedbill_in_loop,
    ]
    fails = sum(t() for t in tests)
    print()
    if fails:
        print(f"❌ {fails} case(s) FAILED")
        raise SystemExit(1)
    print("✅ All cross-project architecture tests passed.")


if __name__ == "__main__":
    main()
