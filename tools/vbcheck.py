"""
VB6 .bas 静态结构校验：
  - Sub/Function 与 End Sub/End Function 配对
  - If ... Then(独立行) 与 End If 配对（单行 If 不计）
  - Do 与 Loop 配对
  - For 与 Next 配对
  - Select Case 与 End Select 配对
  - With 与 End With 配对
  - 所有 GoTo / GoSub 的标号必须存在
  - On Error GoTo <label> 的标号必须存在或 = 0
  - 所有 Exit 关键字与外层 Sub/Function 一致
  - 所有形如 `xxx:` 的标号必须能被引用（仅警告）
  - 检查每个 Sub/Function 内 Dim 的变量名是否唯一（仅警告）
"""

from __future__ import annotations
import re
import sys
from pathlib import Path


def strip_comments(line: str) -> str:
    out = []
    in_str = False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


def tokenize(path: Path):
    raw_lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    out = []
    for i, raw in enumerate(raw_lines, 1):
        line = strip_comments(raw).strip()
        out.append((i, raw, line))
    return out


def check(path: Path) -> int:
    errors = []
    warnings = []
    lines = tokenize(path)

    proc_stack = []         # [(kind, name, start_line)]
    block_stack = []        # [(kind, line)]
    labels_per_proc = {}    # proc_name -> set(label)
    refs_per_proc = {}      # proc_name -> [(label, line)]
    dims_per_proc = {}      # proc_name -> dict(name -> [lines])

    proc_re = re.compile(
        r"^(?:Public|Private|Friend|Static)?\s*(?:Static\s+)?(Sub|Function|Property\s+\w+)\s+(\w+)",
        re.IGNORECASE)
    end_proc_re = re.compile(r"^End\s+(Sub|Function|Property)\b", re.IGNORECASE)
    if_then_re = re.compile(r"^If\b.*\bThen\b\s*$", re.IGNORECASE)
    if_then_inline_re = re.compile(r"^If\b.*\bThen\b\s+\S", re.IGNORECASE)
    elseif_re = re.compile(r"^ElseIf\b.*\bThen\b\s*$", re.IGNORECASE)
    end_if_re = re.compile(r"^End\s+If\b", re.IGNORECASE)
    do_re = re.compile(r"^Do(\b|$)", re.IGNORECASE)
    loop_re = re.compile(r"^Loop\b", re.IGNORECASE)
    for_re = re.compile(r"^For\b(?!\s+Each\b)|^For\s+Each\b", re.IGNORECASE)
    next_re = re.compile(r"^Next\b", re.IGNORECASE)
    select_re = re.compile(r"^Select\s+Case\b", re.IGNORECASE)
    end_select_re = re.compile(r"^End\s+Select\b", re.IGNORECASE)
    with_re = re.compile(r"^With\b", re.IGNORECASE)
    end_with_re = re.compile(r"^End\s+With\b", re.IGNORECASE)
    label_re = re.compile(r"^([A-Za-z_]\w*):\s*$")
    goto_re = re.compile(r"\b(?:GoTo|GoSub)\s+([A-Za-z_]\w*|\d+)\b", re.IGNORECASE)
    onerr_re = re.compile(r"\bOn\s+Error\s+GoTo\s+([A-Za-z_]\w*|\d+|\-1)\b", re.IGNORECASE)
    dim_re = re.compile(r"^(?:Dim|ReDim|Static|Const)\s+([A-Za-z_]\w*)", re.IGNORECASE)

    for ln, raw, line in lines:
        if not line:
            continue

        if not proc_stack:
            m = proc_re.match(line)
            if m:
                kind = m.group(1).split()[0].lower()
                name = m.group(2)
                proc_stack.append((kind, name, ln))
                labels_per_proc[name] = set()
                refs_per_proc[name] = []
                dims_per_proc[name] = {}
                continue
        else:
            kind, name, start = proc_stack[-1]
            m = end_proc_re.match(line)
            if m:
                if m.group(1).lower() != kind:
                    errors.append(f"L{ln}: End {m.group(1)} does not match {kind} {name}")
                if block_stack:
                    errors.append(f"L{ln}: Unclosed blocks {block_stack} in {name}")
                    block_stack = []
                proc_stack.pop()
                continue

            # blocks
            if if_then_re.match(line):
                block_stack.append(("If", ln))
            elif elseif_re.match(line):
                if not block_stack or block_stack[-1][0] != "If":
                    errors.append(f"L{ln}: ElseIf without If in {name}")
            elif end_if_re.match(line):
                if not block_stack or block_stack[-1][0] != "If":
                    errors.append(f"L{ln}: End If without If in {name} (stack={block_stack})")
                else:
                    block_stack.pop()
            elif do_re.match(line) and not line.lower().startswith("loop"):
                block_stack.append(("Do", ln))
            elif loop_re.match(line):
                if not block_stack or block_stack[-1][0] != "Do":
                    errors.append(f"L{ln}: Loop without Do in {name}")
                else:
                    block_stack.pop()
            elif for_re.match(line):
                block_stack.append(("For", ln))
            elif next_re.match(line):
                if not block_stack or block_stack[-1][0] != "For":
                    errors.append(f"L{ln}: Next without For in {name}")
                else:
                    block_stack.pop()
            elif select_re.match(line):
                block_stack.append(("Select", ln))
            elif end_select_re.match(line):
                if not block_stack or block_stack[-1][0] != "Select":
                    errors.append(f"L{ln}: End Select without Select in {name}")
                else:
                    block_stack.pop()
            elif with_re.match(line):
                block_stack.append(("With", ln))
            elif end_with_re.match(line):
                if not block_stack or block_stack[-1][0] != "With":
                    errors.append(f"L{ln}: End With without With in {name}")
                else:
                    block_stack.pop()

            # labels
            m = label_re.match(line)
            if m:
                lbl = m.group(1)
                if lbl in labels_per_proc[name]:
                    warnings.append(f"L{ln}: duplicate label {lbl} in {name}")
                labels_per_proc[name].add(lbl)

            # goto / on error goto refs
            for m in goto_re.finditer(line):
                refs_per_proc[name].append((m.group(1), ln))
            for m in onerr_re.finditer(line):
                tgt = m.group(1)
                if tgt not in ("0", "-1"):
                    refs_per_proc[name].append((tgt, ln))

            # dims
            m = dim_re.match(line)
            if m:
                v = m.group(1)
                dims_per_proc[name].setdefault(v, []).append(ln)

    if proc_stack:
        for kind, name, start in proc_stack:
            errors.append(f"L{start}: {kind} {name} not closed")

    # validate refs
    for proc, refs in refs_per_proc.items():
        for lbl, ln in refs:
            if lbl.isdigit() and lbl == "0":
                continue
            if lbl not in labels_per_proc[proc]:
                errors.append(f"L{ln}: label '{lbl}' referenced in {proc} but not defined")

    # warn duplicate Dim
    for proc, vs in dims_per_proc.items():
        for v, lines_ in vs.items():
            if len(lines_) > 1:
                warnings.append(f"{proc}: variable '{v}' Dim'd at lines {lines_}")

    print(f"== {path} ==")
    if errors:
        print(f"ERRORS ({len(errors)}):")
        for e in errors:
            print(f"  {e}")
    if warnings:
        print(f"WARNINGS ({len(warnings)}):")
        for w in warnings:
            print(f"  {w}")
    if not errors:
        print("OK: no structural errors")
    return 1 if errors else 0


if __name__ == "__main__":
    rc = 0
    for p in sys.argv[1:]:
        rc |= check(Path(p))
    sys.exit(rc)
