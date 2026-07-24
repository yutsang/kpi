#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AST bundler：合併 5 build 模組 + make_report → 單一 build_report.py（de-qualify B./S./O./N./R.）。"""
import ast
import re
from pathlib import Path

SRC = Path(__file__).resolve().parent
OUT = SRC / "build_report.py"
INTERNAL = {"build_project_review_table", "build_summary_tables", "build_overview_tables",
            "build_narrative", "render_review_table_pptx", "biao2", "inspect_biao2"}
# 依賴序：被用者先定義
MODULES = ["render_review_table_pptx", "build_narrative", "build_project_review_table",
           "build_summary_tables", "build_overview_tables", "make_report"]

imports = []          # 收集非內部 import 原文
seen = set()          # 已定義 top-level 名（dedup）
body_parts = []
main_src = None       # make_report 嘅 main()

for mod in MODULES:
    src = (SRC / f"{mod}.py").read_text(encoding="utf-8")
    tree = ast.parse(src)
    for node in tree.body:
        seg = ast.get_source_segment(src, node)
        if seg is None:
            continue
        # import：非內部先收集
        if isinstance(node, ast.Import):
            if not any(a.name.split(".")[0] in INTERNAL for a in node.names):
                imports.append(seg)
            continue
        if isinstance(node, ast.ImportFrom):
            root = (node.module or "").split(".")[0]
            if root in INTERNAL or root == "kpi":
                continue
            imports.append(seg)
            continue
        # 模組 docstring（第一個 str Expr）
        if isinstance(node, ast.Expr) and isinstance(getattr(node, "value", None), ast.Constant) \
                and isinstance(node.value.value, str):
            continue
        # if __name__ == "__main__"
        if isinstance(node, ast.If):
            test = node.test
            if isinstance(test, ast.Compare) and isinstance(test.left, ast.Name) \
                    and test.left.id == "__name__":
                continue
        # def main：sub-module 掉；make_report 留起最後
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == "main":
            if mod == "make_report":
                main_src = seg
            continue
        # sys.path.insert(...) 呢類 top-level expr 掉
        if isinstance(node, ast.Expr) and isinstance(node.value, ast.Call):
            call = node.value
            if isinstance(call.func, ast.Attribute) and call.func.attr == "insert" \
                    and isinstance(call.func.value, ast.Attribute) and call.func.value.attr == "path":
                continue
        # 名稱 dedup（_rate 等）
        name = getattr(node, "name", None)
        if name is None and isinstance(node, ast.Assign) and node.targets \
                and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
        if name and name in seen:
            continue
        if name:
            seen.add(name)
        body_parts.append(f"# ── from {mod} ──\n{seg}" if name and name not in
                          ("try",) else seg)

# de-qualify module aliases
assembled = "\n\n\n".join(body_parts + ([main_src] if main_src else []))
assembled = re.sub(r"\b(B2|IB|B|S|O|N|R)\.", "", assembled)

# 去重 import（保序）
seen_imp, imp_lines = set(), []
for line in imports:
    if line not in seen_imp:
        seen_imp.add(line); imp_lines.append(line)

header = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_report.py — 單一自足檔：由底層數據（feed + 清單）生成表二審查報告 pptx。
唔需要任何 prerequisite / 其他模組 / 前置 command，直接：
    python build_report.py [entity]        # 預設 mgm
LLM 敘述（可選）：若同目錄有 {entity}_llm_narrative.json 就自動採用，否則用清單原文 fallback。
（此檔由 5 個 build 模組 + make_report 自動合併而成；報告只作 ref，全部由底層數據生成。）
"""
''' + "\n".join(imp_lines) + "\n\n\n"

OUT.write_text(header + assembled + '\n\n\nif __name__ == "__main__":\n    main()\n', encoding="utf-8")
print(f"✓ {OUT}  ({len((header+assembled).splitlines())} 行；dedup {len(seen)} 名)")
