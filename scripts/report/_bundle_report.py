#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AST bundler：合併所有模組 → 單一自足 build_report.py（de-qualify B./S./O./N./R./B2./IB.）。
包含 build 鏈（render/narrative/review/summary/overview + make_report）+ LLM 鏈
（workbench/inspect_biao2/biao2/build_llm_narrative，供 --llm 用；heavy imports 全 lazy）。
改任何模組後重跑： python scripts/report/_bundle_report.py"""
import ast
import re
from pathlib import Path

SRC = Path(__file__).resolve().parent
WB = SRC.parents[1] / "src" / "kpi" / "lib" / "workbench.py"   # workbench 喺 src/
OUT = SRC / "build_report.py"
INTERNAL = {"build_project_review_table", "build_summary_tables", "build_overview_tables",
            "build_narrative", "render_review_table_pptx", "biao2", "inspect_biao2",
            "build_llm_narrative", "layout"}
# 依賴序：被用者先定義（layout 最先，其餘全部用佢；LLM 鏈喺 make_report 之前）
MODULES = [
    ("layout", SRC / "layout.py"),
    ("render_review_table_pptx", SRC / "render_review_table_pptx.py"),
    ("build_narrative", SRC / "build_narrative.py"),
    ("build_project_review_table", SRC / "build_project_review_table.py"),
    ("build_summary_tables", SRC / "build_summary_tables.py"),
    ("build_overview_tables", SRC / "build_overview_tables.py"),
    ("workbench", WB),
    ("inspect_biao2", SRC / "inspect_biao2.py"),
    ("biao2", SRC / "biao2.py"),
    ("build_llm_narrative", SRC / "build_llm_narrative.py"),
    ("make_report", SRC / "make_report.py"),
]

future_imports = []   # from __future__ … 必須喺最頂
imports = []          # 其餘非內部 import
seen = set()          # 已定義 top-level 名（dedup）
body_parts = []
main_src = None       # make_report 嘅 main()

for mod, path in MODULES:
    src = path.read_text(encoding="utf-8")
    tree = ast.parse(src)
    for node in tree.body:
        seg = ast.get_source_segment(src, node)
        if seg is None:
            continue
        if isinstance(node, ast.ImportFrom) and node.module == "__future__":
            if seg not in future_imports:
                future_imports.append(seg)
            continue
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
        if isinstance(node, ast.Expr) and isinstance(getattr(node, "value", None), ast.Constant) \
                and isinstance(node.value.value, str):
            continue                                        # 模組 docstring
        if isinstance(node, ast.If):
            test = node.test
            if isinstance(test, ast.Compare) and isinstance(test.left, ast.Name) \
                    and test.left.id == "__name__":
                continue                                    # if __name__ == "__main__"
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == "main":
            if mod == "make_report":
                main_src = seg
            continue                                        # sub-module main 掉、make_report 留最後
        if isinstance(node, ast.Expr) and isinstance(node.value, ast.Call):
            call = node.value
            if isinstance(call.func, ast.Attribute) and call.func.attr == "insert" \
                    and isinstance(call.func.value, ast.Attribute) and call.func.value.attr == "path":
                continue                                    # sys.path.insert(...)
        name = getattr(node, "name", None)
        if name is None and isinstance(node, ast.Assign) and node.targets \
                and isinstance(node.targets[0], ast.Name):
            name = node.targets[0].id
        if name and name in seen:
            continue                                        # dedup（_rate/_norm 等）
        if name:
            seen.add(name)
        body_parts.append(f"# ── from {mod} ──\n{seg}")

assembled = "\n\n\n".join(body_parts + ([main_src] if main_src else []))
assembled = re.sub(r"\b(B2|IB|B|S|O|N|R|L)\.", "", assembled)   # de-qualify module aliases

seen_imp, imp_lines = set(), []
for line in imports:
    if line not in seen_imp:
        seen_imp.add(line); imp_lines.append(line)

header = "#!/usr/bin/env python3\n# -*- coding: utf-8 -*-\n"
for fi in future_imports:                                   # __future__ 必須喺最頂
    header += fi + "\n"
header += '''"""
build_report.py — 單一自足檔：由底層數據（feed + 清單）生成表二審查報告 pptx。
毋須任何 prerequisite / 其他模組 / 前置 command：
    python build_report.py [entity]            # 預設 mgm；用現有 {entity}_llm_narrative.json 或清單 fallback
    python build_report.py [entity] --llm      # 即場生成 LLM 敘述（需 KPMG 網 + workbench creds）再出報告
（此檔由各 build/LLM 模組自動合併；LLM 相關 heavy import [openai/msoffcrypto] 全 lazy；報告只作 ref。）
"""
''' + "\n".join(imp_lines) + "\n\n\n"

OUT.write_text(header + assembled + '\n\n\nif __name__ == "__main__":\n    main()\n', encoding="utf-8")
print(f"✓ {OUT}  ({len((header + assembled).splitlines())} 行；dedup {len(seen)} 名；"
      f"future {len(future_imports)}；imports {len(imp_lines)}）")
