# -*- coding: utf-8 -*-
"""雲端建置入口：在 highhand repo checkout 內執行（雲端 routine 或任何環境）。
跑 tools/builder.py（自動抓「93H進度盤點雲端」Google Sheet）→
產出 93h/index.html 與 93h/report/plan.json。需要 openpyxl。"""
import os
import runpy
import shutil

TOOLS = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(TOOLS)
os.environ["PLAN_DIR"] = os.path.join(REPO, "93h", "report")
runpy.run_path(os.path.join(TOOLS, "builder.py"), run_name="__main__")
shutil.copyfile(os.path.join(TOOLS, "93H儀表板.html"), os.path.join(REPO, "93h", "index.html"))
print("cloud build done ->", os.path.join(REPO, "93h", "index.html"))
