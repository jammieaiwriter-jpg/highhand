# -*- coding: utf-8 -*-
"""雲端建置入口：在 highhand repo checkout 內執行（雲端 routine 或任何環境）。
先把會勘裡還沒追蹤的點補進缺失改善（失敗不擋儀表板），
再跑 tools/builder.py（自動抓「93H進度盤點雲端」Google Sheet）→
產出 93h/index.html 與 93h/report/plan.json。需要 openpyxl。

會勘入帳可關：環境變數 H93_SURVEY_IMPORT=off。
工地回報是另一本試算表時，設 H93_REPORT_SSID 為那本的 ID。
"""
import os
import runpy
import shutil
import subprocess
import sys

TOOLS = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(TOOLS)


def _import_surveys():
    """班次順手入帳。讀不到會勘或寫入失敗都只印訊息，儀表板照常重產。"""
    flag = os.environ.get("H93_SURVEY_IMPORT", "on").strip().lower()
    if flag in ("0", "off", "no", "false"):
        print("survey import skipped (H93_SURVEY_IMPORT off)")
        return
    try:
        r = subprocess.run(
            [sys.executable, os.path.join(TOOLS, "h93_tools.py"), "survey"],
            cwd=REPO, capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=180,
        )
    except Exception as e:  # noqa: BLE001
        print("survey import skipped:", e)
        return
    print("survey import rc", r.returncode)
    tail = ((r.stdout or "") + "\n" + (r.stderr or "")).strip()
    if tail:
        print(tail[-800:])


_import_surveys()
os.environ["PLAN_DIR"] = os.path.join(REPO, "93h", "report")
runpy.run_path(os.path.join(TOOLS, "builder.py"), run_name="__main__")
shutil.copyfile(os.path.join(TOOLS, "93H儀表板.html"), os.path.join(REPO, "93h", "index.html"))
print("cloud build done ->", os.path.join(REPO, "93h", "index.html"))
