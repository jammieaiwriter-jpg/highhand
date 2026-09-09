# -*- coding: utf-8 -*-
"""93H 進度儀表板產生器（唯一正本）
資料源＝Google Sheet「93H進度盤點雲端」（抓不到才退本機封存檔），產出「93H儀表板.html」。
本檔同時被 build_dashboard_entry.py 複製到 highhand_repo/tools/builder.py 供雲端 routine 使用，
所以桌機與雲端永遠跑同一份程式——改這裡就好，不要另外改 tools/builder.py。

2026-09-09 改版（玲嬅定）：
  1. 取消「📥 今日工地回報（即時）」區塊（工地不用回報頁，回報走 TG→雲端）。
  2. 出工預核：同廠商×同工項的多筆排程合併成一列接力鏈（9/13丈量→9/27安裝）；
     廠商欄一律顯示「廠商(工項)」。
  3. 未完成／改期：同一工班只留一列（列出事項×次數）；原定日之後該工班已出工→自動消失。
  5. 行事曆與交辦：只列今天以後的提醒，過期即拿掉；改期的另列在改期區。
  6. 廠商協調：已定案／結案／已到貨的不顯示。
  7. 待長官裁決：加「待長官現場確認／拍板」自各分頁自動抓取。
  8. 缺失改善：已完成（含「已完成(剩保護工程)」）不顯示。
  9. 粗工打石：新增與會議紀錄「沛誼工作報告」週統計對帳（雲端分頁「粗工對帳」）。
"""
import glob
import html
import json
import os
import re
import urllib.request
from datetime import date, timedelta

from openpyxl import load_workbook

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_HTML = os.path.join(BASE, "93H儀表板.html")
TODAY = date.today()

# 儀表板直接改狀態（2026-09-09）：每列 ✏️ → 瀏覽器 POST update_cell 到回報 Apps Script → 寫回雲端 Sheet 單格
# ＋自動記異動紀錄；頁面下次重產（整點桌機排程／雲端班次）才反映，改完的列先標灰「已送出」。
REPORT_API = "https://script.google.com/macros/s/AKfycbxxBXks2VtVjBJZhkzkzTVexEpUmcyCKMKjE18-dpiEh2SnfHcwP2VG6SMPV1NfAblJjw/exec"
SHEET_ID = "1JBdrqfzQ_AF0fG8x_RT7wjZSTCjCfgdgnX9FRkQJSd0"
# 可在儀表板上改的欄位：分頁 → [(欄號1-based, 欄名, 常用值)]
EDIT_FIELDS = {
    "出工預核": [(5, "實際出工", ["有,人", "未出工()", "⚠️有到人,但做別的", "已完成", "已重排"]), (1, "日期", []), (6, "備註", [])],
    "行政時程": [(5, "完成", ["已完成", "改期→", "取消"]), (1, "日期", []), (6, "備註", [])],
    "廠商協調": [(5, "狀態", ["已定案", "結案✓", "待約", "確認中", "待報價"]), (6, "備註", [])],
    "進料追蹤": [(5, "實際到料", ["已到料", "部分到料", "未到"]), (1, "日期", []), (6, "備註", [])],
    "待裁決": [(4, "長官裁決", []), (5, "裁決日", []), (6, "備註", [])],
    "缺失改善": [(7, "狀態", ["已完成", "進行中", "已排程", "待處理", "觀察中", "結案"]), (6, "預計完成", []), (8, "備註", [])],
    "發包與工作事項": [(3, "現況", ["完成", "結案"]), (4, "下一步", [])],
    "燈具進度": [(5, "日期狀態", ["已完成", "已進場", "施作中"]), (6, "備註", [])],
    "石材追蹤": [(3, "狀態", ["已完成", "已進場", "施作中"]), (4, "備註", []), (6, "進度", [])],
    "磁磚未進場": [(3, "狀態", ["已進場", "已完成"])],
    "新美鐵件": [(6, "安裝", ["已完成", "安裝中"]), (5, "丈量", ["已完成"]), (7, "備註", []), (9, "進度", [])],
    "門禁追蹤": [(3, "日期/狀態", ["已完成", "已進場"]), (4, "備註", [])],
    "使照檢附": [(4, "目前狀態", ["已取得✓", "已交", "已結案✓"]), (5, "備註", [])],
    "送審變更": [(5, "實際核准", ["已核准"]), (3, "送件日", []), (7, "備註", [])],
}
EDIT_KEYCOL = {"燈具進度": 3, "磁磚未進場": 1, "新美鐵件": 4, "使照檢附": 1, "送審變更": 1}


class Row(list):
    """rows_of 的列：帶原始 Sheet 列號 rn（update_cell 用）"""
    rn = 0


def ebtn(sheet, rn, key, label, vals):
    """產生 ✏️ 按鈕（data-e 帶分頁/列號/目前值）；rn 為 0 或分頁不可編輯時不產生。"""
    if not rn or sheet not in EDIT_FIELDS:
        return ""
    v = {str(c): (str(vals[c - 1]) if c - 1 < len(vals) and vals[c - 1] is not None else "") for c, _, _ in EDIT_FIELDS[sheet]}
    d = json.dumps({"s": sheet, "r": rn, "k": str(key)[:6], "kc": EDIT_KEYCOL.get(sheet, 2),
                    "l": str(label)[:40], "v": v}, ensure_ascii=False)
    return f' <button class="eb" data-e="{html.escape(d, quote=True)}" title="改狀態" type="button">✏️</button>'

# ---------------- 里程碑 ----------------
# ★ 治理規則：里程碑為長官核定版，只有玲嬅同意才能改（柏銘表動到里程碑時
#   只能「標示警告」，不得逕改此處）。詳見 工地/AGENTS.md。
MILESTONES = [
    ("1F門禁完成(玻璃/新美門/鐵捲門/防火門)", date(2026, 9, 30), ""),
    ("拆圍籬、景觀開挖", date(2026, 10, 1), "借地10/1~116/2/28已定"),
    ("2F~1F外牆拆架完成", date(2026, 11, 15), "9/2會議三度延(11/13~11/15拆);目標提前待2F~7F清潔後調整"),
    ("消防檢查開始", date(2026, 11, 15), "9/2會議延11/15~116/1/15⚠️尾端壓使照;前置:消防變更9/15~10/15"),
    ("無障礙掛件", date(2026, 11, 15), ""),
    ("無障礙、使管處看現場", date(2026, 12, 5), "使照取得前送資料即可"),
    ("取得使用執照", date(2027, 1, 15), ""),
]

# ---------------- 工具 ----------------
DONE_PAT = re.compile(r"已完成|已進場|完成$")
READY_PAT = re.compile(r"已可安裝|可安排")


def minguo(d: date) -> str:
    return f"{d.year - 1911}/{d.month}/{d.day}"


def parse_date(text: str):
    """從 115/09/01、116.03.30、115/8/21前完工、115/9月底 之類字串抓出西元日期。"""
    if not text:
        return None
    m = re.search(r"(11[4-9])[./](\d{1,2})[./](\d{1,2})", text)
    if m:
        y, mo, dd = int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, dd)
        except ValueError:
            return None
    m = re.search(r"(11[4-9])[./](\d{1,2})\s*月(底|初|中)?", text)
    if m:
        y, mo = int(m.group(1)) + 1911, int(m.group(2))
        part = m.group(3) or "底"
        dd = {"初": 5, "中": 15}.get(part, 28)
        try:
            return date(y, mo, dd)
        except ValueError:
            return None
    return None


def is_yellow(cell) -> bool:
    f = cell.fill
    if f is None or f.patternType != "solid":
        return False
    rgb = getattr(f.fgColor, "rgb", None) or ""
    return str(rgb).upper().endswith(("FFFF00", "FFE699", "FFF2CC", "FFEB9C"))


def classify(status_text: str, note: str):
    """回傳 (bucket, 到期日)。bucket: done/overdue/soon/later/ready/waiting"""
    s = (status_text or "").strip()
    n = (note or "").strip()
    if DONE_PAT.search(s):
        return "done", None
    if "施作中" in s or "安裝中" in s:
        d = parse_date(n)
        if d:
            if d < TODAY:
                return "overdue", d
            if d <= TODAY + timedelta(days=14):
                return "soon", d
            return "later", d
        return "doing", None
    d = parse_date(s)
    if d:
        if d < TODAY:
            return "overdue", d
        if d <= TODAY + timedelta(days=14):
            return "soon", d
        return "later", d
    if READY_PAT.search(s) or READY_PAT.search(n):
        return "ready", None
    return "waiting", None


# ---- 廠商／工項正規化（出工預核顯示用）----
STEP_WORDS = re.compile(r"完成期限|進場安裝|進場施作|丈量|安裝|進場|開始|期限|完成|施作|復工|退場|放樣|吊料")


def vbase(v) -> str:
    """廠商基底名：去括號、去問號。『鈦翔(林文信)』→鈦翔、『清潔(永盛?)』→清潔"""
    v = str(v or "").strip()
    v = re.sub(r"[（(].*?[)）]", "", v)
    v = re.sub(r"[?？]", "", v).strip()
    return v


def work_core(w) -> str:
    """工項核心詞（≤8字）：去括號/日期/人數/工序動詞，用來判斷同一工項。"""
    w = str(w or "")
    w = re.sub(r"[（(][^()（）]*[)）]", "", w)
    w = re.sub(r"^\d+人\s*[:：]?", "", w)
    w = re.sub(r"\d{1,3}/\d{1,2}(\s*[~～\-]\s*\d{1,3}/\d{1,2})?", "", w)
    w = re.split(r"[+＋→；;]", w)[0]
    w = STEP_WORDS.sub("", w)
    w = re.sub(r"[\s,，:：、⭐⚠️]+", "", w)
    return w[:8]


def vendor_label(r) -> str:
    """廠商欄統一格式：廠商(工項)。"""
    b = vbase(r[1]) or str(r[1] or "").strip()
    c = work_core(r[2])
    if not c or c == b or b.startswith(c):
        return b or c
    return f"{b}({c})"


def attended(v) -> bool:
    """實際出工欄是否代表『人有到』（有,N人／已完成／已丈量／⚠️有到但做別的）。"""
    return bool(re.match(r"^[\s⚠️✅]*(有|已)", str(v or "")))


# ---------------- 讀資料源 ----------------
# 正本＝Google Sheet「93H進度盤點雲端」（2026-09-03 雲端化）。每次執行先抓雲端匯出；
# 抓不到（斷網）才退回本機最新 xlsx。本機 93H進度盤點_1150826.xlsx 已封存不再維護。
CLOUD_XLSX = ("https://docs.google.com/spreadsheets/d/"
              "1JBdrqfzQ_AF0fG8x_RT7wjZSTCjCfgdgnX9FRkQJSd0/export?format=xlsx")
CLOUD_CACHE = os.path.join(BASE, "_雲端盤點快取.xlsx")
SRC = None
try:
    req = urllib.request.Request(CLOUD_XLSX, headers={"User-Agent": "93H-dashboard"})
    data = urllib.request.urlopen(req, timeout=30).read()
    if len(data) > 10000:
        with open(CLOUD_CACHE, "wb") as _f:
            _f.write(data)
        SRC = CLOUD_CACHE
        print("資料源：雲端 93H進度盤點雲端（Google Sheets）")
except Exception as _e:
    print(f"雲端抓取失敗（{_e}），改用本機檔")
if SRC is None:
    files = sorted(glob.glob(os.path.join(BASE, "93H進度盤點_*.xlsx")))
    if os.path.exists(CLOUD_CACHE):
        files.append(CLOUD_CACHE)
    if not files:
        raise SystemExit("找不到資料源（雲端失敗且無本機 93H進度盤點_*.xlsx）")
    SRC = files[-1]
SRC_LABEL = "93H進度盤點雲端(Google Sheet)" if SRC == CLOUD_CACHE else os.path.basename(SRC)
wb = load_workbook(SRC, data_only=True)

items = []  # dict: 類別, 位置, 項目, 前段(送樣/丈量), 日期狀態, 備註, 使照前, bucket, due


def add(cat, loc, name, pre, status, note, permit, prog="", ms="", sheet="", rn=0, raw=None):
    b, due = classify(status, note)
    items.append(dict(cat=cat, loc=loc, name=name, pre=pre, status=status,
                      note=note, permit=permit, bucket=b, due=due, prog=prog,
                      ms=ms or "報完工", sheet=sheet, rn=rn, raw=raw or []))


def txt(v):
    return "" if v is None else str(v).strip()


ws = wb["燈具進度"]
for row in ws.iter_rows(min_row=2):
    vals = [txt(c.value) for c in row[:7]]
    loc, _, name, sample, plan, note = vals[:6]
    if not name:
        continue
    permit = is_yellow(row[2])
    add("燈具", loc, name, sample, plan, note, permit, ms=vals[6] if len(vals) > 6 else "",
        sheet="燈具進度", rn=row[0].row, raw=vals)

ws = wb["石材追蹤"]
for row in ws.iter_rows(min_row=2):
    vals = [txt(c.value) for c in row[:7]]
    grp, name, status, note = vals[0], vals[1], vals[2], vals[3]
    permit = (len(vals) > 4 and vals[4].upper() in ("V", "★")) or is_yellow(row[1])
    if not name:
        continue
    add("石材", grp, name, "", status, note, permit,
        prog=vals[5] if len(vals) > 5 else "", ms=vals[6] if len(vals) > 6 else "",
        sheet="石材追蹤", rn=row[0].row, raw=vals)

ws = wb["磁磚未進場"]
for row in ws.iter_rows(min_row=2):
    vals = [txt(c.value) for c in row[:4]]
    name, size, status = vals[:3]
    if not name:
        continue
    add("磁磚", size, name, "", status, "", is_yellow(row[0]), ms=vals[3] if len(vals) > 3 else "",
        sheet="磁磚未進場", rn=row[0].row, raw=vals)

ws = wb["新美鐵件"]
for row in ws.iter_rows(min_row=2):
    vals = [txt(c.value) for c in row[:10]]
    grp, loc, _, name, measure, install, note = vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], vals[6]
    permit = (len(vals) > 7 and vals[7].upper() in ("V", "★")) or is_yellow(row[3])
    if not name:
        continue
    status = install if (install and not DONE_PAT.search(measure)) or DONE_PAT.search(install) else (install or measure)
    add("鐵件", f"{grp}/{loc}", name, measure, status or install, note, permit,
        prog=vals[8] if len(vals) > 8 else "", ms=vals[9] if len(vals) > 9 else "",
        sheet="新美鐵件", rn=row[0].row, raw=vals)


# ---- 出工預核 / 缺失改善 / 粗工打石 ----
def rows_of(name, ncol):
    if name not in wb.sheetnames:
        return []
    out = []
    for i, row in enumerate(wb[name].iter_rows(min_row=2, values_only=True), 2):
        vals = [txt(v) for v in (row[:ncol] if row else [])]
        if any(vals):
            r = Row(vals + [""] * (ncol - len(vals)))
            r.rn = i
            out.append(r)
    return out


dispatch_all = rows_of("出工預核", 6)


def d_key(r):
    d = parse_date(r[0])
    return d or date.max


# 未來排程（明日以後、實際出工空白）＋今日列（全部）
dispatch_upcoming = sorted(
    [r for r in dispatch_all
     if parse_date(r[0]) == TODAY or ((parse_date(r[0]) or date.max) > TODAY and not r[4])],
    key=lambda r: (d_key(r), r[1]))
dispatch_today = [r for r in dispatch_all if parse_date(r[0]) == TODAY]
# 鐵則（玲嬅 2026-09-07 定）：應出未回報＝預定日已到(含今天)且「實際出工」欄空白 → 一律紅底顯示
# 補充（9/9）：同廠商當天另一列已填實際出工 → 該廠商當日已有回報，不重複列
_reported_days = {(vbase(r[1]), parse_date(r[0])) for r in dispatch_all if str(r[4]).strip()}
dispatch_unverified = sorted(
    [r for r in dispatch_all
     if (parse_date(r[0]) or date.max) <= TODAY and not str(r[4]).strip()
     and (vbase(r[1]), parse_date(r[0])) not in _reported_days],
    key=d_key)
# 同工班最近一次「人有到」的日期（用來自動結案改期列）
last_attend = {}
for _r in dispatch_all:
    _d = parse_date(_r[0])
    if _d and _d <= TODAY and attended(_r[4]):
        _k = vbase(_r[1])
        if _d > last_attend.get(_k, date.min):
            last_attend[_k] = _d
FAIL_PAT = re.compile(r"未出|未完成|未進|未派|改期|未施作")
dispatch_failed = sorted(
    [r for r in dispatch_all
     if (parse_date(r[0]) or date.max) <= TODAY and FAIL_PAT.search(str(r[4]))
     and "已重排" not in str(r[4]) and "已完成" not in str(r[4])
     # 規則3（玲嬅 9/9）：原定日之後同一工班已出工 → 改期列自動取消
     and not (last_attend.get(vbase(r[1]), date.min) > d_key(r))],
    key=d_key)

mat_all = rows_of("進料追蹤", 6)
mat_upcoming = sorted(
    [r for r in mat_all
     if parse_date(r[0]) == TODAY or ((parse_date(r[0]) or date.max) > TODAY and not r[4]) or not r[0]],
    key=d_key)
mat_unverified = [r for r in mat_all if r[0] and (parse_date(r[0]) or date.max) < TODAY and not r[4]]

admin_all = rows_of("行政時程", 6)
# 規則5（玲嬅 9/9）：行事曆只提醒今天以後的事，過期即拿掉（不再列「未記結果」）
admin_upcoming = sorted(
    [r for r in admin_all
     if parse_date(r[0]) == TODAY or ((parse_date(r[0]) or date.max) > TODAY and not r[4])],
    key=d_key)
_admin_last_done = {}
for _r in admin_all:
    _d = parse_date(_r[0])
    if _d and _d <= TODAY and _r[4] and not re.search(r"改期|未", str(_r[4])):
        _k = vbase(_r[1])
        if _d > _admin_last_done.get(_k, date.min):
            _admin_last_done[_k] = _d
admin_failed = [r for r in admin_all
                if r[0] and (parse_date(r[0]) or date.max) <= TODAY
                and re.search(r"改期|未", str(r[4])) and "已重排" not in str(r[4])
                and not (_admin_last_done.get(vbase(r[1]), date.min) > d_key(r))]

decisions = [r for r in rows_of("待裁決", 6) if not r[3]]

# 規則8（玲嬅 9/9）：已完成（含「已完成(剩保護工程)」）／結案 一律不顯示
defects = [r for r in rows_of("缺失改善", 8)
           if not re.match(r"^\s*(已完成|完成|結案|已結案)", r[6]) and "結案" not in r[6]]

# 規則6（玲嬅 9/9）：廠商協調只列還沒定案的
coord_all = rows_of("廠商協調", 6)
COORD_DONE = re.compile(r"已定案|結案|✓|已到貨|已完成|^完成|定案$")
coord_open = [r for r in coord_all if not COORD_DONE.search(r[4])]

labor = rows_of("粗工打石", 8)


def month_key(r):
    d = parse_date(r[0])
    return (d.year, d.month) if d else None


def labor_n(r):
    try:
        return float(re.sub(r"[^\d.]", "", str(r[3])) or 0)
    except ValueError:
        return 0.0


this_month = (TODAY.year, TODAY.month)
labor_stats = {}
for r in labor:
    who = r[5] or "未填"
    n = labor_n(r)
    amt = 0
    try:
        amt = float(re.sub(r"[^\d.]", "", str(r[6])) or 0)
    except ValueError:
        pass
    s = labor_stats.setdefault(who, {"m_n": 0, "m_amt": 0, "t_n": 0, "t_amt": 0})
    s["t_n"] += n; s["t_amt"] += amt
    if month_key(r) == this_month:
        s["m_n"] += n; s["m_amt"] += amt

# 規則9（玲嬅 9/9）：與會議紀錄「沛誼工作報告」週統計對帳
# 雲端分頁「粗工對帳」：週期間|起日|迄日|沛誼總工|沛誼公司|沛誼廠商|超工累計|電鑽工|本週廢棄物|累計台數|來源|備註
recon_rows = rows_of("粗工對帳", 12)
LABOR_LEDGER_START = date(2026, 8, 26)  # 粗工打石自此日起記，既往不究


def labor_week(d0, d1):
    tot = comp = vend = 0.0
    for r in labor:
        d = parse_date(r[0])
        if d and d0 <= d <= d1:
            n = labor_n(r)
            tot += n
            if str(r[5]).startswith("公司"):
                comp += n
            else:
                vend += n
    return tot, comp, vend


recon_disp = []
_last_end = None
for r in recon_rows:
    d0, d1 = parse_date(r[1]), parse_date(r[2])
    if not (d0 and d1):
        continue
    tot, comp, vend = labor_week(d0, d1)
    try:
        ptot = float(re.sub(r"[^\d.]", "", r[3]) or 0)
        pcomp = float(re.sub(r"[^\d.]", "", r[4]) or 0)
        pvend = float(re.sub(r"[^\d.]", "", r[5]) or 0)
    except ValueError:
        ptot = pcomp = pvend = 0
    ok = abs(tot - ptot) < 0.01 and abs(comp - pcomp) < 0.01 and abs(vend - pvend) < 0.01
    if d0 < LABOR_LEDGER_START and not ok:
        ok = None  # 盤點表自 8/26 起記，更早的週只能部分核對，不算不符
    def _num(v):
        v = str(v or "")
        return v[:-2] if re.fullmatch(r"\d+\.0", v) else v
    recon_disp.append(dict(period=r[0], ptot=ptot, pcomp=pcomp, pvend=pvend, tot=tot, comp=comp, vend=vend,
                           over=_num(r[6]), drill=_num(r[7]), waste=_num(r[8]), trucks=_num(r[9]), src=r[10], note=r[11], ok=ok))
    if _last_end is None or d1 > _last_end:
        _last_end = d1
recon_pending = None  # 沛誼尚未報告的期間（自上週迄日+1 到今天）
if _last_end and _last_end < TODAY:
    _p0 = _last_end + timedelta(days=1)
    recon_pending = (_p0, TODAY) + labor_week(_p0, TODAY)

if "門禁追蹤" in wb.sheetnames:
    ws = wb["門禁追蹤"]
    for row in ws.iter_rows(min_row=2):
        vals = [txt(c.value) for c in row[:7]]
        vendor, name, status, note = vals[0], vals[1], vals[2], vals[3]
        if not name:
            continue
        permit = (len(vals) > 4 and vals[4].upper() in ("V", "★"))
        add("門禁" if (vals[6] if len(vals) > 6 else "") == "1F門禁" else "工項", vendor, name, "", status, note, permit,
            prog=vals[5] if len(vals) > 5 else "", ms=vals[6] if len(vals) > 6 else "1F門禁",
            sheet="門禁追蹤", rn=row[0].row, raw=vals)

submissions = rows_of("送審變更", 7)
for s in submissions:
    if s[5] in ("1F門禁", "外牆拆架", "消防檢查", "使照檢查") and s[0] != "消防檢查(消檢)":
        add("送審", s[1], s[0], "", "已完成" if s[4] else (s[3] or s[2]), s[6], True, ms=s[5],
            sheet="送審變更", rn=s.rn, raw=s)

work_rows = []
ws = wb["發包與工作事項"]
for _ri, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
    vals = [txt(v) for v in (row[:4] if row else [])]
    if len(vals) < 2 or not vals[1]:
        continue
    if "完成" == vals[2].strip() or (vals[2].strip().startswith("完成") and len(vals[2].strip()) <= 4):
        continue
    if re.match(r"^\s*(結案|已結案)", vals[2]) or re.match(r"^\s*(結案|已結案)", vals[3]):
        continue
    _wr = Row(vals + [""] * (4 - len(vals)))
    _wr.rn = _ri
    work_rows.append(_wr)

# ---------------- 規則7：待長官現場確認／拍板（自各分頁自動抓取） ----------------
PEND_PAT = re.compile(
    r"(待|需|請|等|盼|由)[^,;，；。]{0,8}?(長官|總經理|公司|董事長)[^,;，；。]{0,8}?"
    r"(確認|裁示|裁決|決定|拍板|核可|定案|審閱|裁|同意|核定|決議|洽談)"
    r"|(長官|總經理|董事長)(現場|到場|來)(確認|看|檢討|裁|再確認)"
    r"|待定案|待裁決|待裁示|提案待|待公司決")
PEND_DONE = re.compile(r"已確認|已定案|已裁|已核|已決|已回覆|結案|已同意|已現場確認|已定|OK")


def pending_hits(text):
    """回傳含『待長官／總經理確認』語意的子句（排除已完成語意）。"""
    hits = []
    for clause in re.split(r"[,;，；。\n]", str(text or "")):
        if PEND_PAT.search(clause) and not PEND_DONE.search(clause):
            hits.append(clause.strip())
    return hits


site_confirm = []  # (分頁, 事項, 子句, 對象/廠商)


def scan_pending(sheet, rows, name_i, text_is, who_i=None):
    for r in rows:
        text = "｜".join(str(r[i]) for i in text_is if i < len(r))
        hits = pending_hits(text)
        if hits:
            site_confirm.append((sheet, r[name_i], "；".join(hits[:2]), r[who_i] if who_i is not None else ""))


scan_pending("發包", work_rows, 1, [2, 3], 0)
scan_pending("缺失", defects, 1, [4, 6, 7], 3)
scan_pending("協調", coord_open, 3, [4, 5], 1)
scan_pending("行事曆", admin_upcoming, 2, [2, 5], 1)
scan_pending("出工", [r for r in dispatch_upcoming if parse_date(r[0]) != TODAY], 2, [2, 5], 1)
scan_pending("送審", submissions, 0, [6], 1)
scan_pending("門禁", rows_of("門禁追蹤", 7), 1, [2, 3], 0)
scan_pending("進料", mat_upcoming, 2, [4, 5], 1)
for i in items:
    if i["bucket"] != "done":
        h = pending_hits(i["status"] + "｜" + i["note"])
        if h:
            site_confirm.append((i["cat"], i["name"], "；".join(h[:2]), i["loc"]))

# ---------------- 統計 ----------------
open_items = [i for i in items if i["bucket"] != "done"]
overdue = sorted([i for i in open_items if i["bucket"] == "overdue"], key=lambda x: x["due"])
soon = sorted([i for i in open_items if i["bucket"] == "soon"], key=lambda x: x["due"])
ready = [i for i in open_items if i["bucket"] == "ready"]
doing = [i for i in open_items if i["bucket"] == "doing"]
later = sorted([i for i in open_items if i["bucket"] == "later"], key=lambda x: x["due"])
waiting = [i for i in open_items if i["bucket"] == "waiting"]
permit_open = [i for i in open_items if i["permit"]]

BUCKET_LABEL = {
    "overdue": "已過預定日", "soon": "兩週內到期", "ready": "已可施作待排",
    "later": "已排程", "waiting": "等前置條件", "done": "已完成", "doing": "施作中",
}
BUCKET_CLASS = {
    "overdue": "b-red", "soon": "b-orange", "ready": "b-blue",
    "later": "b-gray", "waiting": "b-gray", "done": "b-green", "doing": "b-green",
}


def esc(s):
    return html.escape(str(s)).replace("\n", "<br>")


def badge(i):
    lbl = BUCKET_LABEL[i["bucket"]]
    if i["due"]:
        delta = (i["due"] - TODAY).days
        lbl += f"｜{minguo(i['due'])}" + (f"（逾{-delta}天）" if delta < 0 else f"（{delta}天後）")
    return f'<span class="badge {BUCKET_CLASS[i["bucket"]]}">{esc(lbl)}</span>'


def item_rows(lst, show_cat=True):
    out = []
    for i in lst:
        star = "⭐" if i["permit"] else ""
        cat = f'<td class="cat">{esc(i["cat"])}</td>' if show_cat else ""
        note = i["note"] or i["pre"]
        if i.get("prog"):
            note = f"<b>{esc(i['prog'])}</b>" + ("｜" + esc(note) if note else "")
        else:
            note = esc(note)
        eb = ebtn(i.get("sheet"), i.get("rn"), i["name"], f"{i['cat']} {i['name']}", i.get("raw") or [])
        out.append(
            f"<tr>{cat}<td>{esc(i['loc'])}</td>"
            f"<td>{star}{esc(i['name'])}{eb}</td><td>{badge(i)}</td>"
            f"<td class='note'>{note}</td></tr>")
    return "\n".join(out)


def section_table(title, lst, tid, note=""):
    if not lst:
        return ""
    head = "<tr><th>類別</th><th>位置</th><th>項目</th><th>狀態</th><th>備註/前置</th></tr>"
    return (f'<details class="sec" id="{tid}"><summary>{title}'
            f'<span class="cnt">{len(lst)}</span>{note}</summary>'
            f'<table>{head}{item_rows(lst)}</table></details>')


# ---------------- 里程碑反推樹 ----------------
# (名稱, 顯示日, 說明, 逾期判定日, 預設展開, 顯示風險警告)
MS_ORDER = [
    ("1F門禁", date(2026, 9, 30), "玻璃/新美門/鐵捲門(青岱)/防火門(佳昇)——拆圍籬(10月初)的前置", date(2026, 9, 30), True, True),
    ("拆架2-7F", date(2026, 10, 7), "9/1會議:因雨延兩週,7F~2F拆架約10/7~10/14；標準層外觀與清潔須趕在該區拆架前完成", date(2026, 10, 14), True, True),
    ("拆架1-2F", date(2026, 11, 13), "2F~1F拆架11/13~11/15(9/2會議)：車道格柵/車道天花板油漆/2F~1F清潔(11/5~11/12)須先完成", date(2026, 11, 15), True, True),
    ("消防檢查", date(2026, 11, 15), "消檢11/15~116/1/15(9/2會議延,尾端壓使照⚠️)；前置鏈：電力變更→消防變更→送電/設備安裝", date(2026, 11, 15), True, True),
    ("使照檢查", date(2026, 12, 5), "12月初無障礙+使管處看現場查核的必要工項；含景觀(園藝核對)；無障礙掛件11/15先行", date(2026, 12, 5), True, True),
    ("景觀工程", date(2026, 10, 1), "借地期間10/1~116/2/28施作(暫置土方/回填)；9/7重排景觀進度；園藝核對項在使照檢查", date(2027, 2, 28), False, False),
    ("報完工", date(2026, 12, 30), "總經理指示：使照前完成，不卡使照", date(2026, 12, 30), False, False),
    ("使照後二工", date(2027, 4, 30), "使照後施作：二工件＋公設裝潢＋樣品屋3A1/13A3＋接待中心S1(先建後售,代銷進駐)", date(2027, 4, 30), False, False),
]
GATING_MS = {"1F門禁", "拆架2-7F", "拆架1-2F", "消防檢查", "使照檢查"}
MS_ICON = {"報完工": "📦", "使照後二工": "🔧", "景觀工程": "🌿"}
MS_LABEL = {"報完工": "使照前完整度", "1F門禁": "門禁(1F/MF/2F)"}
SUB_DEF = [
    ("overdue", "🔴 已過期", True),
    ("soon", "🟠 兩週內到期", True),
    ("later", "📅 已安排", True),
    ("ready", "🔵 可施作但未排日期", True),
    ("doing", "🟢 施作中(缺完成日)", True),
    ("waiting", "⏳ 卡住－等前置", True),
]


def tree_html():
    parts = []
    for ms_name, ms_date, ms_note, ms_deadline, ms_open, ms_risk in MS_ORDER:
        group = [i for i in items if i["ms"] == ms_name]
        if not group:
            continue
        done = [i for i in group if i["bucket"] == "done"]
        open_i = [i for i in group if i["bucket"] != "done"]
        pct = int(len(done) / len(group) * 100) if group else 0
        delta = (ms_date - TODAY).days
        risky = [i for i in open_i if i["due"] and i["due"] > ms_deadline] if ms_risk else []
        head = (f'<div class="ms-head">'
                f'<span class="ms-meta">完成 {len(done)}/{len(group)}（{pct}%）</span>'
                f'<div class="bar"><div class="bar-in" style="width:{pct}%"></div></div>'
                f'<div class="ms-sub">{esc(ms_note)}</div>'
                + (f'<div class="ms-risk">⚠️ {len(risky)} 項排程日晚於里程碑：'
                   + "、".join(esc(i["name"]) for i in risky[:6]) + "</div>" if risky else "")
                + "</div>")
        subs = []
        for key, label, _ in SUB_DEF:
            lst = [i for i in open_i if i["bucket"] == key]
            if key == "waiting":
                lst = sorted(lst, key=lambda x: x["note"] or x["pre"])
            elif key in ("overdue", "soon", "later"):
                lst = sorted(lst, key=lambda x: x["due"] or date.max)
            if not lst:
                continue
            subs.append(f'<div class="subh">{label}（{len(lst)}）</div>'
                        f'<table>{item_rows(lst)}</table>')
        cd = f"剩{delta}天" if delta >= 0 else f"逾{-delta}天"
        cd_cls = ' style="color:#c0392b"' if (delta < 21 and ms_risk) else ''
        icon = MS_ICON.get(ms_name, "🎯")
        label = MS_LABEL.get(ms_name, ms_name)
        parts.append(f'<details class="sec ms-tree"><summary>{icon} {esc(label)}'
                     f'<span class="ms-date-tag">{minguo(ms_date)}</span>'
                     f'<span class="ms-cd-tag"{cd_cls}>{cd}</span>'
                     f'<span class="cnt">{len(open_i)}項未完</span>'
                     f'<span class="pct">{pct}%</span></summary>{head}{"".join(subs)}</details>')
    return "\n".join(parts)


# ---------------- 里程碑 HTML ----------------
ms_html = []
for name, d, note in MILESTONES:
    delta = (d - TODAY).days
    cls = "ms-past" if delta < 0 else ("ms-hot" if delta <= 30 else "ms-ok")
    cd = f"逾{-delta}天" if delta < 0 else f"剩{delta}天"
    ms_html.append(
        f'<div class="ms {cls}"><div class="ms-date">{minguo(d)}</div>'
        f'<div class="ms-name">{esc(name)}</div>'
        f'<div class="ms-cd">{cd}</div>'
        + (f'<div class="ms-note">{esc(note)}</div>' if note else "") + "</div>")

license_day = (MILESTONES[-1][1] - TODAY).days

work_93h = [r for r in work_rows if r[0] in ("海興段", "其他", "")]
work_other = [r for r in work_rows if r[0] not in ("海興段", "其他", "")]


def work_table(rows):
    return "".join(
        f"<tr><td class='cat'>{esc(r[0])}</td><td>{esc(r[1])}{ebtn('發包與工作事項', getattr(r, 'rn', 0), r[1], r[1], r)}</td>"
        f"<td class='note'>{esc(r[2])}</td><td class='note'>{esc(r[3])}</td></tr>"
        for r in rows)


work_html = work_table(work_93h)
work_other_html = work_table(work_other)

# ---------------- 一週天氣與出工部署（2026-09-03 加入） ----------------
# 工地：台南市安平區海興段。Open-Meteo 免金鑰，抓不到用快取，再不行整段隱藏。
WX_LAT, WX_LON = 22.995, 120.163
WX_CACHE = os.path.join(BASE, "_天氣快取.json")
WX_URL = ("https://api.open-meteo.com/v1/forecast"
          f"?latitude={WX_LAT}&longitude={WX_LON}"
          "&daily=weather_code,temperature_2m_max,temperature_2m_min,"
          "precipitation_sum,precipitation_probability_max,wind_gusts_10m_max"
          "&timezone=Asia%2FTaipei&forecast_days=7")
WX_ICON = [(0, "☀️晴"), (2, "🌤多雲時晴"), (3, "☁️陰"), (48, "🌫霧"), (57, "🌦毛毛雨"),
           (67, "🌧雨"), (77, "🌧雨"), (82, "🌧陣雨"), (86, "🌧陣雨"), (99, "⛈雷雨")]
OUTDOOR_PAT = re.compile(
    r"拆架|鷹架|圍籬|圍牆|開挖|景觀|外牆|外壁|吊(?!頂)|防水|試水|露臺|露台|陽台|屋頂|帆布|"
    r"廢棄物|清運|植栽|放樣|排水溝|鋪面|洗窗|鐵捲門|台電外線|戶外|外部")
INDOOR_PAT = re.compile(r"室內|浴室|梯廳|天花|地下|B\d|店舖地板|各樓層")
WEEKDAY_ZH = "一二三四五六日"


def wx_icon(code):
    for c, s in WX_ICON:
        if code <= c:
            return s
    return "🌧"


def wx_level(rain, prob, gust, code):
    """紅=不宜戶外 黃=有風險 綠=可出工"""
    if rain >= 10 or code >= 95 or gust >= 60 or (prob >= 80 and rain >= 3):
        return "red"
    if prob >= 40 or rain >= 1 or gust >= 45:
        return "yellow"
    return "green"


weather_days = []  # dict: d(date), icon, tmax, tmin, rain, prob, gust, level
_wj = None
try:
    _wreq = urllib.request.Request(WX_URL, headers={"User-Agent": "93H-dashboard"})
    _wraw = urllib.request.urlopen(_wreq, timeout=20).read().decode("utf-8")
    _wj = json.loads(_wraw)["daily"]
    with open(WX_CACHE, "w", encoding="utf-8") as _f:
        _f.write(_wraw)
except Exception as _e:
    print(f"天氣抓取失敗（{_e}），改用快取")
    try:
        _wj = json.load(open(WX_CACHE, encoding="utf-8"))["daily"]
    except Exception:
        _wj = None
if _wj:
    for _i, _ds in enumerate(_wj["time"]):
        _d = date.fromisoformat(_ds)
        if _d < TODAY:
            continue
        _rain = _wj["precipitation_sum"][_i] or 0
        _prob = _wj["precipitation_probability_max"][_i] or 0
        _gust = _wj["wind_gusts_10m_max"][_i] or 0
        _code = _wj["weather_code"][_i] or 0
        weather_days.append(dict(
            d=_d, icon=wx_icon(_code), rain=_rain, prob=_prob, gust=_gust,
            tmax=_wj["temperature_2m_max"][_i], tmin=_wj["temperature_2m_min"][_i],
            level=wx_level(_rain, _prob, _gust, _code)))

wx_by_date = {w["d"]: w for w in weather_days}
wx_bad = [w for w in weather_days if w["level"] == "red"]
wx_risky = [w for w in weather_days if w["level"] != "green"]

# 出工預核 × 天氣：未來7天預定的戶外工項落在紅/黃日 → 警示，並找最近綠燈日建議
weather_alerts = []
for r in dispatch_upcoming:
    _d = parse_date(r[0])
    if not _d or _d not in wx_by_date:
        continue
    _work = r[2] + " " + r[1]
    if not OUTDOOR_PAT.search(_work) or INDOOR_PAT.search(_work):
        continue
    w = wx_by_date[_d]
    if w["level"] == "green":
        continue
    greens = [x["d"] for x in weather_days if x["level"] == "green" and x["d"] != _d]
    sugg = ("建議改期或備室內替代工項；最近綠燈日 " + minguo(min(greens, key=lambda g: abs((g - _d).days)))
            if greens else "一週內無綠燈日，需備雨天方案（帆布/室內工項）")
    weather_alerts.append((r, w, sugg))

wx_week_bad = len(wx_risky) >= 5 or len(wx_bad) >= 3  # 整週爛天氣 → 提前部署橫幅


def weather_html():
    if not weather_days:
        return ""
    lv_bg = {"red": "#fbe3e0", "yellow": "#fff7e0", "green": "#e8f6ee"}
    lv_txt = {"red": "🔴不宜戶外", "yellow": "🟡有雨風險", "green": "🟢可出工"}
    cards = "".join(
        f'<div class="ms" style="background:{lv_bg[w["level"]]};min-width:118px">'
        f'<div class="ms-date">{minguo(w["d"])}(週{WEEKDAY_ZH[w["d"].weekday()]})'
        f'{"　今天" if w["d"] == TODAY else ""}</div>'
        f'<div class="ms-name">{w["icon"]}　{w["tmin"]:.0f}~{w["tmax"]:.0f}°C</div>'
        f'<div class="ms-sub">降雨{w["prob"]:.0f}%｜{w["rain"]:.0f}mm｜陣風{w["gust"]:.0f}km/h</div>'
        f'<div class="ms-cd">{lv_txt[w["level"]]}</div></div>'
        for w in weather_days)
    banner = ""
    if wx_week_bad:
        bad_str = "、".join(minguo(w["d"]) for w in wx_bad) or "—"
        banner = (f'<div style="background:#c0392b;color:#fff;border-radius:10px;padding:10px 14px;'
                  f'margin:10px 0;font-weight:700">⛈️ 未來一週天候不佳（紅燈 {len(wx_bad)} 天：{bad_str}），'
                  f'戶外工項建議提前部署：改排室內工項、雨天備案先叫料、提早通知廠商改期，避免臨時應變。</div>')
    return (banner
            + '<div class="kpi-row-label" style="padding:0 14px">🌦 一週天氣（台南安平）'
              '<span style="font-weight:400;color:#888;font-size:.85em">　Open-Meteo 預報</span></div>'
            + f'<div class="msbar" style="padding:0 14px 10px">{cards}</div>')


done_cnt = len(items) - len(open_items)


def wx_cell(d):
    w = wx_by_date.get(d)
    if not w:
        return "<td>—</td>"
    bg = {"red": "#fbe3e0", "yellow": "#fff7e0", "green": ""}[w["level"]]
    warn = "⚠️" if w["level"] != "green" else ""
    return f'<td style="background:{bg};white-space:nowrap">{w["icon"]}{w["prob"]:.0f}%{warn}</td>'


# ---- 規則3：改期待重排——同工班一列（事項×次數），資料端歷史列不動 ----
_fg = {}
for _r in dispatch_failed:
    _fg.setdefault(vbase(_r[1]) or str(_r[1]), []).append(_r)
dispatch_failed_disp = []
for _k, _rs in _fg.items():
    _rs = sorted(_rs, key=d_key)
    _last = _rs[-1]
    cores = []
    for _x in _rs:
        _c = work_core(_x[2]) or str(_x[2])[:8]
        if _c not in cores:
            cores.append(_c)
    when = minguo(d_key(_rs[-1])) if len(_rs) == 1 else f"{minguo(d_key(_rs[0]))}~{minguo(d_key(_rs[-1]))}"
    dispatch_failed_disp.append(dict(
        when=when, vendor=_k, works="、".join(cores), n=len(_rs),
        status=str(_last[4]), note=str(_last[5]), sort=d_key(_last), row=_last))
dispatch_failed_disp.sort(key=lambda x: x["sort"])

# ---- 規則2：未來排程——同廠商×同工項合併成接力鏈，一列一工項 ----
dispatch_future = [r for r in dispatch_upcoming if parse_date(r[0]) != TODAY]
_dg = {}
for _r in dispatch_future:
    _dg.setdefault((vbase(_r[1]) or str(_r[1]), work_core(_r[2])), []).append(_r)
dispatch_future_disp = []
for (_v, _c), _rs in _dg.items():
    _rs = sorted(_rs, key=d_key)
    first = _rs[0]
    if len(_rs) == 1:
        chain = esc(first[2]) + ebtn('出工預核', first.rn, first[1], f'{first[0]} {first[1]} {first[2]}', first)
    else:
        chain = " → ".join(f"<b>{minguo(d_key(x))}</b> {esc(x[2])}{ebtn('出工預核', x.rn, x[1], f'{x[0]} {x[1]} {x[2]}', x)}" for x in _rs)
    notes = []
    for x in _rs:
        if x[5] and x[5] not in notes:
            notes.append(x[5])
    srcs = []
    for x in _rs:
        if x[3] and x[3] not in srcs:
            srcs.append(x[3])
    dispatch_future_disp.append(dict(
        d=d_key(first), label=vendor_label(first), chain=chain, n=len(_rs),
        src="／".join(srcs), note="；".join(notes)))
dispatch_future_disp.sort(key=lambda x: (x["d"], x["label"]))

# 今日列（實際出工已填者：黃底=有到、紅底=未出；空白者只在「應出未回報」出現，不重複）
today_filled = [r for r in dispatch_today if str(r[4]).strip()]

# 應出未回報：紅底表（鐵則，永遠在出工區最前面顯示，不得省略）
unverified_html = ""
if dispatch_unverified:
    unverified_html = (
        '<div class="subh" style="color:#c0392b;font-weight:700">🔴 應出未回報（預定日已到，實際出工空白——請追工地補報）</div>'
        '<table><tr><th>日期</th><th>廠商(工項)</th><th>預定工作</th><th>來源</th><th>備註</th></tr>'
        + ''.join(f"<tr style=background:#fbe3e0><td>{esc(r[0])}</td><td><b>{esc(vendor_label(r))}</b>{ebtn('出工預核', r.rn, r[1], f'{r[0]} {r[1]} {r[2]}', r)}</td>"
                  f"<td>{esc(r[2])}</td><td class='note'>{esc(r[3])}</td><td class='note'>{esc(r[5])}</td></tr>"
                  for r in dispatch_unverified)
        + '</table>')

today_html = ""
if today_filled:
    def _bg(r):
        return "#fbe3e0" if FAIL_PAT.search(str(r[4])) else "#fff7e0"
    today_html = (
        '<div class="subh">📍 今日出工（已回報）</div>'
        '<table><tr><th>日期</th><th>廠商(工項)</th><th>預定工作</th><th>實際出工</th><th>備註</th></tr>'
        + ''.join(f"<tr style=background:{_bg(r)}><td>{esc(r[0])}</td><td><b>{esc(vendor_label(r))}</b>{ebtn('出工預核', r.rn, r[1], f'{r[0]} {r[1]} {r[2]}', r)}</td>"
                  f"<td>{esc(r[2])}</td><td>{esc(r[4])}</td><td class='note'>{esc(r[5])}</td></tr>"
                  for r in today_filled)
        + '</table>')

future_html = ""
if dispatch_future_disp:
    future_html = (
        '<div class="subh">📅 未來排程（出工預核；同工項多日合併為接力鏈）</div>'
        '<table><tr><th>日期</th><th>天氣</th><th>廠商(工項)</th><th>預定工作／接力</th><th>來源</th><th>備註</th></tr>'
        + ''.join(f"<tr><td>{minguo(x['d']) if x['d'] != date.max else '待定'}</td>{wx_cell(x['d'])}"
                  f"<td><b>{esc(x['label'])}</b>{'<span class=cnt style=background:#888>' + str(x['n']) + '段</span>' if x['n'] > 1 else ''}</td>"
                  f"<td>{x['chain']}</td><td class='note'>{esc(x['src'])}</td><td class='note'>{esc(x['note'])}</td></tr>"
                  for x in dispatch_future_disp)
        + '</table>')
else:
    future_html = '<div style="padding:0 14px 12px;color:#888">尚無明日以後的出工排程</div>'

# ---- 改期區 HTML ----
failed_html = ""
if dispatch_failed_disp or admin_failed:
    failed_html = (
        '<table><tr><th>原定日</th><th>工班／對象</th><th>事項（×延誤次數）</th><th>最新狀態</th><th>備註</th></tr>'
        + ''.join(f"<tr style=background:#fbe3e0><td>{esc(x['when'])}</td><td><b>{esc(x['vendor'])}</b>{ebtn('出工預核', x['row'].rn, x['row'][1], x['row'][0] + ' ' + x['row'][1] + ' ' + x['row'][2], x['row'])}</td>"
                  f"<td>{esc(x['works'])}{'<b>　×' + str(x['n']) + '次</b>' if x['n'] > 1 else ''}</td>"
                  f"<td>{esc(x['status'])}</td><td class='note'>{esc(x['note'])}</td></tr>"
                  for x in dispatch_failed_disp)
        + ''.join(f"<tr style=background:#fbe3e0><td>{esc(r[0])}</td><td><b>{esc(r[1])}</b>{ebtn('行政時程', r.rn, r[1], f'{r[0]} {r[1]} {r[2]}', r)}</td>"
                  f"<td>{esc(r[2])}</td><td>{esc(r[4])}</td><td class='note'>{esc(r[5])}</td></tr>"
                  for r in admin_failed)
        + '</table>'
        + '<div style="padding:4px 14px 10px;color:#888;font-size:.8em">規則：同一工班只列一列；原定日之後該工班已出工即自動消失；已標「已重排」的不列。</div>')

# ---- 裁決區 HTML（待裁決分頁 ＋ 自各分頁抓取的待長官現場確認）----
dec_html = ""
if decisions or site_confirm:
    dec_html = ""
    if decisions:
        dec_html += ('<div class="subh">⚖️ 待裁決（待裁決分頁）</div>'
                     '<table><tr><th>提出日</th><th>事項</th><th>工地/廠商建議</th><th>備註</th></tr>'
                     + ''.join(f"<tr><td>{esc(r[0])}</td><td><b>{esc(r[1])}</b>{ebtn('待裁決', r.rn, r[1], r[1], r)}</td><td class='note'>{esc(r[2])}</td>"
                               f"<td class='note'>{esc(r[5])}</td></tr>" for r in decisions) + '</table>')
    if site_confirm:
        dec_html += ('<div class="subh">👀 待長官現場確認／拍板（自各分頁自動抓取）</div>'
                     '<table><tr><th>分頁</th><th>事項</th><th>對象</th><th>待確認內容</th></tr>'
                     + ''.join(f"<tr><td class='cat'>{esc(s[0])}</td><td><b>{esc(s[1])}</b></td><td>{esc(s[3])}</td>"
                               f"<td class='note'>{esc(s[2])}</td></tr>" for s in site_confirm) + '</table>')

# ---- 粗工對帳 HTML ----
recon_html = ""
if recon_disp or recon_pending:
    def _f(n):
        return f"{n:g}"
    rows = ''.join(
        f"<tr style=background:{'#e8f6ee' if x['ok'] else ('#f2f4f4' if x['ok'] is None else '#fbe3e0')}><td>{esc(x['period'])}</td>"
        f"<td>{_f(x['ptot'])}（公司{_f(x['pcomp'])}/廠商{_f(x['pvend'])}）</td>"
        f"<td>{_f(x['tot'])}（公司{_f(x['comp'])}/廠商{_f(x['vend'])}）</td>"
        f"<td>{'✅相符' if x['ok'] else ('➖部分(盤點自8/26起)' if x['ok'] is None else '⚠️不符')}</td>"
        f"<td class='note'>超工累計{esc(x['over'])}｜電鑽{esc(x['drill'])}｜廢棄物{esc(x['waste'])}｜累計{esc(x['trucks'])}台</td>"
        f"<td class='note'>{esc(x['src'])}{'｜' + esc(x['note']) if x['note'] else ''}</td></tr>"
        for x in recon_disp)
    if recon_pending:
        p0, p1, t, c, v = recon_pending
        rows += (f"<tr style=background:#fff7e0><td>{minguo(p0)}~{minguo(p1)}</td><td>—（待沛誼下次報告）</td>"
                 f"<td>{_f(t)}（公司{_f(c)}/廠商{_f(v)}）</td><td>⏳</td><td class='note'></td>"
                 f"<td class='note'>盤點表本週累計，週三會後對帳</td></tr>")
    recon_html = ('<div class="subh">🧾 與會議紀錄「沛誼工作報告」週統計對帳</div>'
                  '<table><tr><th>週期間</th><th>沛誼報告(工)</th><th>盤點表(工)</th><th>核對</th><th>其他統計</th><th>來源/備註</th></tr>'
                  + rows + '</table>')

_wx_t = wx_by_date.get(TODAY)
wx_today_str = f"{_wx_t['icon']}{_wx_t['tmax']:.0f}°" if _wx_t else ""
failed_hint = ("｜最急：" + esc(dispatch_failed_disp[0]["vendor"])) if dispatch_failed_disp else ""
_sl = []
if dispatch_unverified:
    _sl.append(f'<a href="#sec-dispatch" onclick="secopen(\'sec-dispatch\')">🔴應出未回報{len(dispatch_unverified)}</a>')
if dispatch_failed_disp or admin_failed:
    _sl.append(f'<a href="#sec-failed" onclick="secopen(\'sec-failed\')">🔴改期{len(dispatch_failed_disp)+len(admin_failed)}</a>')
if decisions or site_confirm:
    _sl.append(f'<a href="#sec-dec" onclick="secopen(\'sec-dec\')">⚖️裁決{len(decisions)}+現場確認{len(site_confirm)}</a>')
if defects:
    _sl.append(f'<a href="#sec-defect" onclick="secopen(\'sec-defect\')">🛠缺失{len(defects)}</a>')
if weather_alerts:
    _sl.append(f'<a href="#sec-dispatch" onclick="secopen(\'sec-dispatch\')">🌧天氣警示{len(weather_alerts)}</a>')
if recon_disp and any(x["ok"] is False for x in recon_disp):
    _sl.append(f'<a href="#sec-labor" onclick="secopen(\'sec-labor\')">⛏粗工對帳不符</a>')
statusline = ('<div class="statusline">' + "".join(_sl) + '</div>'
              '<script>function secopen(i){var d=document.getElementById(i);if(d)d.open=true;}</script>')

page = f"""<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>93H 進度儀表板</title>
<link rel="apple-touch-icon" sizes="180x180" href="icon-180.png">
<link rel="icon" type="image/png" sizes="32x32" href="favicon-32.png">
<meta name="theme-color" content="#1a3c6e">
<meta property="og:title" content="93H 海興段 進度儀表板">
<meta property="og:description" content="好瀚建設 93H 海興段：里程碑倒數、出工預核、到期警示、缺失改善">
<meta property="og:image" content="https://jammieaiwriter-jpg.github.io/highhand/93h/icon-512.png">
<style>
:root {{ --red:#c0392b; --orange:#d68910; --blue:#2471a3; --green:#1e8449; --gray:#707b7c; }}
* {{ box-sizing:border-box; }}
body {{ font-family:"Microsoft JhengHei","PingFang TC",sans-serif; margin:0; background:#f4f6f7; color:#212f3c; }}
header {{ background:#1a3c6e; color:#fff; padding:14px 20px; }}
header h1 {{ margin:0; font-size:1.3em; }}
header .sub {{ font-size:.85em; opacity:.85; margin-top:4px; }}
.wrap {{ max-width:1100px; margin:0 auto; padding:12px; }}
.kpis {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(140px,1fr)); gap:10px; margin:12px 0; }}
.kpi {{ background:#fff; border-radius:10px; padding:12px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,.1); }}
.kpi .num {{ font-size:1.9em; font-weight:700; }}
.kpi .lbl {{ font-size:.82em; color:#555; }}
.k-red .num {{ color:var(--red); }} .k-orange .num {{ color:var(--orange); }}
.k-blue .num {{ color:var(--blue); }} .k-green .num {{ color:var(--green); }} .k-star .num {{ color:#8e44ad; }}
.msbar {{ display:flex; gap:8px; overflow-x:auto; padding:4px 0 10px; }}
.ms {{ min-width:150px; background:#fff; border-radius:10px; padding:10px; border-top:4px solid var(--gray); box-shadow:0 1px 3px rgba(0,0,0,.1); flex:1; }}
.ms-hot {{ border-top-color:var(--orange); border-top-width:6px; box-shadow:0 2px 8px rgba(214,137,16,.35); }}
.ms-hot .ms-cd {{ color:var(--orange); font-size:.92em; }}
.ms-past {{ border-top-color:var(--red); }} .ms-ok {{ border-top-color:var(--green); }}
.zone {{ margin:20px 2px 4px; font-weight:700; color:#5d6d7e; font-size:.88em; letter-spacing:2px; border-bottom:2px solid #d5dbdb; padding-bottom:4px; }}
.statusline {{ background:#fff; border-radius:10px; padding:10px 14px; margin:10px 0; box-shadow:0 1px 3px rgba(0,0,0,.1); font-size:.95em; font-weight:700; }}
.statusline a {{ text-decoration:none; color:#212f3c; margin-right:14px; white-space:nowrap; }}
.ms-date {{ font-weight:700; font-size:.9em; }}
.ms-name {{ font-size:.85em; margin:4px 0; }}
.ms-cd {{ font-size:.8em; font-weight:700; color:var(--blue); }}
.ms-note {{ font-size:.72em; color:#777; margin-top:3px; }}
.sec {{ background:#fff; border-radius:10px; margin:12px 0; box-shadow:0 1px 3px rgba(0,0,0,.1); overflow:hidden; }}
.sec summary {{ cursor:pointer; padding:12px 14px; font-weight:700; font-size:1.02em; background:#eaf0f6; }}
.sec .cnt {{ background:#1a3c6e; color:#fff; border-radius:10px; padding:1px 9px; font-size:.8em; margin-left:8px; }}
table {{ width:100%; border-collapse:collapse; font-size:.88em; }}
th {{ background:#f2f4f4; text-align:left; padding:7px 9px; border-bottom:2px solid #d5dbdb; white-space:nowrap; }}
td {{ padding:7px 9px; border-bottom:1px solid #eaeded; vertical-align:top; }}
td.cat {{ white-space:nowrap; color:#555; }}
td.note {{ color:#666; font-size:.92em; }}
.badge {{ display:inline-block; border-radius:6px; padding:2px 8px; font-size:.85em; white-space:nowrap; color:#fff; }}
.b-red {{ background:var(--red); }} .b-orange {{ background:var(--orange); }}
.b-blue {{ background:var(--blue); }} .b-green {{ background:var(--green); }} .b-gray {{ background:var(--gray); }}
.permit summary {{ background:#fdf2d0; }}
footer {{ text-align:center; color:#888; font-size:.78em; padding:16px; }}
.ms-head {{ padding:10px 14px 4px; }}
.ms-title {{ font-weight:700; font-size:1.05em; margin-right:10px; }}
.ms-meta {{ color:#555; font-size:.85em; }}
.bar {{ height:8px; background:#e5e8ec; border-radius:4px; margin:8px 0 4px; }}
.bar-in {{ height:8px; background:#1e8449; border-radius:4px; }}
.ms-sub {{ color:#888; font-size:.78em; }}
.ms-risk {{ color:#c0392b; font-size:.85em; font-weight:700; margin-top:4px; }}
.subh {{ padding:8px 14px 2px; font-weight:700; font-size:.9em; color:#333; }}
.pct {{ float:right; color:#1e8449; font-weight:700; }}
.ms-date-tag {{ margin-left:10px; color:#1a3c6e; font-weight:700; }}
.ms-cd-tag {{ margin-left:8px; color:#2471a3; font-weight:700; }}
.kpi-row-label {{ font-weight:700; margin:10px 0 4px; font-size:.95em; }}
.kpi-row-label.sub {{ color:#777; font-size:.85em; }}
.kpis-sub .kpi {{ padding:6px; }}
.kpis-sub .kpi .num {{ font-size:1.3em; }}
.kpis-sub .kpi .lbl {{ font-size:.72em; }}
.tw {{ overflow-x:auto; -webkit-overflow-scrolling:touch; }}
.tw table {{ min-width:620px; }}
.eb {{ display:none; border:0; background:transparent; cursor:pointer; font-size:.9em; padding:0 3px; opacity:.55; }}
body.can-edit .eb {{ display:inline; }}
.eb:hover {{ opacity:1; }}
tr.edited td {{ background:#eef2f5 !important; color:#888; }}
tr.edited .eb {{ opacity:1; }}
#em-bg {{ position:fixed; inset:0; background:rgba(0,0,0,.35); display:none; z-index:50; }}
#em {{ position:fixed; left:50%; bottom:0; transform:translateX(-50%); width:min(560px,100%); background:#fff; border-radius:14px 14px 0 0; box-shadow:0 -4px 20px rgba(0,0,0,.25); padding:14px 16px 18px; display:none; z-index:51; max-height:85vh; overflow:auto; }}
#em h3 {{ margin:0 0 6px; font-size:1em; }}
#em .fld {{ margin:8px 0; }}
#em select, #em input {{ width:100%; font-size:1em; padding:8px; border:1px solid #ccc; border-radius:8px; box-sizing:border-box; }}
#em .chips button {{ margin:4px 4px 0 0; padding:5px 10px; border-radius:16px; border:1px solid #2471a3; background:#fff; color:#2471a3; cursor:pointer; font-size:.9em; }}
#em .act {{ display:flex; gap:8px; margin-top:12px; }}
#em .act button {{ flex:1; padding:10px; border-radius:8px; border:0; font-size:1em; cursor:pointer; }}
#em .ok {{ background:#1e8449; color:#fff; }} #em .no {{ background:#e5e8ec; }}
#em .msg {{ font-size:.85em; color:#c0392b; min-height:1.2em; margin-top:6px; }}
@media (max-width:640px) {{
  header h1 {{ font-size:1.05em; }}
  header .sub {{ font-size:.72em; }}
  .wrap {{ padding:8px; }}
  .kpis {{ grid-template-columns:repeat(3,1fr); gap:6px; }}
  .kpi {{ padding:8px 4px; }}
  .kpi .num {{ font-size:1.4em; }}
  .kpi .lbl {{ font-size:.68em; }}
  .ms {{ min-width:128px; padding:8px; }}
  table {{ font-size:.8em; }}
  th, td {{ padding:5px 6px; }}
  .sec summary {{ font-size:.92em; padding:10px 12px; }}
  .badge {{ font-size:.78em; }}
}}
@media print {{ .sec {{ break-inside:avoid; }} }}
</style></head><body>
<header><h1>93H 海興段 進度儀表板</h1>
<div class="sub">產出：{minguo(TODAY)}（{TODAY.isoformat()}）｜資料：{esc(SRC_LABEL)}｜距取得使照 <b>{license_day} 天</b></div>
</header>
<div class="wrap">
<div class="msbar">{''.join(ms_html)}</div>
{statusline}
<div class="zone">日常追蹤</div>
<details class="sec dispatch" id="sec-dispatch"><summary>👷 出工預核與回報<span style="font-weight:400;color:#555;margin-left:8px">{esc(wx_today_str)}　今日已報{len(today_filled)}組</span><span class="cnt">{len(dispatch_future_disp)}</span></summary>{unverified_html}{today_html}{weather_html()}{future_html}</details>
{'<details class="sec"><summary>🤝 廠商協調（送樣/圖說/工序/到場，未定案）<span class="cnt">' + str(len(coord_open)) + '</span></summary><table><tr><th>日期</th><th>廠商</th><th>類型</th><th>事項</th><th>狀態</th><th>備註</th></tr>' + ''.join(f"<tr><td>{esc(r[0]) or '—'}</td><td><b>{esc(r[1])}</b>{ebtn('廠商協調', r.rn, r[1], f'{r[1]} {r[3]}', r)}</td><td>{esc(r[2])}</td><td>{esc(r[3])}</td><td>{esc(r[4])}</td><td class='note'>{esc(r[5])}</td></tr>" for r in coord_open) + '</table></details>' if coord_open else ''}
{'<details class="sec"><summary>🚚 進料追蹤<span class="cnt">' + str(len(mat_upcoming) + len(mat_unverified)) + '</span></summary><table><tr><th>日期</th><th>廠商</th><th>料項</th><th>來源</th><th>實際到料</th><th>備註</th></tr>' + ''.join(f"<tr{' style=background:#fff7e0' if parse_date(r[0])==TODAY else ''}><td>{esc(r[0]) or '待定'}</td><td><b>{esc(r[1])}</b>{ebtn('進料追蹤', r.rn, r[1], f'{r[1]} {r[2]}', r)}</td><td>{esc(r[2])}</td><td class='note'>{esc(r[3])}</td><td>{esc(r[4]) or '—'}</td><td class='note'>{esc(r[5])}</td></tr>" for r in mat_upcoming) + ''.join(f"<tr><td>{esc(r[0])}</td><td><b>{esc(r[1])}</b>{ebtn('進料追蹤', r.rn, r[1], f'{r[1]} {r[2]}', r)}</td><td>{esc(r[2])}</td><td class='note'>{esc(r[3])}</td><td>❓未記到料</td><td class='note'>{esc(r[5])}</td></tr>" for r in mat_unverified) + '</table></details>' if (mat_upcoming or mat_unverified) else ''}
{'<details class="sec"><summary>📅 行事曆與交辦（今天起）<span class="cnt">' + str(len(admin_upcoming)) + '</span></summary><table><tr><th>日期</th><th>對象/窗口</th><th>事項</th><th>來源</th><th>備註</th></tr>' + ''.join(f"<tr{' style=background:#fff7e0' if parse_date(r[0])==TODAY else ''}><td>{esc(r[0])}</td><td><b>{esc(r[1])}</b>{ebtn('行政時程', r.rn, r[1], f'{r[0]} {r[1]} {r[2]}', r)}</td><td>{esc(r[2])}</td><td class='note'>{esc(r[3])}</td><td class='note'>{esc(r[5])}</td></tr>" for r in admin_upcoming) + '</table></details>' if admin_upcoming else '<details class="sec"><summary>📅 行事曆與交辦（今天起）<span class="cnt">0</span></summary><div style="padding:0 14px 12px;color:#888">今天以後沒有排定的會議／到場／交辦</div></details>'}
<div class="zone">風險與決策</div>
{'<details class="sec" id="sec-failed" style="border-left:4px solid #c0392b"><summary>🔴 未完成／改期待重排' + failed_hint + '<span class="cnt">' + str(len(dispatch_failed_disp) + len(admin_failed)) + '</span></summary>' + failed_html + '</details>' if failed_html else ''}
{'<details class="sec" id="sec-dec" style="border-left:4px solid #7030A0"><summary>⚖️ 待長官裁決／現場確認<span class="cnt">' + str(len(decisions) + len(site_confirm)) + '</span></summary>' + dec_html + '</details>' if dec_html else ''}
{'<details class="sec" id="sec-defect"><summary>🛠️ 缺失改善追蹤（未完成）<span class="cnt">' + str(len(defects)) + '</span></summary><table><tr><th>發現日</th><th>缺失內容</th><th>位置</th><th>責任廠商</th><th>改善方式</th><th>狀態</th><th>備註</th></tr>' + ''.join(f"<tr><td>{esc(r[0])}</td><td>{esc(r[1])}{ebtn('缺失改善', r.rn, r[1], r[1], r)}</td><td>{esc(r[2])}</td><td><b>{esc(r[3])}</b></td><td class='note'>{esc(r[4])}</td><td>{esc(r[6])}</td><td class='note'>{esc(r[7])}</td></tr>" for r in defects) + '</table></details>' if defects else ''}
<div class="zone">行政</div>
{'<details class="sec"><summary>🏛️ 送審與變更（審核單位，時程可能拖）<span class="cnt">' + str(len(submissions)) + '</span></summary><table><tr><th>事項</th><th>受理/審核單位</th><th>送件日</th><th>預計核准</th><th>實際核准</th><th>備註</th></tr>' + ''.join(f"<tr><td><b>{esc(s[0])}</b>{ebtn('送審變更', s.rn, s[0], s[0], s)}</td><td>{esc(s[1])}</td><td>{esc(s[2]) or '—'}</td><td>{esc(s[3]) or '—'}</td><td>{esc(s[4]) or '—'}</td><td class='note'>{esc(s[6])}</td></tr>" for s in submissions) + '</table></details>' if submissions else ''}
{'<details class="sec"><summary>📑 使照檢附盤點（建照附款）<span class="cnt">' + str(len(rows_of("使照檢附", 5))) + '</span></summary><table><tr><th>檢附項目</th><th>附款</th><th>主辦</th><th>目前狀態</th><th>備註</th></tr>' + ''.join(f"<tr style=background:{'#e8f6ee' if re.search('已取得|已交|已完成|結案|✓', r[3]) else '#fbe3e0'}><td><b>{esc(r[0])}</b>{ebtn('使照檢附', r.rn, r[0], r[0], r)}</td><td>{esc(r[1])}</td><td>{esc(r[2])}</td><td>{esc(r[3])}</td><td class='note'>{esc(r[4])}</td></tr>" for r in rows_of("使照檢附", 5)) + '</table></details>' if "使照檢附" in wb.sheetnames else ''}
<div class="zone">里程碑</div>
{tree_html()}
<div class="zone">盤點與其他</div>
<details class="sec"><summary>📋 發包與工作事項（93H）<span class="cnt">{len(work_93h)}</span></summary>
<table><tr><th>區/類</th><th>事項</th><th>現況</th><th>下一步</th></tr>{work_html}</table></details>
{'<details class="sec" id="sec-labor"><summary>⛏️ 粗工/打石統計（依費用歸屬）<span class="cnt">' + str(len(labor)) + '</span></summary>' + recon_html + '<div class="subh">📊 費用歸屬統計</div><table><tr><th>費用歸屬</th><th>本月工數</th><th>本月金額</th><th>整場工數</th><th>整場金額</th></tr>' + ''.join(f"<tr><td><b>{esc(k)}</b></td><td>{v['m_n']:g}</td><td>{v['m_amt']:,.0f}</td><td>{v['t_n']:g}</td><td>{v['t_amt']:,.0f}</td></tr>" for k, v in sorted(labor_stats.items())) + '</table><div class="subh">最近 15 筆</div><table><tr><th>日期</th><th>工種</th><th>廠商</th><th>人數</th><th>工作內容</th><th>費用歸屬</th><th>備註</th></tr>' + ''.join(f"<tr><td>{esc(r[0])}</td><td>{esc(r[1])}</td><td>{esc(r[2])}</td><td>{esc(r[3])}</td><td class='note'>{esc(r[4])}</td><td><b>{esc(r[5])}</b></td><td class='note'>{esc(r[7])}</td></tr>" for r in labor[-15:]) + '</table></details>' if labor else ''}
{'<details class="sec"><summary>🏘️ 其他案場（總安段/福智/新家波…）<span class="cnt">' + str(len(work_other)) + '</span></summary><table><tr><th>案場</th><th>事項</th><th>現況</th><th>下一步</th></tr>' + work_other_html + '</table></details>' if work_other else ''}
</div>
<footer>93H 進度儀表板｜資料源：雲端 Sheet「93H進度盤點雲端」，桌機/雲端排程自動重產｜⭐=使照前重點｜✏️=直接改雲端狀態（下次重產反映）</footer>
<div id="em-bg"></div>
<div id="em"><h3 id="em-title"></h3><div class="ms-sub" id="em-sub"></div>
<div class="fld"><select id="em-col"></select></div>
<div class="chips" id="em-chips"></div>
<div class="fld"><input id="em-val" placeholder="輸入新值"></div>
<div class="msg" id="em-msg"></div>
<div class="act"><button class="no" id="em-no" type="button">取消</button><button class="ok" id="em-ok" type="button">送出寫回雲端</button></div></div>
<script>
(function(){{
  var API="__API__", SSID="__SSID__", FIELDS=__FIELDS__;
  var LS="93h_edits", cur=null;
  function loadLS(){{ try{{ var o=JSON.parse(localStorage.getItem(LS)||"{{}}"); var t=Date.now(); Object.keys(o).forEach(function(k){{ if(t-o[k].ts>36e5*36) delete o[k]; }}); return o; }}catch(e){{ return {{}}; }} }}
  function saveLS(o){{ try{{ localStorage.setItem(LS, JSON.stringify(o)); }}catch(e){{}} }}
  var edits=loadLS();
  // 後端版本探針：已部署 update_cell 才顯示 ✏️（避免打到舊版 doPost 誤寫回報表）
  fetch(API+"?ver=1").then(function(r){{ return r.json(); }}).then(function(j){{ if(j && j.features && j.features.indexOf("update_cell")>=0) document.body.classList.add("can-edit"); }}).catch(function(){{}});
  function keyOf(e,c){{ return e.s+"|"+e.r+"|"+c; }}
  function markRow(btn,txt){{ var tr=btn.closest("tr"); if(tr){{ tr.classList.add("edited"); }} btn.title=txt; btn.textContent="✅"; }}
  document.querySelectorAll(".eb").forEach(function(b){{
    var e=JSON.parse(b.getAttribute("data-e")); var hit=[];
    Object.keys(e.v).forEach(function(c){{ var st=edits[keyOf(e,c)]; if(st && st.prev===e.v[c]) hit.push(st.name+"："+st.val); }});
    if(hit.length) markRow(b,"已送出（等重產）："+hit.join("；"));
  }});
  var bg=document.getElementById("em-bg"), em=document.getElementById("em"), sel=document.getElementById("em-col"),
      chips=document.getElementById("em-chips"), val=document.getElementById("em-val"), msg=document.getElementById("em-msg");
  function close(){{ bg.style.display=em.style.display="none"; cur=null; }}
  function renderCol(){{
    var f=FIELDS[cur.e.s][sel.selectedIndex]; var st=edits[keyOf(cur.e,f[0])];
    val.value=(st && st.prev===cur.e.v[f[0]])? st.val : (cur.e.v[f[0]]||"");
    chips.innerHTML=""; f[2].forEach(function(p){{ var b=document.createElement("button"); b.type="button"; b.textContent=p;
      b.onclick=function(){{ val.value=p; val.focus(); var i=p.indexOf("人"); if(p.indexOf("()")>=0){{ var q=p.indexOf("(")+1; val.setSelectionRange(q,q); }} else if(i>0) val.setSelectionRange(i,i); }}; chips.appendChild(b); }});
    msg.textContent="";
  }}
  function open(btn){{
    var e=JSON.parse(btn.getAttribute("data-e")); cur={{e:e,btn:btn}};
    document.getElementById("em-title").textContent="✏️ "+e.l;
    document.getElementById("em-sub").textContent="分頁「"+e.s+"」第 "+e.r+" 列｜寫回雲端 Sheet 並記異動紀錄，儀表板下次重產（整點）反映";
    sel.innerHTML=""; FIELDS[e.s].forEach(function(f){{ var o=document.createElement("option"); o.textContent=f[1]+"（現：" + (e.v[f[0]]||"空") + "）"; sel.appendChild(o); }});
    renderCol(); bg.style.display=em.style.display="block"; setTimeout(function(){{ val.focus(); }},50);
  }}
  sel.onchange=renderCol;
  document.addEventListener("click",function(ev){{ var b=ev.target.closest(".eb"); if(b){{ ev.preventDefault(); ev.stopPropagation(); open(b); }} }});
  bg.onclick=close; document.getElementById("em-no").onclick=close;
  document.getElementById("em-ok").onclick=function(){{
    if(!cur) return; var f=FIELDS[cur.e.s][sel.selectedIndex]; var nv=val.value.trim(); var ov=cur.e.v[f[0]]||"";
    var st=edits[keyOf(cur.e,f[0])]; var expect=(st && st.prev===ov)? st.val : ov;
    if(nv===expect){{ msg.textContent="值沒有改變"; return; }}
    var ok=document.getElementById("em-ok"); ok.disabled=true; ok.textContent="寫入中…"; msg.textContent="";
    var body={{key:"93h",type:"update_cell",ssId:SSID,name:cur.e.s,row:cur.e.r,col:f[0],value:nv,expect:expect,keyCol:cur.e.kc,keyText:cur.e.k,colName:f[1],label:cur.e.l,who:"玲嬅(儀表板)"}};
    fetch(API,{{method:"POST",body:JSON.stringify(body)}}).then(function(r){{ return r.text(); }}).then(function(t){{
      var j; try{{ j=JSON.parse(t); }}catch(e){{ j={{ok:false,err:"回應非JSON（可能已寫入，請重整確認）"}}; }}
      if(j.ok){{ edits[keyOf(cur.e,f[0])]={{val:nv,prev:ov,name:f[1],ts:Date.now()}}; saveLS(edits); markRow(cur.btn,"已送出（等重產）："+f[1]+"："+nv); close(); }}
      else {{ msg.textContent="寫入失敗："+(j.err==="stale"?"雲端該格已被改成「"+j.cur+"」，請重整頁面再改":(j.err==="row moved"?"列位置已變動（雲端第"+cur.e.r+"列現在是「"+j.cur+"」），請重整頁面":j.err)); }}
      ok.disabled=false; ok.textContent="送出寫回雲端";
    }}).catch(function(e){{ msg.textContent="連線失敗："+e; ok.disabled=false; ok.textContent="送出寫回雲端"; }});
  }};
}})();
</script>
<script>
(function() {{
  var loaded = Date.now();
  document.addEventListener("visibilitychange", function() {{
    if (document.visibilityState === "visible" && Date.now() - loaded > 5 * 60 * 1000) {{
      location.reload();
    }}
  }});
}})();
</script>
</body></html>"""

page = page.replace("<table", "<div class=\"tw\"><table").replace("</table>", "</table></div>")
page = (page.replace("__API__", REPORT_API).replace("__SSID__", SHEET_ID)
        .replace("__FIELDS__", json.dumps({k: [[c, n, p] for c, n, p in v] for k, v in EDIT_FIELDS.items()}, ensure_ascii=False)))

with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(page)
print(f"OK -> {OUT_HTML}")

# ---- 工地回報頁用：未來7天預計出工 plan.json（隨 bat 一起推上 GitHub Pages）----
plan_days = []
for i in range(0, 8):
    d = TODAY + timedelta(days=i)
    _pi = [{"vendor": r[1], "work": r[2], "source": r[3], "note": r[5]}
           for r in dispatch_all if parse_date(r[0]) == d]
    if _pi:
        plan_days.append({"date": d.isoformat(), "roc": minguo(d), "items": _pi})
_plan_dir = os.environ.get("PLAN_DIR") or os.path.join(BASE, "highhand_repo", "93h", "report")
PLAN_PATH = os.path.join(_plan_dir, "plan.json")
if os.path.isdir(os.path.dirname(PLAN_PATH)):
    with open(PLAN_PATH, "w", encoding="utf-8") as f:
        json.dump({"generated": TODAY.isoformat(), "days": plan_days}, f, ensure_ascii=False, indent=1)
    print(f"OK -> {PLAN_PATH}")
print(f"統計：逾期{len(overdue)} 兩週內{len(soon)} 可施作{len(ready)} 施作中{len(doing)} 使照前未完{len(permit_open)} "
      f"排程{len(later)} 等前置{len(waiting)} 完成{done_cnt}/{len(items)}｜"
      f"應出未回報{len(dispatch_unverified)} 改期{len(dispatch_failed_disp)}+{len(admin_failed)} "
      f"裁決{len(decisions)}+現場確認{len(site_confirm)} 缺失{len(defects)} 協調{len(coord_open)} 行事曆{len(admin_upcoming)}")
