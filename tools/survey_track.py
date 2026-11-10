# -*- coding: utf-8 -*-
"""會勘點 → 缺失改善：純邏輯（不連網、不寫 Sheet）。

會勘分頁一列一個巡視點（日期/時間/巡視/陪同/序/位置/交辦/缺失/照片數/填報人）。
簽核用語（Ok、已確認、再議…）不建列。已有來源標記，或日期＋位置＋內容對得上
既有缺失改善列的，不重複新增、也不改那一列。
"""
import re
from datetime import date, datetime, timedelta

SURVEY_SHEET = "會勘"
DEFECT_SHEET = "缺失改善"
DEFECT_HEADERS = ["發現日", "缺失內容", "位置", "責任廠商", "改善方式", "預計完成", "狀態", "備註"]
OPEN_STATUS = "待分派"

# 整格（或括號裡每一段）都是這類詞才算簽核，句子裡順帶出現「再議」仍要追蹤。
_SIGNOFF = re.compile(
    r"^(?:ok|okay|ｏｋ|已確認|再議|無|無缺失|無異狀|沒問題|無問題|確認ok|可接受|"
    r"結案|暫不處理|維持原案|照舊|無須處理|不用處理|無須改善|沒有問題)$",
    re.IGNORECASE)
_TAG = re.compile(r"會勘\s*((?:0?1[0-9]{2}/)?\d{1,2}/\d{1,2})\s*[#＃]\s*(\d+)")
_CLOSED = re.compile(r"^\s*(已完成|完成|結案|已結案)")
_STOP = {
    "工程", "施工", "缺失", "改善", "現場", "確認", "施作", "問題", "處理", "需要",
    "已經", "目前", "公司", "工地", "廠商", "預計", "完成", "進行", "今天", "明日",
    "請", "已", "未", "及", "與", "或", "的", "了",
}
# 會勘表頭 → 內部欄名。交辦/缺失都進內容。
_HEADER = {
    "日期": "date", "發現日": "date",
    "時間": "time",
    "巡視": "inspector", "巡視人": "inspector",
    "陪同": "accompany",
    "序": "seq", "編號": "seq", "項次": "seq", "序號": "seq",
    "位置": "loc",
    "交辦": "assign", "交辦事項": "assign",
    "缺失": "defect", "缺失內容": "defect", "內容": "defect", "說明": "defect",
    "照片數": "photos", "照片": "photos",
    "填報人": "reporter", "填報": "reporter",
}


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else ("%g" % v)
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return str(v).strip()


def roc_date(d):
    """與盤點表一致：0115/9/23（年前置 0，月日不補 0）。"""
    return f"0{d.year - 1911}/{d.month}/{d.day}"


def source_tag(d, seq):
    return f"會勘 {roc_date(d)}#{int(seq)}"


def _flat(s):
    s = norm(s).replace("（", "(").replace("）", ")").replace("　", "")
    s = re.sub(r"[\s,，、。．.！!？?；;：:／/]+", "", s)
    return s.lower()


def is_signoff(text):
    """整段沒有可追蹤的內容，或每一段都只是簽核詞。"""
    raw = norm(text)
    if not raw:
        return True
    parts = [p for p in re.findall(r"[^()（）]+", _flat(raw)) if p]
    if not parts:
        return True
    return all(_SIGNOFF.match(p) for p in parts)


def is_closed(status):
    s = norm(status)
    return bool(_CLOSED.match(s) or "結案" in s)


def parse_loose_date(value, today=None):
    """西元 date。接受 datetime、2026-09-23、0115/9/23、115/9/23、9/23、115/11月中。"""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = norm(value)
    if not text:
        return None
    m = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", text)
    if m:
        return _ymd(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"0?(1[0-9]{2})[./](\d{1,2})[./](\d{1,2})", text)
    if m:
        return _ymd(int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3)))
    m = re.search(r"0?(1[0-9]{2})[./](\d{1,2})\s*月(底|初|中)?", text)
    if m:
        part = m.group(3) or "底"
        dd = {"初": 5, "中": 15}.get(part, 28)
        return _ymd(int(m.group(1)) + 1911, int(m.group(2)), dd)
    m = re.search(r"(?<!\d)(\d{1,2})[./](\d{1,2})(?!\d)", text)
    if m and today:
        return _bare_md(int(m.group(1)), int(m.group(2)), today)
    return None


def _ymd(y, mo, dd):
    try:
        return date(y, mo, dd)
    except ValueError:
        return None


def _bare_md(month, day, today):
    d = _ymd(today.year, month, day)
    if d and d > today + timedelta(days=2):
        d = _ymd(today.year - 1, month, day)
    return d


def source_keys(text, today):
    """備註裡的「會勘 日期#序」→ {(date, seq), ...}。沒有 #序 的『8/25會勘』不算。"""
    keys = set()
    for m in _TAG.finditer(norm(text)):
        d = parse_loose_date(m.group(1), today)
        if d:
            keys.add((d, int(m.group(2))))
    return keys


def is_survey_origin(note, today):
    return bool(source_keys(note, today))


def attention(due_text, owner_text, today):
    """(逾期, 缺責任廠商, 缺預計完成)。"""
    due = parse_loose_date(due_text, today)
    overdue = bool(due and due < today)
    missing_owner = not norm(owner_text)
    missing_due = not norm(due_text)
    return overdue, missing_owner, missing_due


def _tokens(text):
    out = []
    for t in re.findall(r"[A-Za-z0-9]+|[\u4e00-\u9fff]{2,}", norm(text)):
        tl = t.lower()
        if tl in _STOP:
            continue
        if re.fullmatch(r"[a-z0-9]+", tl) and len(tl) < 2:
            continue
        out.append(tl)
    return out


def _contains(a, b):
    fa, fb = _flat(a), _flat(b)
    if len(fa) < 4 or len(fb) < 4:
        return False
    return fa in fb or fb in fa


def _loc_compatible(a, b):
    fa, fb = _flat(a), _flat(b)
    if len(fa) < 1 or len(fb) < 1:
        return False
    return fa in fb or fb in fa


def _cell(row, i):
    return row[i] if i < len(row) else ""


def same_finding(point, defect, today):
    """既有列（含人工抄入、含已完成）是否就是這個會勘點。對上就不新增、也不改原列。"""
    content_d = _cell(defect, 1)
    loc_d = _cell(defect, 2)
    date_d = parse_loose_date(_cell(defect, 0), today)
    date_ok = bool(point["date"] and date_d and point["date"] == date_d)
    loc_ok = _loc_compatible(point["loc"], loc_d)
    shared = set(_tokens(point["content"])) & set(_tokens(content_d))
    contain = _contains(point["content"], content_d)
    toks_s, toks_d = set(_tokens(point["content"])), set(_tokens(content_d))
    jac = (len(shared) / len(toks_s | toks_d)) if (toks_s and toks_d) else 0.0
    strong = contain or jac >= 0.55 or len(shared) >= 2
    if strong and date_ok:
        return True
    if strong and loc_ok and (contain or jac >= 0.45 or len(shared) >= 2):
        return True
    if date_ok and loc_ok and shared and max(len(t) for t in shared) >= 3:
        return True
    return False


def _header_map(header):
    mapping = {}
    for i, cell in enumerate(header):
        key = _HEADER.get(norm(cell).replace(" ", ""))
        if key and key not in mapping:
            mapping[key] = i
    return mapping


def _looks_like_header(header):
    mapped = _header_map(header)
    return "loc" in mapped or "defect" in mapped or "assign" in mapped or "date" in mapped


def _combine(assign, defect):
    a, d = norm(assign), norm(defect)
    if is_signoff(a) and is_signoff(d):
        return a or d
    if is_signoff(a):
        return d
    if is_signoff(d):
        return a
    if not a:
        return d
    if not d:
        return a
    if _flat(a) in _flat(d):
        return d
    if _flat(d) in _flat(a):
        return a
    return a + "／" + d


def _seq_of(value, fallback):
    text = norm(value)
    m = re.search(r"\d+", text)
    if not m:
        return int(fallback)
    return int(m.group(0))


def parse_survey_grid(grid, today):
    """會勘表（含標題列）→ 巡視點 list。簽核列先留著，交給 plan 標成略過。"""
    rows = [list(r) for r in (grid or []) if any(norm(c) for c in r)]
    if not rows:
        return []
    start = 0
    mapping = {}
    if _looks_like_header(rows[0]):
        mapping = _header_map(rows[0])
        start = 1
    else:
        # 使用者描述的欄序：日期/時間/巡視/陪同/序/位置/交辦/缺失/照片數/填報人
        mapping = {"date": 0, "time": 1, "inspector": 2, "accompany": 3, "seq": 4,
                   "loc": 5, "assign": 6, "defect": 7, "photos": 8, "reporter": 9}

    def col(row, name):
        i = mapping.get(name)
        if i is None or i >= len(row):
            return ""
        return row[i]

    points = []
    per_day = {}
    for raw in rows[start:]:
        content = _combine(col(raw, "assign"), col(raw, "defect"))
        loc = norm(col(raw, "loc"))
        if not norm(content) and not loc:
            continue
        d = parse_loose_date(col(raw, "date"), today)
        if d is None:
            continue
        per_day[d] = per_day.get(d, 0) + 1
        seq = _seq_of(col(raw, "seq"), per_day[d])
        points.append({
            "date": d,
            "seq": seq,
            "loc": loc,
            "content": norm(content),
            "inspector": norm(col(raw, "inspector")),
            "reporter": norm(col(raw, "reporter")),
            "tag": source_tag(d, seq),
        })
    # 同一天同一個序只留第一筆有內容的（回傳重複列）
    seen = set()
    uniq = []
    for p in points:
        key = (p["date"], p["seq"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    return uniq


def defect_row(point):
    """8 欄，對齊缺失改善。責任廠商、改善方式、預計完成留白，狀態固定待分派。"""
    note = point["tag"]
    extra = []
    if point.get("inspector"):
        extra.append("巡視:" + point["inspector"])
    if point.get("reporter"):
        extra.append("填報:" + point["reporter"])
    if extra:
        note = note + "；" + "；".join(extra)
    return [
        roc_date(point["date"]),
        point["content"],
        point["loc"],
        "",
        "",
        "",
        OPEN_STATUS,
        note,
    ]


def plan_additions(points, defects, today):
    """回傳 {add: [8欄], skip: [{reason, tag, content, loc, match_row}]}。

    不修改 defects。已完成的列若對得上，也算已追蹤，不會再開一列。
    """
    defects = [list(r) for r in defects]
    tagged = set()
    for i, row in enumerate(defects):
        tagged |= {(k, i) for k in source_keys(_cell(row, 7), today)}
    tagged_keys = {k for k, _ in tagged}

    add = []
    skip = []
    planned = set(tagged_keys)
    for p in points:
        key = (p["date"], p["seq"])
        if is_signoff(p["content"]):
            skip.append({"reason": "signoff", "tag": p["tag"], "content": p["content"], "loc": p["loc"]})
            continue
        if key in planned:
            skip.append({"reason": "tagged", "tag": p["tag"], "content": p["content"], "loc": p["loc"]})
            continue
        match_row = None
        for i, row in enumerate(defects):
            if same_finding(p, row, today):
                match_row = i + 2  # 標題是第 1 列
                break
        if match_row is not None:
            skip.append({"reason": "manual", "tag": p["tag"], "content": p["content"],
                         "loc": p["loc"], "match_row": match_row})
            planned.add(key)
            continue
        add.append(defect_row(p))
        planned.add(key)
    return {"add": add, "skip": skip}


def format_report(result, dry, source_note=""):
    """給班次與人工看的純文字。"""
    adds = result["add"]
    skip = result["skip"]
    lines = []
    lines.append("會勘→缺失改善（乾跑，未寫入）" if dry else "會勘→缺失改善")
    if source_note:
        lines.append(source_note)
    verb = "將新增" if dry else "新增"
    lines.append(f"{verb} {len(adds)} 列" + ("：" if adds else "。"))
    for row in adds:
        lines.append(f"  + {row[0]}  {row[2] or '（位置未填）'}  {row[1]}  狀態={row[6]}  備註={row[7]}")
    by = {}
    for s in skip:
        by.setdefault(s["reason"], []).append(s)
    labels = {"signoff": "略過簽核/無內容", "tagged": "略過已有來源標記", "manual": "略過已在缺失改善（不覆寫）"}
    for key, label in labels.items():
        items = by.get(key) or []
        if not items:
            continue
        lines.append(f"{label} {len(items)}：")
        for s in items[:12]:
            extra = f"（對上約第 {s['match_row']} 列）" if s.get("match_row") else ""
            lines.append(f"  - {s['tag']} {s['loc']} {s['content']}{extra}")
        if len(items) > 12:
            lines.append(f"  …另 {len(items) - 12} 列")
    if not dry and adds:
        lines.append("既有列不改狀態、不改內容，只在表尾追加。")
    elif dry:
        lines.append("這是乾跑，雲端一格都沒寫。")
    return "\n".join(lines)
