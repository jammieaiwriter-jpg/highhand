# -*- coding: utf-8 -*-
"""會勘 → 缺失改善。不連網。python3 tools/test_survey_track.py"""
import os
import sys
import tempfile
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import survey_track as st

TODAY = date(2026, 11, 10)

HEADER = ["日期", "時間", "巡視", "陪同", "序", "位置", "交辦", "缺失", "照片數", "填報人"]

SAMPLE = [
    HEADER,
    ["11/6", "10:00", "總經理", "玲嬅", 1, "1F", "YKK通風門", "門扇無法密合", 2, "建麟"],
    ["11/6", "10:10", "總經理", "", 2, "R3", "", "R3排風扇未固定", 1, "建麟"],
    ["11/6", "10:20", "總經理", "", 3, "3F", "Ok", "", 0, "建麟"],
    ["11/6", "10:30", "總經理", "", 4, "4F", "已確認", "", 0, "建麟"],
    ["11/6", "10:40", "總經理", "", 5, "5F", "再議", "", 0, "建麟"],
    ["11/6", "10:50", "總經理", "", 6, "6F", "已確認(Ok)", "", 0, "建麟"],
    ["11/5", "14:00", "總經理", "", 1, "管理室", "", "天花板滲水", 3, "建麟"],
    ["11/5", "14:10", "總經理", "", 2, "S1", "石階", "高低差需修", 1, "建麟"],
    ["11/5", "14:20", "總經理", "", 3, "3FA3", "", "客廁台面盆矽利康位置不當", 0, "建麟"],
    ["11/5", "14:30", "總經理", "", 4, "14F", "請確認YKK門下緣", "間隙偏大", 1, "建麟"],
]

# 盤點表現有的一列（人工抄過 11/5 會勘，備註沒有「會勘 日期#序」）
MANUAL = [
    "0115/11/5",
    "3FA3客廁台面盆矽利康位置不當(合檠原打在臉盆旁)",
    "3FA3", "鴻成(水電)", "請水電改打在臉盆底", "", "待處理",
    "11/5會勘(總經理交辦)",
]
OTHER = [
    "0115/9/23", "陽台天花板油漆SA2000手刷痕明顯", "各戶工作陽台",
    "品豪油漆(柯)", "", "", "已確認(Ok)", "9/23會議總經理交辦7.",
]
OLD_NOTE = [
    "0115/8/26", "樓梯間乳膠漆脫落", "甲乙梯樓梯間", "待議", "", "115/11月中", "進行中",
    "總額157,949；8/25會勘工地概算15-16萬",
]


def _plan(extra_defects=None, grid=None):
    points = st.parse_survey_grid(grid if grid is not None else SAMPLE, TODAY)
    defects = [list(MANUAL), list(OTHER), list(OLD_NOTE)]
    if extra_defects:
        defects.extend(extra_defects)
    before = [list(r) for r in defects]
    result = st.plan_additions(points, defects, TODAY)
    return points, defects, before, result


class SurveyTrackTest(unittest.TestCase):
    def test_signoff_and_real_findings(self):
        self.assertTrue(st.is_signoff("Ok"))
        self.assertTrue(st.is_signoff("已確認"))
        self.assertTrue(st.is_signoff("再議"))
        self.assertTrue(st.is_signoff("已確認(Ok)"))
        self.assertTrue(st.is_signoff("  "))
        self.assertFalse(st.is_signoff("請確認YKK門下緣／間隙偏大"))
        self.assertFalse(st.is_signoff("Ok,填縫改善"))
        self.assertFalse(st.is_signoff("外牆清洗再議"))
        self.assertTrue(st.is_deferred("外牆清洗再議"))
        self.assertTrue(st.is_deferred("使照後要換，再議"))
        self.assertFalse(st.is_deferred("Ok,填縫改善"))

    def test_adds_untracked_and_skips_signoff_and_manual(self):
        _, defects, before, result = _plan()
        self.assertEqual(defects, before)  # 不改既有列
        contents = [r[1] for r in result["add"]]
        self.assertIn("YKK通風門／門扇無法密合", contents)
        self.assertIn("R3排風扇未固定", contents)
        self.assertIn("天花板滲水", contents)
        self.assertIn("石階／高低差需修", contents)
        self.assertIn("請確認YKK門下緣／間隙偏大", contents)
        self.assertNotIn("客廁台面盆矽利康位置不當", contents)
        for row in result["add"]:
            self.assertEqual(row[3], "")
            self.assertEqual(row[4], "")
            self.assertEqual(row[5], "")
            self.assertEqual(row[6], "待分派")
            self.assertTrue(row[7].startswith("會勘 "))
            self.assertIn("#", row[7])
        reasons = {s["reason"] for s in result["skip"]}
        self.assertIn("signoff", reasons)
        self.assertIn("manual", reasons)
        signoff_text = " ".join(s["content"] for s in result["skip"] if s["reason"] == "signoff")
        self.assertIn("Ok", signoff_text)
        self.assertIn("已確認", signoff_text)
        self.assertIn("再議", signoff_text)
        manual = [s for s in result["skip"] if s["reason"] == "manual"]
        self.assertTrue(any("矽利康" in s["content"] for s in manual))
        # 油漆那列不是 YKK，不能被誤判成已追蹤
        self.assertFalse(any("YKK" in s["content"] and s["reason"] == "manual" for s in result["skip"]))

    def test_idempotent_second_run(self):
        _, _, _, first = _plan()
        _, _, _, second = _plan(extra_defects=first["add"])
        self.assertEqual(second["add"], [])
        self.assertTrue(all(s["reason"] == "tagged" or s["reason"] == "signoff" or s["reason"] == "manual"
                            for s in second["skip"]))
        self.assertGreaterEqual(sum(1 for s in second["skip"] if s["reason"] == "tagged"), len(first["add"]))

    def test_completed_manual_row_is_not_reopened(self):
        done = [
            "0115/9/23", "YKK通風門門扇無法密合", "1F", "YKK", "", "0115/9/30", "已完成",
            "9/23現場",
        ]
        _, _, _, result = _plan(extra_defects=[done])
        self.assertFalse(any("YKK通風門" in r[1] for r in result["add"]))

    def test_closed_status_drops_off_open_list(self):
        open_row = ["0115/9/23", "天花板滲水", "管理室", "", "", "", "待分派", "會勘 0115/9/23#1；巡視:總經理"]
        done_row = ["0115/9/23", "R3排風扇未固定", "R3", "水電", "", "0115/10/1", "已完成", "會勘 0115/9/23#2"]
        closed = ["0115/11/5", "石階", "S1", "石材", "", "", "結案", "會勘 0115/11/5#2"]
        self.assertTrue(st.is_survey_origin(open_row[7], TODAY))
        self.assertFalse(st.is_survey_origin("8/25會勘工地概算15-16萬", TODAY))
        self.assertFalse(st.is_survey_origin("11/5會勘(總經理交辦)", TODAY))
        kept = [r for r in (open_row, done_row, closed)
                if st.is_survey_origin(r[7], TODAY) and not st.is_closed(r[6])]
        self.assertEqual(kept, [open_row])

    def test_attention_flags(self):
        overdue, miss_o, miss_d = st.attention("0115/10/1", "", TODAY)
        self.assertTrue(overdue and miss_o and not miss_d)
        overdue, miss_o, miss_d = st.attention("", "鴻成", TODAY)
        self.assertFalse(overdue)
        self.assertFalse(miss_o)
        self.assertTrue(miss_d)
        overdue, miss_o, miss_d = st.attention("0115/11/30", "鴻成", TODAY)
        self.assertFalse(any((overdue, miss_o, miss_d)))
        overdue, _, _ = st.attention("115/11月中", "待議", TODAY)
        self.assertFalse(overdue)  # 11/15 還沒到

    def test_tag_formats_match(self):
        self.assertEqual(st.source_keys("會勘 9/23#2；巡視:總經理", TODAY),
                         {(date(2026, 9, 23), 2)})
        self.assertEqual(st.source_keys("會勘 0115/11/5#1", TODAY),
                         {(date(2026, 11, 5), 1)})
        self.assertEqual(st.source_tag(date(2026, 9, 23), 2), "會勘 0115/9/23#2")

    def test_positional_grid_without_header(self):
        points = st.parse_survey_grid(SAMPLE[1:], TODAY)
        self.assertTrue(any(p["content"].startswith("YKK通風門") for p in points))

    def test_duplicate_seq_kept_once(self):
        grid = [HEADER, SAMPLE[1], SAMPLE[1]]
        points = st.parse_survey_grid(grid, TODAY)
        self.assertEqual(len(points), 1)

    def test_combined_column_and_weekday_date(self):
        """工地回報的會勘分頁是單一欄「交辦/缺失」，日期帶星期。"""
        grid = [
            ["日期", "時間", "巡視", "陪同", "序", "位置", "交辦/缺失", "照片數", "填報人"],
            ["9/23 (三)", None, "總經理", None, 1, "R3", "電梯機房要裝排風扇", 0, "建麟"],
            ["9/23 (三)", None, "總經理", None, 2, "3F", "Ok", 0, "建麟"],
            ["11/5 (四)", None, "總經理", None, 3, None,
             "【3FA3 客廁台面盆】合檠目前矽利康打在臉盆旁，請水電改打在臉盆底", 0, "建麟"],
            ["11/5 (四)", None, "總經理", None, 4, None, "【1F路樹】使照後要換，再議", 0, "建麟"],
        ]
        points = st.parse_survey_grid(grid, TODAY)
        by_seq = {p["seq"]: p for p in points}
        self.assertEqual(by_seq[1]["content"], "電梯機房要裝排風扇")
        self.assertEqual(by_seq[1]["date"], date(2026, 9, 23))
        self.assertTrue(st.is_signoff(by_seq[2]["content"]))
        self.assertEqual(by_seq[3]["loc"], "3FA3 客廁台面盆")
        self.assertIn("矽利康", by_seq[3]["content"])
        self.assertFalse(st.is_signoff(by_seq[4]["content"]))
        result = st.plan_additions(points, [list(MANUAL)], TODAY)
        added = " ".join(r[1] for r in result["add"])
        self.assertNotIn("電梯機房要裝排風扇", added)  # 9/23，早於入帳日起算
        self.assertNotIn("矽利康", added)
        self.assertTrue(any(s["reason"] == "before" and "排風扇" in s["content"] for s in result["skip"]))
        self.assertTrue(any(s["reason"] == "manual" and "矽利康" in s["content"] for s in result["skip"]))
        self.assertTrue(any(s["reason"] == "defer" and "路樹" in s["loc"] for s in result["skip"]))

    def test_import_cutoff_include_and_exclude(self):
        """9/23 只入填縫改善；11/5 入帳但路樹（再議）不入；11/5 之後照常。"""
        grid = [
            ["日期", "時間", "巡視", "陪同", "序", "位置", "交辦/缺失", "照片數", "填報人"],
            ["9/23 (三)", None, "總經理", None, 3, "1F電梯前異材質石材", "Ok,填縫改善", 0, "建麟"],
            ["9/23 (三)", None, "總經理", None, 5, "YKK三合一通風門", "9F以上雙扣，8F以下單扣", 0, "建麟"],
            ["9/23 (三)", None, "總經理", None, 6, "R3", "電梯機房要裝排風扇", 0, "建麟"],
            ["11/5 (四)", None, "總經理", None, 9, None, "【1F路樹】使照後要換，再議", 0, "建麟"],
            ["11/5 (四)", None, "總經理", None, 12, None, "【S1店舖安北路出入口】做石階（3踏）", 0, "建麟"],
            ["11/6 (五)", None, "總經理", None, 1, "RF", "新發現的漏水要補", 0, "建麟"],
            ["11/6 (五)", None, "總經理", None, 2, "B1", "這案再議", 0, "建麟"],
        ]
        result = st.plan_additions(st.parse_survey_grid(grid, TODAY), [], TODAY)
        added = [(r[2], r[1]) for r in result["add"]]
        self.assertIn(("1F電梯前異材質石材", "Ok,填縫改善"), added)
        self.assertIn(("S1店舖安北路出入口", "【S1店舖安北路出入口】做石階（3踏）"), added)
        self.assertIn(("RF", "新發現的漏水要補"), added)
        self.assertFalse(any("通風門" in c or "排風扇" in c or "路樹" in c or "再議" in c for _, c in added))
        reasons = {s["tag"]: s["reason"] for s in result["skip"]}
        self.assertEqual(reasons["會勘 0115/9/23#5"], "before")
        self.assertEqual(reasons["會勘 0115/9/23#6"], "before")
        self.assertEqual(reasons["會勘 0115/11/5#9"], "defer")
        self.assertEqual(reasons["會勘 0115/11/6#2"], "defer")

    def test_phrase_already_in_ledger_note_is_not_readded(self):
        grid = [
            ["日期", "時間", "巡視", "陪同", "序", "位置", "交辦/缺失", "照片數", "填報人"],
            ["11/5 (四)", None, "總經理", None, 10, None, "【RF滴水條】已切割磁磚、丈量尺寸", 1, "建麟"],
        ]
        existing = ["0115/9/9", "RF滴水條破口已加封板仍有縫隙", "RF", "新美", "", "", "施作中",
                    "11/5已切割磁磚、丈量尺寸"]
        result = st.plan_additions(st.parse_survey_grid(grid, TODAY), [existing], TODAY)
        self.assertEqual(result["add"], [])
        self.assertEqual(result["skip"][0]["reason"], "manual")

    def test_report_text_lists_adds(self):
        _, _, _, result = _plan()
        text = st.format_report(result, dry=True, source_note="來源：樣本")
        self.assertIn("乾跑，未寫入", text)
        self.assertIn("YKK通風門", text)
        self.assertIn("待分派", text)
        self.assertIn("會勘 0115/11/6#1", text)
        self.assertNotIn("+ 0115/11/6  3F  Ok", text)


class CliDryRunTest(unittest.TestCase):
    def test_cli_dry_run_sample_files(self):
        import subprocess
        from openpyxl import Workbook
        tmp = tempfile.mkdtemp()
        ledger = os.path.join(tmp, "ledger.xlsx")
        survey = os.path.join(tmp, "survey.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "缺失改善"
        ws.append(st.DEFECT_HEADERS)
        ws.append(MANUAL)
        ws.append(OTHER)
        wb.save(ledger)
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = "會勘"
        for row in SAMPLE:
            ws2.append(row)
        wb2.save(survey)
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "h93_tools.py")
        r = subprocess.run(
            [sys.executable, script, "survey", "--dry-run", "--from", ledger, "--survey", survey],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertIn("乾跑，未寫入", r.stdout)
        self.assertIn("YKK通風門", r.stdout)
        self.assertIn("天花板滲水", r.stdout)
        self.assertNotIn("import_sheet", r.stdout + r.stderr)


if __name__ == "__main__":
    unittest.main()
