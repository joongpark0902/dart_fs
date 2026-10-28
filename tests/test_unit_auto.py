"""표시 단위 자동 판정(unit_display=0)·설정 셀(cfg_unit)로 mFS 즉시 전환."""
import json, os, tempfile, unittest
from datetime import date
from types import SimpleNamespace
from core import pipeline as P, mapping as M
from core.model import FsRecord, Period
from excel import compose as C, refresh as R
from tests.test_pipeline import FakeClient


def _settings_row(s, label):
    return next(r for (r, c), cell in s.cells.items() if c == 2 and cell.v == label)


def _hand_col(bs_unit, unit_display=0, fs_divs=("OFS",)):
    """BS 단위 bs_unit 인 손수 만든 Collected(자산·부채·자본총계)."""
    pers = [Period("FY2024", date(2024, 12, 31), "annual", 26), Period("FY2025", date(2025, 12, 31), "annual", 27)]
    cfg = P.BuildConfig("X", "X", date(2026, 10, 26), 2024, 2025, False, list(fs_divs), unit_display, None, "default")
    recs = []
    for pk, (a, l, e) in (("FY2024", (100, 60, 40)), ("FY2025", (110, 70, 40))):
        for i, (lab, v) in enumerate((("자산총계", a), ("부채총계", l), ("자본총계", e))):
            recs.append(FsRecord(pk, "OFS", "BS", i, lab, lab, v, bs_unit, None, None, "r", "XML", "cur"))
    col = P.Collected(cfg, SimpleNamespace(periods=pers, reports={}), ["FY2024", "FY2025"],
                      fs_records=recs, mapping=M.load())
    col.units[("OFS", "BS")] = bs_unit
    return col


class TestJudge(unittest.TestCase):
    def test_judge_bs_then_is_then_one(self):
        col = _hand_col(1000)
        self.assertEqual(C.judge_unit(col), (1000, "rFS_별도 재무상태표 단위 행"))
        col.units = {("OFS", "IS"): 1000000}
        self.assertEqual(C.judge_unit(col), (1000000, "rFS_별도 손익계산서 단위 행"))
        col.units = {}
        self.assertEqual(C.judge_unit(col)[0], 1)

    def test_primary_is_consolidated_when_selected(self):
        col = _hand_col(1000, fs_divs=("CFS", "OFS"))
        col.units = {("OFS", "BS"): 1, ("CFS", "BS"): 1000}
        self.assertEqual(C.judge_unit(col), (1000, "rFS_연결 재무상태표 단위 행"))

    def test_from_dict_keeps_auto_zero(self):
        d = {"corp_code": "X", "as_of": "2026-10-26", "start_year": 2024, "end_year": 2025, "unit_display": 0}
        self.assertEqual(P.BuildConfig.from_dict(d).unit_display, 0)
        d.pop("unit_display")
        self.assertEqual(P.BuildConfig.from_dict(d).unit_display, 0)          # 기본 = 자동
        d["unit_display"] = 1000
        self.assertEqual(P.BuildConfig.from_dict(d).unit_display, 1000)


class TestComposeAutoThousand(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wb = C.compose(_hand_col(1000), None, "p", "s")

    def test_mfs_formulas_scale_raw_unit_over_cfg_unit(self):
        mfs = self.wb.get("mFS")
        rows = {cell.v: r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL}
        f = mfs.cells[(rows["자산총계"], C.RAW_FIRST_VALUE_COL)].v
        self.assertRegex(f, r"^='rFS_별도'!E\d+\*1000/cfg_unit$")

    def test_settings_cfg_unit_holds_judged_divisor(self):
        s = self.wb.get("설정")
        r = int(s.names["cfg_unit"][1:])
        self.assertEqual(s.cells[(r, 3)].v, 1000)
        js = json.loads(s.cells[(int(s.names["cfg_json"][1:]), 3)].v)
        self.assertEqual((js["unit_display"], js["unit_judged"]), (0, 1000))
        self.assertIn("자동 판정: 천원 (근거: rFS_별도 재무상태표 단위 행)", s.cells[(_settings_row(s, "단위 판정"), 3)].v)
        self.assertEqual((self.wb.unit_value, self.wb.unit_auto), (1000, True))

    def test_checks_compare_in_won_rounded_to_thousand(self):
        chk = self.wb.get("검증")
        f = next(c.v for c in chk.cells.values() if c.formula and str(c.v).startswith("=ROUND("))
        self.assertRegex(f, r"^=ROUND\('mFS'!E\d+\*cfg_unit/1000,0\)=ROUND\(\('mFS'!E\d+\+'mFS'!E\d+\)\*cfg_unit/1000,0\)$")
        self.assertFalse([c.v for c in self.wb.get("mFS").cells.values() if c.formula and str(c.v).startswith("=ROUND(")])

    def test_log_has_judgement(self):
        msgs = [c.v for c in self.wb.get("로그").cells.values()]
        self.assertIn("표시 단위 자동 판정: 천원 (근거: rFS_별도 재무상태표 단위 행)", msgs)


class TestComposeAutoOb(unittest.TestCase):
    """오비 픽스처(XML 경로, 재무상태표 단위 원) → 자동 판정 원."""
    @classmethod
    def setUpClass(cls):
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 0,
                            {"recommended": True}, "default")
        cls.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        cls.wb = C.compose(cls.col, None, "p", "s")

    def test_judged_won(self):
        self.assertEqual((self.wb.unit_value, self.wb.unit_auto), (1, True))
        s = self.wb.get("설정")
        self.assertEqual(s.cells[(int(s.names["cfg_unit"][1:]), 3)].v, 1)

    def test_mfs_formulas_and_title(self):
        mfs = self.wb.get("mFS")
        title = mfs.cells[(1, C.RAW_LABEL_COL)]
        self.assertTrue(title.formula)
        self.assertIn("cfg_unit", title.v)
        data = [c.v for (r, cc), c in mfs.cells.items() if c.formula and c.fmt == "num" and not c.v.startswith("=COUNTIF(")]
        self.assertTrue(data)
        self.assertTrue(all(v.endswith("/cfg_unit") for v in data), data[:3])
        self.assertFalse(any("*1/" in v or "*1000" in v for v in data))   # 재무제표 raw 는 원 단위

    def test_checks_multiply_by_cfg_unit_no_tolerance(self):
        chk = self.wb.get("검증")
        checks = [c.v for (r, cc), c in chk.cells.items() if c.formula and c.fmt == "check"]
        self.assertTrue(checks)
        self.assertTrue(all("cfg_unit" in f for f in checks if "'mFS'!" in f), checks)
        self.assertFalse([f for f in checks if "ABS(" in f or "<=1" in f])
        sga = [f for f in checks if "'rN_판관비'!" in f and "'mFS'!" in f]
        self.assertTrue(sga)
        # 원 단위 BS 와 천원 단위 주석: 양쪽을 원으로 환산해 천원 반올림 후 비교(허용폭 없음)
        self.assertRegex(sga[0], r"^=ROUND\('rN_판관비'!E\d+\*1000/1000,0\)=ROUND\('mFS'!E\d+\*cfg_unit/1000,0\)$")

    def test_settings_validation_and_helper(self):
        s = self.wb.get("설정")
        r = int(s.names["cfg_unit"][1:])
        self.assertIn((f"C{r}", '"1,1000,1000000"'), s.validations)
        h = s.cells[(r, 4)]
        self.assertTrue(h.formula)
        self.assertEqual(h.v, '=IF(cfg_unit=1,"원",IF(cfg_unit=1000,"천원",IF(cfg_unit=1000000,"백만원","?")))')

    def test_openpyxl_writes_validation(self):
        import openpyxl
        from excel import write_openpyxl as W
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "t.xlsx")
            W.write(self.wb, p)
            ws = openpyxl.load_workbook(p)["설정"]
            dvs = ws.data_validations.dataValidation
            self.assertEqual(len(dvs), 1)
            self.assertEqual(dvs[0].formula1, '"1,1000,1000000"')
            self.assertEqual(str(dvs[0].sqref), "C10")

    def test_unit_cell_input_style(self):
        """cfg_unit 칸: 연노랑(FFFFCC) 단색 채움 + 사방 점선, 숫자 서식·목록 유효성 유지."""
        import openpyxl
        from excel import write_openpyxl as W
        s = self.wb.get("설정")
        r = int(s.names["cfg_unit"][1:])
        self.assertEqual(s.cells[(r, 3)].fmt, "input")
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "t.xlsx")
            W.write(self.wb, p)
            ws = openpyxl.load_workbook(p)["설정"]
            c = ws[f"C{r}"]
            self.assertEqual(c.fill.fill_type, "solid")
            self.assertEqual(c.fill.fgColor.rgb[-6:], "FFFFCC")
            for side in (c.border.left, c.border.right, c.border.top, c.border.bottom):
                self.assertEqual(side.style, "dotted")
            self.assertEqual(c.number_format, W.NUM_FMT)
            self.assertEqual(str(ws.data_validations.dataValidation[0].sqref), f"C{r}")
            self.assertIsNone(ws[f"C{r + 1}"].fill.fill_type)   # 다른 칸은 채움 없음

    def test_check_note_row_and_names(self):
        """'단위 판정' 바로 다음 '검증 안내'(기울임), 이후 행·정의된 이름이 한 칸씩 내려감."""
        import openpyxl
        from excel import write_openpyxl as W
        s = self.wb.get("설정")
        r = _settings_row(s, "단위 판정")
        self.assertEqual(s.cells[(r + 1, 2)].v, "검증 안내")
        note = s.cells[(r + 1, 3)]
        self.assertEqual(note.fmt, "italic")
        self.assertTrue(note.v.startswith("주석 표는 항목별 천원 반올림으로 공시되어"))
        self.assertTrue(note.v.endswith("허용폭 없이 그대로 표시합니다."))
        self.assertEqual(s.names, {"cfg_unit": "C10", "cfg_python": "C14", "cfg_script": "C15",
                                   "cfg_status": "C17", "cfg_json": "C18"})
        self.assertEqual(_settings_row(s, "마지막 갱신"), 16)
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "t.xlsx")
            W.write(self.wb, p)
            x = openpyxl.load_workbook(p)
            ws = x["설정"]
            self.assertTrue(ws[f"C{r + 1}"].font.italic)
            self.assertEqual(x.defined_names["cfg_status"].attr_text, "'설정'!$C$17")
            self.assertEqual(x.defined_names["cfg_json"].attr_text, "'설정'!$C$18")
            self.assertEqual(ws["B17"].value, "상태")
            self.assertEqual(ws["B18"].value, "설정 JSON")
            from excel import refresh as R
            rows = [list(row) for row in ws.iter_rows(values_only=True)]
            self.assertEqual(R._setting(rows, "설정 JSON"), ws["C18"].value)
            self.assertEqual(R._setting(rows, "상태"), "생성 완료")


class TestComposeExplicit(unittest.TestCase):
    def test_explicit_choice_kept(self):
        wb = C.compose(_hand_col(1000, unit_display=1000000), None, "p", "s")
        s = wb.get("설정")
        self.assertEqual(s.cells[(int(s.names["cfg_unit"][1:]), 3)].v, 1000000)
        self.assertEqual((wb.unit_value, wb.unit_auto), (1000000, False))
        self.assertIn("직접 선택", s.cells[(_settings_row(s, "단위 판정"), 3)].v)
        msgs = [c.v for c in wb.get("로그").cells.values()]
        self.assertFalse(any(str(m).startswith("표시 단위 자동 판정") for m in msgs))


class TestRefreshUnitResolution(unittest.TestCase):
    def test_resolve(self):
        auto = {"unit_display": 0, "unit_judged": 1000}
        self.assertEqual(R.resolve_unit(auto, 1000.0), 0)            # 자동 + 셀이 판정값 그대로 → 다시 판정
        self.assertEqual(R.resolve_unit(auto, None), 0)
        self.assertEqual(R.resolve_unit(auto, 1000000.0), 1000000)   # 엑셀에서 바꿈 → 그 값(직접 선택)
        self.assertEqual(R.resolve_unit({"unit_display": 1000}, 1.0), 1)   # 직접 선택 → 셀 값
        self.assertEqual(R.resolve_unit({"unit_display": 1000}, None), 1000)
        self.assertEqual(R.resolve_unit({"unit_display": 1000}, "x"), 1000)
        self.assertEqual(R.resolve_unit({"unit_display": 0}, 1.0), 0)       # 판정값 기록 없는 자동 → 자동

    def test_settings_updates_rewrites_unit_when_auto(self):
        old = P.BuildConfig("OB", "t", date(2026, 10, 26), 2023, 2025, False, ["OFS"], 0, None, "x")
        d = old.to_dict()
        d["unit_judged"] = 1
        rows = [[None, "mFS 표시 단위", 1], [None, "단위 판정", "자동 판정: 원"],
                [None, "설정 JSON", json.dumps(d, ensure_ascii=False)]]
        got = R.settings_updates(rows, old, {"value": 1000, "auto": True, "text": "자동 판정: 천원 (근거: x)"})
        self.assertIn((1, 3, 1000), got)
        self.assertIn((2, 3, "자동 판정: 천원 (근거: x)"), got)
        js = json.loads(next(v for r, c, v in got if r == 3))
        self.assertEqual((js["unit_display"], js["unit_judged"]), (0, 1000))
        # 직접 선택이면 셀·판정값은 그대로(판정 문구만 갱신)
        got = R.settings_updates(rows, old, {"value": 1000000, "auto": False, "text": "직접 선택: 백만원"})
        self.assertNotIn(1, [r for r, c, v in got])
        js = json.loads(next(v for r, c, v in got if r == 3))
        self.assertEqual((js["unit_display"], js["unit_judged"]), (0, 1))


if __name__ == "__main__":
    unittest.main()
