"""조서 양식 규칙: 머리글 2행 실선·합계 아래 실선·소제목 행 라벨만·비율 숫자화·단위 미표기·COM 동일 적용."""
import os, tempfile, unittest, openpyxl
from datetime import date
from core import pipeline as P, mapping as M, normalize as N
from excel import compose as C, write_openpyxl as W, write_com as WC
from tests.test_pipeline import FakeClient

D, E = C.RAW_LABEL_COL, C.RAW_FIRST_VALUE_COL


def _find(ws, col, text, start=1):
    for r in range(start, ws.max_row + 1):
        if ws.cell(r, col).value == text:
            return r
    raise AssertionError(f"{text!r} 없음")


def _last_col(ws, r):
    return max(c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value not in (None, ""))


class TestFormatRules(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, {"recommended": True}, "default")
        cls.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        cls.model = C.compose(cls.col, None, r"C:\py\python.exe", r"C:\app\excel\refresh.py")
        cls.tmp = tempfile.TemporaryDirectory()
        cls.x = openpyxl.load_workbook(W.write(cls.model, os.path.join(cls.tmp.name, "f.xlsx")))

    @classmethod
    def tearDownClass(cls):
        cls.x.close()
        cls.tmp.cleanup()

    # 1. 머리글 2행: 첫 행 위 실선, 둘째 행 아래 실선, D열부터 마지막 기간 열까지
    def test_raw_header_two_rows(self):
        ws = self.x["rFS_별도"]
        r = _find(ws, D, "재무상태표")
        h1, h2 = r + 2, r + 3
        last = _last_col(ws, h1)
        self.assertGreaterEqual(last, E + 1)
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(h1, c).border.top.style, "thin", (h1, c))
            self.assertIsNone(ws.cell(h1, c).border.bottom.style, (h1, c))
            self.assertEqual(ws.cell(h2, c).border.bottom.style, "thin", (h2, c))
            self.assertIsNone(ws.cell(h2, c).border.top.style, (h2, c))
            self.assertEqual(ws.cell(h1, c).alignment.horizontal, "center")
        self.assertEqual(ws.cell(h2, E).number_format, "yyyy-mm-dd")
        self.assertEqual(ws.cell(h2, E).alignment.horizontal, "center")

    def test_note_block_header_two_rows(self):
        ws = self.x["rN_판관비"]
        r = next(i for i in range(1, ws.max_row + 1) if ws.cell(i, 3).value == "§")
        last = _last_col(ws, r + 2)
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(r + 2, c).border.top.style, "thin")
            self.assertEqual(ws.cell(r + 3, c).border.bottom.style, "thin")

    # 2. 합계·총계 행: 굵게 + 아래 실선(D..마지막 기간 열)
    def test_total_rows_bold_bottom_line(self):
        ws = self.x["rFS_별도"]
        r0 = _find(ws, D, "재무상태표")
        last = _last_col(ws, r0 + 2)
        r = next(i for i in range(r0, ws.max_row + 1) if N.norm_label(ws.cell(i, D).value or "") == "자산총계")
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(r, c).border.bottom.style, "thin", c)
            self.assertTrue(ws.cell(r, c).font.bold, c)
        self.assertEqual(ws.cell(r, E).number_format, W.NUM_FMT)
        cash = next(i for i in range(r0, ws.max_row + 1) if ws.cell(i, D).value == "1. 현금및현금성자산")
        self.assertIsNone(ws.cell(cash, D).border.bottom.style)
        self.assertFalse(ws.cell(cash, D).font.bold)

    def test_is_total(self):
        for s in ("합계", "총계", "소계", "계", "자산총계", "부채와자본총계", "판매비와관리비 합계"):
            self.assertTrue(C.is_total(N.norm_label(s)), s)
        for s in ("당기순이익", "유동자산", "회계", "설계"):
            self.assertFalse(C.is_total(N.norm_label(s)), s)

    # 3. 소제목 행(값이 전부 라벨과 같은 글자) → 라벨만
    def test_subheading_row_label_only_data_level(self):
        rs = [r for r in self.col.note_records if r.note_title == "리스" and r.row_label == "사용권자산" and r.raw_cell is None]
        self.assertTrue(rs)
        self.assertTrue(all(r.value is None for r in rs), [r.value for r in rs])

    def test_subheading_row_label_only_sheet(self):
        found = False
        for ws in self.x.worksheets:
            if not ws.title.startswith("rN_") or ws.title == C.ETC_SHEET:
                continue
            for r in range(1, ws.max_row + 1):
                lab = ws.cell(r, D).value
                if not isinstance(lab, str) or ws.cell(r, 3).value == "§":
                    continue
                for c in range(E, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and v == lab:
                        self.fail(f"{ws.title}!{r}: 소제목 값 {v!r}")
                if lab == "사용권자산":
                    found = True
        self.assertTrue(found)

    # 4. 비율 → 숫자 + 0.0%
    def test_parse_percent(self):
        self.assertAlmostEqual(N.parse_percent("83.83%"), 0.8383)
        self.assertIsInstance(N.parse_percent("23.33 %"), N.Percent)
        self.assertAlmostEqual(N.parse_percent("(1.5%)"), -0.015)
        self.assertAlmostEqual(N.parse_percent("-2%"), -0.02)
        self.assertIsNone(N.parse_percent("지분율(%)"))
        self.assertIsNone(N.parse_percent("12월"))
        self.assertIsNone(N.parse_percent(None))

    def test_percent_values_numeric_with_pct_format(self):
        rs = [r for r in self.col.note_records if r.row_label == "유효세율"]
        self.assertTrue(rs)
        self.assertTrue(all(isinstance(r.value, N.Percent) for r in rs), [r.value for r in rs])
        ws, r = None, None
        for w in self.x.worksheets:
            if not w.title.startswith("rN_"):
                continue
            for i in range(1, w.max_row + 1):
                if w.cell(i, D).value == "유효세율":
                    ws, r = w, i
                    break
            if r:
                break
        cells = [ws.cell(r, c) for c in range(E, E + 2)]
        self.assertTrue(any(isinstance(c.value, float) and abs(c.value - 0.2859) < 1e-9 for c in cells),
                        [c.value for c in cells])
        self.assertTrue(all(c.number_format == "0.0%" for c in cells if isinstance(c.value, float)))
        dws = self.x["데이터"]
        pct = [c for c in dws["F"] if isinstance(c.value, float) and c.number_format == "0.0%"]
        self.assertTrue(pct)

    def test_raw_cells_percent(self):
        class G_:
            header = [["구분", "지분율"]]
            body = [["A사", "83.83%"], ["B사", "12월"]]
            ncols = 2

        class NT:
            grid = G_()
        out = P._raw_cells(NT())
        self.assertAlmostEqual(out[1][2], 0.8383)
        self.assertIsInstance(out[1][2], N.Percent)
        self.assertEqual(out[3][2], "12월")

    def test_raw_cells_subheading_blank(self):
        class G_:
            header = [["구분", "당기", "전기"]]
            body = [["(1) 판매비:", "(1) 판매비:", "(1) 판매비:"], ["급여", "10", "9"]]
            ncols = 3

        class NT:
            grid = G_()
        out = P._raw_cells(NT())
        self.assertEqual(out[0][2], "(1) 판매비:")      # 라벨 열은 그대로
        self.assertEqual(out[1][2], "")
        self.assertEqual(out[2][2], "")
        self.assertEqual(out[4][2], "10")

    # 4. 원표: 단위 미표기·머리글 실선
    def test_raw_table_unit_unknown_and_header_border(self):
        ws = self.x[C.ETC_SHEET]
        raws = [i for i in range(1, ws.max_row + 1) if ws.cell(i, 3).value == "§"
                and isinstance(ws.cell(i + 1, E).value, str) and "주석" in ws.cell(i + 1, E).value]
        self.assertTrue(raws)
        units = {ws.cell(i + 1, D).value for i in raws}
        self.assertIn("(단위 : 원문 미표기)", units)
        r = raws[0]
        last = _last_col(ws, r + 2)
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(r + 2, c).border.top.style, "thin")
            self.assertEqual(ws.cell(r + 2, c).border.bottom.style, "thin")
            self.assertEqual(ws.cell(r + 2, c).alignment.horizontal, "center")

    def test_raw_table_unit_known_flag(self):
        recs = [r for r in self.col.note_records if r.raw_cell is not None]
        self.assertTrue(recs)
        for r in recs:
            self.assertEqual(r.unit_known, N.parse_unit(r.caption) is not None)

    # 5. 날짜 서식: 로그
    def test_log_dates(self):
        ws = self.x["로그"]
        dates = [c for row in ws.iter_rows() for c in row if hasattr(c.value, "year")]
        self.assertTrue(dates)
        self.assertTrue(all(c.number_format == "yyyy-mm-dd" for c in dates))

    # 6. mFS: 제목 12 굵게, 구역 머리글 실선, 검증 구역 머리글, 합계류 굵게+아래 실선
    def test_mfs(self):
        ws = self.x["mFS"]
        self.assertEqual(ws.cell(1, D).font.size, 12); self.assertTrue(ws.cell(1, D).font.bold)
        r = _find(ws, D, "별도 재무상태표")
        self.assertEqual(ws.cell(r, 3).value, "§"); self.assertEqual(ws.cell(r, D).font.size, 11)
        last = _last_col(ws, r + 1)
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(r + 1, c).border.top.style, "thin")
            self.assertEqual(ws.cell(r + 1, c).border.bottom.style, "thin")
        t = next(i for i in range(r, ws.max_row + 1) if ws.cell(i, D).value == "자산총계")
        self.assertTrue(ws.cell(t, D).font.bold)
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(t, c).border.bottom.style, "thin")
            self.assertTrue(ws.cell(t, c).font.bold)
        self.assertEqual(ws.cell(t, E).number_format, W.NUM_FMT)
        k = _find(ws, D, "별도 검증 요약")
        self.assertEqual(ws.cell(k + 1, E).value, ws.cell(r + 1, E).value)      # 기간 머리글
        for c in range(D, last + 1):
            self.assertEqual(ws.cell(k + 1, c).border.top.style, "thin")
            self.assertEqual(ws.cell(k + 1, c).border.bottom.style, "thin")
        self.assertEqual(ws.cell(k + 2, D).value, "검증 TRUE 수")
        self.assertTrue(str(ws.cell(k + 2, E).value).startswith("=COUNTIF('검증'!E"))
        self.assertEqual(ws.cell(k + 2, E).number_format, W.NUM_FMT)

    # 6b. 검증 시트: 머리글 2행 실선, 구역 제목 11 굵게, 검증 셀 가운데 General, 차이 열 숫자 서식, 해당 없음 기울임
    def test_check_sheet(self):
        ws = self.x["검증"]
        self.assertEqual(ws.sheet_names if False else self.x.sheetnames[:5], ["설정", "목차", "mFS", "검증", "RAW >>"])
        last = _last_col(ws, 2)
        for c in range(1, last + 1):
            self.assertEqual(ws.cell(2, c).border.top.style, "thin", c)
            self.assertEqual(ws.cell(3, c).border.bottom.style, "thin", c)
        self.assertEqual(ws.cell(2, E).value, "FY2024")
        r = _find(ws, D, "A. 재무제표 내부 정합 — 별도")
        self.assertEqual(ws.cell(r, D).font.size, 11); self.assertTrue(ws.cell(r, D).font.bold)
        a1 = next(i for i in range(1, ws.max_row + 1) if ws.cell(i, 2).value == "A1")
        self.assertTrue(str(ws.cell(a1, E).value).startswith("=ROUND("))
        self.assertEqual(ws.cell(a1, E).alignment.horizontal, "center")
        self.assertEqual(ws.cell(a1, E).number_format, "General")
        hdr = {ws.cell(2, c).value: c for c in range(1, last + 1)}
        self.assertEqual(ws.cell(a1, hdr["차이 FY2024"]).number_format, W.NUM_FMT)
        na = next(i for i in range(1, ws.max_row + 1) if ws.cell(i, E).value == "해당 없음")
        self.assertTrue(ws.cell(na, E).font.italic)
        self.assertEqual(ws.cell(a1, D).font.name, "맑은 고딕"); self.assertEqual(ws.cell(a1, D).font.size, 10)

    # 7. COM 경로: 테두리·비율 묶음이 openpyxl 과 같은 셀을 가리킨다
    def test_com_border_groups_match(self):
        sh = self.model.get("rFS_별도")
        groups = WC.border_groups(sh)
        r = next(rr for (rr, c), cell in sh.cells.items() if c == D and cell.v == "재무상태표")
        self.assertIn((r + 2, D), groups["top"])
        self.assertIn((r + 3, D), groups["bottom"])
        self.assertNotIn((r + 2, D), groups.get("both", []))
        self.assertNotIn((r + 2, D), groups["bottom"])
        tot = next(rr for (rr, c), cell in sh.cells.items() if c == D and N.norm_label(str(cell.v)) == "자산총계")
        self.assertIn((tot, E), groups["bottom"])
        self.assertEqual(WC.COM_DATE_FMT, 'yyyy"-"mm"-"dd')

    def test_percent_cells_grouped_for_com(self):
        sh = next(s for s in self.model.sheets if s.name.startswith("rN_")
                  and any(isinstance(c.v, N.Percent) for c in s.cells.values()))
        self.assertTrue(WC.style_groups(sh).get("pct"))


if __name__ == "__main__":
    unittest.main()
