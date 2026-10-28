import re, unittest
from datetime import date
from core import pipeline as P, mapping as M
from excel import compose as C, grid as G
from tests.test_pipeline import FakeClient


class TestCompose(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, {"recommended": True}, "default")
        cls.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        cls.wb = C.compose(cls.col, None, r"C:\py\python.exe", r"C:\app\excel\refresh.py")

    def test_sheet_order(self):
        names = [s.name for s in self.wb.sheets]
        self.assertEqual(names[:5], ["설정", "목차", "mFS", "검증", "RAW >>"])
        self.assertIn("rFS_별도", names)
        self.assertIn("rN_판관비", names); self.assertIn("rN_유형자산", names)
        self.assertEqual(names[-3:], ["데이터", "매핑표", "로그"])
        self.assertLess(names.index("rN_기타"), names.index("데이터"))
        self.assertTrue(all(C.is_data_sheet(n) for n in names if n != "설정"))
        self.assertFalse(C.is_data_sheet("설정")); self.assertFalse(C.is_data_sheet("모델"))

    def test_rfs_blocks_layout(self):
        s = self.wb.get("rFS_별도")
        vals = s.values_matrix()
        titles = [r[C.RAW_LABEL_COL - 1] for r in vals if r[C.RAW_LABEL_COL - 1] in ("재무상태표", "포괄손익계산서", "현금흐름표", "자본변동표")]
        self.assertEqual(titles[0], "재무상태표")
        r_title = next(i for i, r in enumerate(vals) if r[C.RAW_LABEL_COL - 1] == "재무상태표")
        self.assertEqual(vals[r_title][2], "§")                       # 블록 경계 표식(C열)
        self.assertEqual(vals[r_title + 1][C.RAW_LABEL_COL - 1], "(단위 : 원)")
        self.assertEqual(vals[r_title + 2][C.RAW_FIRST_VALUE_COL - 1], "제 27 기")
        self.assertEqual(vals[r_title + 3][C.RAW_FIRST_VALUE_COL - 1], date(2024, 12, 31))
        self.assertEqual(vals[r_title + 3][C.RAW_FIRST_VALUE_COL], date(2025, 12, 31))
        cash = next(r for r in vals if r[C.RAW_LABEL_COL - 1] == "1. 현금및현금성자산")
        self.assertEqual(cash[C.RAW_FIRST_VALUE_COL], 171311183566)
        self.assertEqual(cash[C.RAW_FIRST_VALUE_COL - 1], 206064972985)
        # 블록 사이 빈 행 RAW_GAP_ROWS 개
        r_is = next(i for i, r in enumerate(vals) if r[C.RAW_LABEL_COL - 1] == "포괄손익계산서")
        self.assertTrue(all(all(v is None for v in vals[i]) for i in range(r_is - C.RAW_GAP_ROWS, r_is)))
        self.assertTrue(any(v is not None for v in vals[r_is - C.RAW_GAP_ROWS - 1]))

    def test_mfs_formulas_reference_raw(self):
        s = self.wb.get("mFS")
        cells = [c for c in s.cells.values() if c.formula]
        self.assertTrue(cells)
        self.assertTrue(any(c.v.startswith("='rFS_별도'!E") or c.v.startswith("=rFS_별도!E") for c in cells))
        labels = [c.v for (r, cc), c in s.cells.items() if cc == C.RAW_LABEL_COL and c.fmt in (None, "text", "bold")]
        self.assertIn("자산총계", labels); self.assertIn("매출액", labels)
        self.assertFalse([c.v for c in cells if str(c.v).startswith("=ROUND(")])      # 검증식은 검증 시트로
        chk = [c.v for c in self.wb.get("검증").cells.values() if c.formula and str(c.v).startswith("=ROUND(")]
        self.assertTrue(chk)

    def test_mfs_refs_point_at_raw_rows(self):
        """mFS 수식이 가리키는 raw 셀이 해당 raw 행이다."""
        raw = self.wb.get("rFS_별도"); mfs = self.wb.get("mFS")
        r_cash = next(r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL and cell.v == "현금및현금성자산")
        f = mfs.cells[(r_cash, C.RAW_FIRST_VALUE_COL)].v
        rr = int(re.search(r"!E(\d+)", f).group(1))
        self.assertEqual(raw.cells[(rr, C.RAW_LABEL_COL)].v, "1. 현금및현금성자산")
        self.assertEqual(f, f"='rFS_별도'!E{rr}/cfg_unit")

    def test_mfs_unmapped_rows_kept(self):
        mfs = self.wb.get("mFS")
        self.assertTrue([1 for cell in mfs.cells.values() if cell.v == "미매핑"])
        labels = {cell.v for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL}
        self.assertIn("XII. 기본주당이익", labels)     # 매핑 안 된 raw 행도 버리지 않는다

    def test_checks_on_check_sheet(self):
        """검증식은 mFS 가 아니라 검증 시트에(자세한 규칙은 test_checks). mFS 에는 요약 행만."""
        vals = self.wb.get("검증").values_matrix()
        nos = {r[1]: r for r in vals if r[1]}
        self.assertTrue(str(nos["A1"][C.RAW_FIRST_VALUE_COL - 1]).startswith("=ROUND('mFS'!E"))
        self.assertIn("'rN_판관비'!E", str(nos["B1"][C.RAW_FIRST_VALUE_COL - 1]))
        self.assertIn("'rN_금융손익'!E", str(nos["B5"][C.RAW_FIRST_VALUE_COL - 1]))
        self.assertIn("/1000", str(nos["B5"][C.RAW_FIRST_VALUE_COL - 1]))
        mvals = self.wb.get("mFS").values_matrix()
        self.assertTrue(any(r[C.RAW_LABEL_COL - 1] == "검증 TRUE 수" for r in mvals))
        self.assertFalse(any(r[C.RAW_LABEL_COL - 1] == "자산 = 부채 + 자본" for r in mvals))

    def test_note_sheet_blocks(self):
        vals = self.wb.get("rN_판관비").values_matrix()
        self.assertTrue(vals[0][C.RAW_LABEL_COL - 1].startswith("주석 27. 판매비, 관리비와 물류비 | "))
        self.assertEqual(vals[1][C.RAW_LABEL_COL - 1], "(단위 : 천원)")
        salary = next(r for r in vals if r[C.RAW_LABEL_COL - 1] == "급여")
        self.assertEqual(salary[C.RAW_FIRST_VALUE_COL - 1:C.RAW_FIRST_VALUE_COL + 1], [130368479, 130500219])

    def test_section_titles_unique_per_sheet(self):
        for s in self.wb.sheets:
            if s.name.startswith("rN_") or s.name.startswith("rFS_"):
                titles = [s.cells[(r, C.RAW_LABEL_COL)].v for (r, c) in s.cells if c == 3]
                self.assertTrue(titles, s.name)
                self.assertEqual(len(titles), len(set(titles)), s.name)

    def test_other_tables_in_rn_etc(self):
        vals = self.wb.get("rN_기타").values_matrix()
        tags = [r[C.RAW_LABEL_COL - 1] for r in vals if r[2] == "§"]
        self.assertTrue(any(t.startswith("주석 15. 종속기업투자주식") for t in tags))
        # 기간 표시 없는 원표는 rN_종속관계 에 섞이지 않는다
        sub = self.wb.get("rN_종속관계").values_matrix()
        self.assertEqual({r[C.RAW_LABEL_COL - 1] for r in sub if r[2] == "§"} & set(tags), set())

    def test_toc_sheet_has_include_flags(self):
        vals = self.wb.get("목차").values_matrix()
        self.assertEqual(vals[0][:5], ["번호", "제목", "표준 주제", "포함", "추천"])
        row27 = next(r for r in vals if r[0] == "27")
        self.assertEqual((row27[2], row27[3], row27[4]), ("판관비", True, "★"))

    def test_settings_names(self):
        s = self.wb.get("설정")
        self.assertEqual(set(s.names), {"cfg_python", "cfg_script", "cfg_status", "cfg_json", "cfg_unit"})
        self.assertEqual(s.cells[(int(s.names["cfg_python"][1:]), 3)].v, r"C:\py\python.exe")
        self.assertEqual(s.cells[(int(s.names["cfg_unit"][1:]), 3)].v, 1)
        self.assertFalse(s.is_data)

    def test_data_sheet_long_format(self):
        vals = self.wb.get("데이터").values_matrix()
        self.assertEqual(vals[0], ["시트", "블록", "열라벨", "행라벨", "기간", "값", "단위", "출처 rcept_no", "출처", "표준계정"])
        self.assertGreater(len(vals), 100)

    def test_existing_order_applied(self):
        wb2 = C.compose(self.col, {"rFS_별도": {"재무상태표": ["자 산", "없어진계정", "1. 현금및현금성자산"]}}, "p", "s")
        vals = wb2.get("rFS_별도").values_matrix()
        r_title = next(i for i, r in enumerate(vals) if r[C.RAW_LABEL_COL - 1] == "재무상태표")
        self.assertEqual(vals[r_title + 4][C.RAW_LABEL_COL - 1], "자 산")
        self.assertEqual(vals[r_title + 5][C.RAW_LABEL_COL - 1], "없어진계정")
        self.assertEqual(vals[r_title + 6][C.RAW_LABEL_COL - 1], "1. 현금및현금성자산")

    def test_existing_order_note_block_uses_section_title(self):
        sec = self.wb.get("rN_판관비").cells[(1, C.RAW_LABEL_COL)].v
        wb2 = C.compose(self.col, {"rN_판관비": {sec: ["퇴직급여", "급여"]}}, "p", "s")
        vals = wb2.get("rN_판관비").values_matrix()
        self.assertEqual([vals[4][C.RAW_LABEL_COL - 1], vals[5][C.RAW_LABEL_COL - 1]], ["퇴직급여", "급여"])


class TestGrid(unittest.TestCase):
    def test_sheet_basics(self):
        s = G.Sheet("x")
        s.set(2, 3, "a"); s.set(1, 1, 5, "num")
        self.assertEqual((s.max_row(), s.max_col()), (2, 3))
        self.assertEqual(s.values_matrix(), [[5, None, None], [None, None, "a"]])
        self.assertEqual(G.Sheet("e").values_matrix(), [])
        self.assertIs(G.WorkbookModel([s]).get("x"), s)


if __name__ == "__main__":
    unittest.main()



class TestSameLabelDifferentAccount(unittest.TestCase):
    def test_lease_liabilities_split_by_account_id(self):
        from types import SimpleNamespace
        from core.model import FsRecord, Period
        pers = [Period("FY2024", date(2024, 12, 31), "annual", 26), Period("FY2025", date(2025, 12, 31), "annual", 27)]
        cfg = P.BuildConfig("X", "X", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, None, "default")
        recs = []
        for pk, (a, b) in (("FY2024", (10, 20)), ("FY2025", (11, 21))):
            recs.append(FsRecord(pk, "OFS", "BS", 1, "리스부채", "리스부채", a, 1, None,
                                 "ifrs-full_CurrentLeaseLiabilities", "r", "XML", "cur"))
            recs.append(FsRecord(pk, "OFS", "BS", 2, "리스부채", "리스부채", b, 1, None,
                                 "ifrs-full_NoncurrentLeaseLiabilities", "r", "XML", "cur"))
        col = P.Collected(cfg, SimpleNamespace(periods=pers, reports={}), ["FY2024", "FY2025"],
                          fs_records=recs, mapping=M.load())
        wb = C.compose(col, None, "p", "s")
        mfs = wb.get("mFS")
        rows = {cell.v: r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL}
        self.assertIn("유동리스부채", rows); self.assertIn("비유동리스부채", rows)
        fc = mfs.cells[(rows["유동리스부채"], C.RAW_FIRST_VALUE_COL)].v
        fn = mfs.cells[(rows["비유동리스부채"], C.RAW_FIRST_VALUE_COL)].v
        self.assertNotEqual(fc, fn)
        self.assertNotIn("+", fc); self.assertNotIn("+", fn)
