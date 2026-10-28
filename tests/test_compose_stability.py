"""업데이트 때 행·열 안정성(설계 §5): 기존 블록 시작 행·열 순서·mFS 행 순서 유지, 넘침 경고."""
import copy, dataclasses, unittest
from datetime import date
from core import pipeline as P, mapping as M, normalize as N
from excel import compose as C
from tests.test_pipeline import FakeClient


def _starts(wb, sheet):
    return {t: b["start"] for t, b in C.existing_from_model(wb)[sheet]["blocks"].items()}


def _labels_at(wb, sheet):
    s = wb.get(sheet)
    return {r: c.v for (r, cc), c in s.cells.items() if cc == C.RAW_LABEL_COL}


def _with_extra_bs(col, n):
    """재무상태표 FY2025 에 새 계정 n 개를 더한 수집 결과(블록이 n 행 커짐)."""
    c2 = copy.copy(col)
    base = next(r for r in col.fs_records if r.sj_div == "BS" and r.period_key == "FY2025")
    extra = [dataclasses.replace(base, label=f"신규계정{i}", label_norm=N.norm_label(f"신규계정{i}"),
                                 ord=10 ** 6 + i, amount=1000 + i) for i in range(n)]
    c2.fs_records = list(col.fs_records) + extra
    return c2


class TestStability(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1,
                            {"recommended": True}, "default")
        cls.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        cls.wb1 = C.compose(cls.col, None, "p", "s")
        cls.ex1 = C.existing_from_model(cls.wb1)

    def test_existing_from_model_shape(self):
        ex = self.ex1["rFS_별도"]
        self.assertEqual(ex["periods"], ["FY2024", "FY2025"])
        self.assertEqual(ex["blocks"]["재무상태표"]["start"], 1)
        self.assertEqual(ex["blocks"]["재무상태표"]["labels"][:2], ["자 산", "I. 유동자산"])
        self.assertIn("rN_판관비", self.ex1)
        self.assertEqual(self.ex1["rN_판관비"]["periods"], ["FY2024", "FY2025"])
        mfs = self.ex1["mFS"]
        self.assertEqual(mfs["periods"], ["FY2024", "FY2025"])
        self.assertIn("별도 재무상태표", mfs["blocks"])
        self.assertIn("별도 검증 요약", mfs["blocks"])
        self.assertEqual(self.ex1["rN_기타"]["periods"], [])        # 원표는 기간 열이 없다

    def test_same_input_same_layout(self):
        wb2 = C.compose(self.col, self.ex1, "p", "s")
        for name in ("rFS_별도", "rN_유형자산", "rN_기타", "mFS"):
            a = {k: c.v for k, c in self.wb1.get(name).cells.items()}
            b = {k: c.v for k, c in wb2.get(name).cells.items()}
            self.assertEqual(a, b, name)
        self.assertEqual(wb2.shift_warnings, [])

    def test_a_growth_consumes_spare_row(self):
        wb2 = C.compose(_with_extra_bs(self.col, 1), self.ex1, "p", "s")
        before, after = _starts(self.wb1, "rFS_별도"), _starts(wb2, "rFS_별도")
        self.assertEqual(before, after)                          # 다음 블록(포괄손익계산서 등) 시작 행 그대로
        bs = self.ex1["rFS_별도"]["blocks"]["재무상태표"]
        old_end = bs["start"] + 4 + len(bs["labels"])            # 예전 블록 끝 다음 행 = 첫 여유 행
        self.assertEqual(_labels_at(wb2, "rFS_별도")[old_end], "신규계정0")
        self.assertLess(old_end, after["포괄손익계산서"])
        self.assertEqual(wb2.shift_warnings, [])

    def test_b_new_period_appended_right(self):
        ex = copy.deepcopy(self.ex1)
        ex["rFS_별도"]["periods"] = ["FY2025"]                    # 예전 파일엔 FY2025 열만 있었다
        wb2 = C.compose(self.col, ex, "p", "s")
        vals = wb2.get("rFS_별도").values_matrix()
        e, f = C.RAW_FIRST_VALUE_COL - 1, C.RAW_FIRST_VALUE_COL
        self.assertEqual((vals[3][e], vals[3][f]), (date(2025, 12, 31), date(2024, 12, 31)))
        cash = next(r for r in vals if r[C.RAW_LABEL_COL - 1] == "1. 현금및현금성자산")
        self.assertEqual((cash[e], cash[f]), (171311183566, 206064972985))
        # mFS 수식도 같은 열을 가리킨다(FY2025 = E열)
        mfs = wb2.get("mFS")
        r_cash = next(r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL and cell.v == "현금및현금성자산")
        rr = int(mfs.cells[(r_cash, C.RAW_FIRST_VALUE_COL + 1)].v.split("!E")[1].split("/")[0])
        self.assertEqual(vals[rr - 1][e], 171311183566)

    def test_b_missing_period_keeps_column(self):
        c2 = copy.copy(self.col)
        c2.periods = ["FY2025"]
        c2.fs_records = [r for r in self.col.fs_records if r.period_key == "FY2025"]
        wb2 = C.compose(c2, self.ex1, "p", "s")
        vals = wb2.get("rFS_별도").values_matrix()
        e, f = C.RAW_FIRST_VALUE_COL - 1, C.RAW_FIRST_VALUE_COL
        self.assertEqual((vals[3][e], vals[3][f]), (date(2024, 12, 31), date(2025, 12, 31)))
        cash = next(r for r in vals if r[C.RAW_LABEL_COL - 1] == "1. 현금및현금성자산")
        self.assertEqual((cash[e], cash[f]), (None, 171311183566))
        mvals = wb2.get("mFS").values_matrix()
        hdr = next(r for r in mvals if r[C.RAW_FIRST_VALUE_COL - 1] == "FY2024")
        self.assertEqual(hdr[C.RAW_FIRST_VALUE_COL], "FY2025")

    def test_c_overflow_shifts_and_warns(self):
        wb2 = C.compose(_with_extra_bs(self.col, 6), self.ex1, "p", "s")
        before, after = _starts(self.wb1, "rFS_별도"), _starts(wb2, "rFS_별도")
        self.assertEqual(after["재무상태표"], before["재무상태표"])
        self.assertEqual(after["포괄손익계산서"] - before["포괄손익계산서"], 6)
        msg = "rFS_별도 '포괄손익계산서': 블록이 6행 밀림 — 모델 참조 확인 필요"
        self.assertIn(msg, wb2.shift_warnings)
        log = [c.v for c in wb2.get("로그").cells.values()]
        self.assertIn(msg, log)

    def test_c_five_extra_rows_still_fit(self):
        wb2 = C.compose(_with_extra_bs(self.col, C.RAW_GAP_ROWS), self.ex1, "p", "s")
        self.assertEqual(_starts(self.wb1, "rFS_별도"), _starts(wb2, "rFS_별도"))
        self.assertFalse([w for w in wb2.shift_warnings if w.startswith("rFS_별도")])

    def test_new_block_goes_after_last_existing(self):
        ex = copy.deepcopy(self.ex1)
        del ex["rFS_별도"]["blocks"]["재무상태표"]                # 예전 파일엔 재무상태표가 없었다
        wb2 = C.compose(self.col, ex, "p", "s")
        before, after = _starts(self.wb1, "rFS_별도"), _starts(wb2, "rFS_별도")
        others = [t for t in before if t != "재무상태표"]
        self.assertEqual({t: after[t] for t in others}, {t: before[t] for t in others})
        last = max(others, key=lambda t: before[t])
        self.assertGreater(after["재무상태표"], after[last])

    def test_d_mfs_existing_row_order(self):
        ex = copy.deepcopy(self.ex1)
        sec = ex["mFS"]["blocks"]["별도 재무상태표"]
        sec["labels"] = list(reversed(sec["labels"])) + ["없어진표준계정"]
        wb2 = C.compose(self.col, ex, "p", "s")
        mfs = wb2.get("mFS")
        new = C.existing_from_model(wb2)["mFS"]["blocks"]["별도 재무상태표"]
        self.assertEqual(new["labels"][:len(sec["labels"])], sec["labels"])
        self.assertEqual(new["start"], self.ex1["mFS"]["blocks"]["별도 재무상태표"]["start"])
        r_gone = next(r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL and cell.v == "없어진표준계정")
        self.assertIn("공시에서 사라짐", [cell.v for (r, c), cell in mfs.cells.items() if r == r_gone])
        self.assertFalse([cell for (r, c), cell in mfs.cells.items() if r == r_gone and cell.formula])
        raw = wb2.get("rFS_별도")                                 # 순서가 바뀌어도 수식은 제 raw 행을 가리킨다
        r_cash = next(r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL and cell.v == "현금및현금성자산")
        rr = int(mfs.cells[(r_cash, C.RAW_FIRST_VALUE_COL)].v.split("!E")[1].split("/")[0])
        self.assertEqual(raw.cells[(rr, C.RAW_LABEL_COL)].v, "1. 현금및현금성자산")
        # 검증 시트의 검증식도 옮겨진 mFS 행을 참조
        chk = wb2.get("검증")
        r_chk = next(r for (r, c), cell in chk.cells.items() if c == 2 and cell.v == "A1")
        r_asset = next(r for (r, c), cell in mfs.cells.items() if c == C.RAW_LABEL_COL and cell.v == "자산총계")
        self.assertIn(f"'mFS'!E{r_asset}*", chk.cells[(r_chk, C.RAW_FIRST_VALUE_COL)].v)

    def test_mfs_skips_valueless_unmapped_headings(self):
        mlabels = {c.v for (r, cc), c in self.wb1.get("mFS").cells.items() if cc == C.RAW_LABEL_COL}
        rlabels = {c.v for (r, cc), c in self.wb1.get("rFS_별도").cells.items() if cc == C.RAW_LABEL_COL}
        self.assertIn("자 산", rlabels)                           # rFS 는 제목 행을 그대로 둔다
        self.assertNotIn("자 산", mlabels)
        self.assertIn("XII. 기본주당이익", mlabels)                 # 값 있는 미매핑 행은 남는다

    def test_old_shape_existing_still_works(self):
        wb2 = C.compose(self.col, {"rFS_별도": {"재무상태표": ["1. 현금및현금성자산"]}}, "p", "s")
        vals = wb2.get("rFS_별도").values_matrix()
        self.assertEqual(vals[4][C.RAW_LABEL_COL - 1], "1. 현금및현금성자산")


class TestHelpers(unittest.TestCase):
    def test_resolve_titles_prefers_existing(self):
        items = [("X", ["a", "b"]), ("X", ["c", "d"])]
        self.assertEqual(C._resolve_titles(items, {}), ["X", "X #2"])
        ex = {"X": {"start": 1, "labels": ["c", "d"]}}
        self.assertEqual(C._resolve_titles(items, ex), ["X #2", "X"])   # 새 표가 앞에 끼어도 기존 표가 'X' 유지
        ex2 = {"X": {"start": 1, "labels": ["a"]}, "X #2": {"start": 9, "labels": ["c"]}}
        items3 = [("X", ["z"]), ("X", ["a"]), ("X", ["c"])]
        self.assertEqual(C._resolve_titles(items3, ex2), ["X #3", "X", "X #2"])
        self.assertEqual(C._resolve_titles([("Y", []), ("X", ["q"])], {"X #2": {"labels": ["q"]}}), ["Y", "X #2"])

    def test_period_of(self):
        self.assertEqual(C._period_of(date(2025, 12, 31)), "FY2025")
        self.assertEqual(C._period_of(46022.0), "FY2025")           # COM Value2 일련번호
        self.assertEqual(C._period_of("2026-06-30 (반기)"), "2026H1")
        self.assertIsNone(C._period_of(True))
        self.assertIsNone(C._period_of(171311183566))
        self.assertIsNone(C._period_of("제 27 기"))

    def test_layout_overflow_math(self):
        warns, shifts = [], []
        ex = {"A": {"start": 1}, "B": {"start": 10}}
        # A 높이 4 → 끝 5, B 는 10 그대로 / A 높이 12 → 끝 13 > 10 → 13+5=18 로 밀림
        self.assertEqual(C._layout("s", [("A", 4), ("B", 3), ("N", 2)], ex, warns, shifts, 1), [1, 10, 18])
        self.assertEqual(C._layout("s", [("A", 12), ("B", 3)], ex, warns, shifts, 1), [1, 18])
        self.assertEqual(shifts, ["s 'B': 블록이 8행 밀림 — 모델 참조 확인 필요"])

    def test_new_block_does_not_run_into_vanished_area(self):
        warns, shifts = [], []
        labels = ["x1", "x2", "x3"]
        ex = {"A": {"start": 1}, "B": {"start": 10}, "Gone": {"start": 20, "labels": labels}}
        got = C._layout("s", [("A", 4), ("B", 3), ("N", 2)], ex, warns, shifts, 1)
        self.assertEqual(got[:2], [1, 10])
        self.assertGreaterEqual(got[2], 20 + 4 + len(labels) + C.RAW_GAP_ROWS)
        self.assertTrue(any("Gone" in w for w in warns))


if __name__ == "__main__":
    unittest.main()
