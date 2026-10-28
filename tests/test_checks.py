"""`검증` 시트(excel/checks.py): 수식 모양(원 환산 → 천원 반올림 비교)·선택 규칙·해당 없음 비고·mFS 요약 행."""
import unittest
from datetime import date
from types import SimpleNamespace as NS
from core import pipeline as P, mapping as M, normalize as N
from core.model import FsRecord, Period
from excel import compose as C, checks as CK, finalize as FZ
from tests.test_pipeline import FakeClient
from tests.test_unit_auto import _hand_col

D, E = C.RAW_LABEL_COL, C.RAW_FIRST_VALUE_COL


def _hdr(s):
    """검증 시트 머리글 1행: {텍스트: 열}."""
    return {c.v: cc for (r, cc), c in s.cells.items() if r == 2 and c.v not in (None, "")}


def _row_of(s, no):
    return next(r for (r, cc), c in s.cells.items() if cc == 2 and c.v == no)


def _val(s, r, c):
    cell = s.cells.get((r, c))
    return cell.v if cell else None


# ---------------------------------------------------------------- 수식 모양

class TestFormulaShape(unittest.TestCase):
    def test_won_then_round_to_thousand(self):
        f, d = CK.check_formula([("+", "'mFS'!E5", C.UNIT_NAME)],
                                [("+", "'mFS'!E6", C.UNIT_NAME), ("+", "'mFS'!E7", C.UNIT_NAME)])
        self.assertEqual(f, "=ROUND('mFS'!E5*cfg_unit/1000,0)=ROUND(('mFS'!E6+'mFS'!E7)*cfg_unit/1000,0)")
        self.assertEqual(d, "=ROUND('mFS'!E5*cfg_unit-('mFS'!E6+'mFS'!E7)*cfg_unit,0)")
        self.assertTrue(f.startswith(FZ.CHECK_PREFIXES))

    def test_unit_one_and_note_unit(self):
        f, d = CK.check_formula([("+", "'rN_판관비'!E19", 1000)], [("+", "'mFS'!E30", C.UNIT_NAME)])
        self.assertEqual(f, "=ROUND('rN_판관비'!E19*1000/1000,0)=ROUND('mFS'!E30*cfg_unit/1000,0)")
        self.assertEqual(d, "=ROUND('rN_판관비'!E19*1000-'mFS'!E30*cfg_unit,0)")
        f, _ = CK.check_formula([("+", "'rFS_별도'!E9", 1)], [("+", "'mFS'!E5", C.UNIT_NAME)])
        self.assertEqual(f, "=ROUND('rFS_별도'!E9/1000,0)=ROUND('mFS'!E5*cfg_unit/1000,0)")

    def test_mixed_sign_and_many_terms(self):
        f, d = CK.check_formula([("+", "'mFS'!E9", C.UNIT_NAME), ("-", "'mFS'!E10", C.UNIT_NAME)],
                                [("+", "'mFS'!E11", C.UNIT_NAME)])
        self.assertEqual(f, "=ROUND(('mFS'!E9-'mFS'!E10)*cfg_unit/1000,0)=ROUND('mFS'!E11*cfg_unit/1000,0)")
        f, d = CK.check_formula([("+", "'rN_x'!E5", 1), ("+", "'rN_x'!E6", 1)], [("+", "'rN_x'!E9", 1)])
        self.assertEqual(f, "=ROUND(('rN_x'!E5+'rN_x'!E6)/1000,0)=ROUND('rN_x'!E9/1000,0)")
        self.assertEqual(d, "=ROUND('rN_x'!E5+'rN_x'!E6-'rN_x'!E9,0)")

    def test_ref_text(self):
        self.assertEqual(CK.ref_text([("+", "'mFS'!E6", C.UNIT_NAME), ("+", "'mFS'!E7", C.UNIT_NAME)]), "mFS!E6+mFS!E7")
        self.assertEqual(CK.ref_text([("+", "'mFS'!E9", C.UNIT_NAME), ("-", "'mFS'!E10", C.UNIT_NAME)]), "mFS!E9-mFS!E10")


# ---------------------------------------------------------------- A 그룹 · 시트 구조 · mFS 요약

class TestSheetA(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wb = C.compose(_hand_col(1000), None, "p", "s")      # 별도 BS 자산·부채·자본총계만, 천원
        cls.s = cls.wb.get("검증")
        cls.h = _hdr(cls.s)

    def test_sheet_position_and_data_flag(self):
        names = [s.name for s in self.wb.sheets]
        self.assertEqual(names[:5], ["설정", "목차", "mFS", "검증", "RAW >>"])
        self.assertIn("검증", C.DATA_SHEETS)
        self.assertTrue(C.is_data_sheet("검증"))
        self.assertFalse(C.is_struct_sheet("검증"))

    def test_header_two_rows_with_rules(self):
        h = self.h
        self.assertEqual((h["구분"], h["번호"], h["항목"]), (1, 2, D))
        self.assertEqual((h["FY2024"], h["FY2025"]), (E, E + 1))
        self.assertEqual(h["차이 FY2024"], E + 3)                      # 빈 열 하나 건너
        self.assertEqual(h["차이 FY2025"], E + 4)
        self.assertEqual((h["좌변 참조"], h["우변 참조"], h["비고"]), (E + 5, E + 6, E + 7))
        last = max(h.values())
        for c in range(1, last + 1):
            self.assertIn(self.s.cells[(2, c)].border, ("top", "both"), c)
            self.assertIn(self.s.cells[(3, c)].border, ("bottom", "both"), c)
        self.assertEqual(self.s.cells[(2, E)].fmt, "hdr")

    def test_section_rows(self):
        secs = [c.v for (r, cc), c in self.s.cells.items() if c.fmt == "section"]
        self.assertEqual(secs, ["A. 재무제표 내부 정합 — 별도", "B. 주석 ↔ 재무제표 — 별도", "C. 주석 합계 검산 — 별도"])

    def test_a1_formula_diff_refs(self):
        s, h = self.s, self.h
        r = _row_of(s, "A1")
        self.assertEqual(_val(s, r, 1), "A")
        self.assertEqual(_val(s, r, D), "자산총계 = 부채총계 + 자본총계")
        cell = s.cells[(r, E)]
        self.assertEqual(cell.v, "=ROUND('mFS'!E5*cfg_unit/1000,0)=ROUND(('mFS'!E6+'mFS'!E7)*cfg_unit/1000,0)")
        self.assertTrue(cell.formula); self.assertEqual(cell.fmt, "check")
        self.assertEqual(s.cells[(r, E + 1)].v, "=ROUND('mFS'!F5*cfg_unit/1000,0)=ROUND(('mFS'!F6+'mFS'!F7)*cfg_unit/1000,0)")
        dcell = s.cells[(r, h["차이 FY2024"])]
        self.assertEqual(dcell.v, "=ROUND('mFS'!E5*cfg_unit-('mFS'!E6+'mFS'!E7)*cfg_unit,0)")
        self.assertTrue(dcell.formula); self.assertEqual(dcell.fmt, "num")
        self.assertEqual(_val(s, r, h["좌변 참조"]), "mFS!E5")
        self.assertEqual(_val(s, r, h["우변 참조"]), "mFS!E6+mFS!E7")
        self.assertIsNone(_val(s, r, h["비고"]))

    def test_a2_not_applicable_with_reason(self):
        s, h = self.s, self.h
        r = _row_of(s, "A2")
        cell = s.cells[(r, E)]
        self.assertEqual((cell.v, cell.fmt, cell.formula), ("해당 없음", "italic", False))
        self.assertIsNone(s.cells.get((r, h["차이 FY2024"])))
        self.assertEqual(_val(s, r, h["비고"]), "없음: 유동자산, 비유동자산")

    def test_a4_only_for_consolidated(self):
        nos = {c.v for (r, cc), c in self.s.cells.items() if cc == 2}
        self.assertNotIn("A4", nos)
        self.assertTrue({"A1", "A5", "A12", "A13", "A14", "B1", "B14"} <= nos)

    def test_a12_half_year_reason(self):
        s, h = self.s, self.h
        r = _row_of(s, "A12")
        self.assertEqual(_val(s, r, E), "해당 없음")
        self.assertIn("직전 연도 없음", _val(s, r, h["비고"]))        # FY2024 열: 직전 연도 열 없음

    def test_mfs_summary_rows(self):
        mfs = self.wb.get("mFS")
        labels = {c.v: r for (r, cc), c in mfs.cells.items() if cc == D}
        self.assertIn("별도 검증 요약", labels)
        self.assertNotIn("별도 검증 (TRUE 여야 함)", labels)
        r1, r2 = CK.row_ranges(self.s)["OFS"]
        rt = labels["검증 TRUE 수"]
        self.assertEqual(mfs.cells[(rt, E)].v, f"=COUNTIF('검증'!E{r1}:E{r2},TRUE)")
        self.assertEqual(mfs.cells[(rt, E + 1)].v, f"=COUNTIF('검증'!F{r1}:F{r2},TRUE)")
        self.assertEqual(mfs.cells[(labels["검증 FALSE 수"], E)].v, f"=COUNTIF('검증'!E{r1}:E{r2},FALSE)")
        self.assertEqual(mfs.cells[(labels["해당 없음 수"], E)].v, f"=COUNTIF('검증'!E{r1}:E{r2},\"해당 없음\")")
        self.assertIn("자세한 내용은 '검증' 시트", [c.v for (r, cc), c in mfs.cells.items() if r == rt])
        self.assertFalse([c.v for c in mfs.cells.values() if c.formula and str(c.v).startswith(FZ.CHECK_PREFIXES)])
        a1 = _row_of(self.s, "A1")
        last = max(r for (r, cc), c in self.s.cells.items() if cc == 2 and str(c.v)[:1] in "ABC" and r > 3)
        self.assertEqual((r1, r2), (a1, last))


class TestCashFlowChecks(unittest.TestCase):
    """A9·A10 은 원문 표시 방식(순증가에 환율변동효과 포함 여부)에 중립 — 데이터로 식을 고르지 않는다."""
    LABS = ["영업활동현금흐름", "투자활동현금흐름", "재무활동현금흐름", "환율변동효과", "현금및현금성자산의 순증가",
            "기초의 현금및현금성자산", "기말의 현금및현금성자산"]

    def _rows(self, labs):
        pers = [Period("FY2024", date(2024, 12, 31), "annual", 26), Period("FY2025", date(2025, 12, 31), "annual", 27)]
        cfg = P.BuildConfig("X", "X", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, None, "default")
        recs = [FsRecord(pk, "OFS", "CF", i, lab, N.norm_label(lab), 10 + i, 1, None, None, "r", "XML", "cur")
                for pk in ("FY2024", "FY2025") for i, lab in enumerate(labs)]
        col = P.Collected(cfg, NS(periods=pers, reports={}), ["FY2024", "FY2025"], fs_records=recs, mapping=M.load())
        wb = C.compose(col, None, "p", "s")
        s = wb.get("검증")
        mfs = wb.get("mFS")
        rows = {c.v: r for (r, cc), c in mfs.cells.items() if cc == D}
        return s, _hdr(s), {no: _row_of(s, no) for no in ("A9", "A10")}, rows

    def test_a9_neutral_formula(self):
        s, h, r, m = self._rows(self.LABS)
        ops, inv, fin, fx = (m[k] for k in ("영업활동현금흐름", "투자활동현금흐름", "재무활동현금흐름", "환율변동효과"))
        beg, end = m["기초현금및현금성자산"], m["기말현금및현금성자산"]
        f = s.cells[(r["A9"], E)].v
        self.assertEqual(f, f"=ROUND(('mFS'!E{ops}+'mFS'!E{inv}+'mFS'!E{fin}+'mFS'!E{fx})*cfg_unit/1000,0)"
                            f"=ROUND(('mFS'!E{end}-'mFS'!E{beg})*cfg_unit/1000,0)")
        self.assertEqual(s.cells[(r["A9"], h["차이 FY2024"])].v,
                         f"=ROUND(('mFS'!E{ops}+'mFS'!E{inv}+'mFS'!E{fin}+'mFS'!E{fx})*cfg_unit-('mFS'!E{end}-'mFS'!E{beg})*cfg_unit,0)")
        self.assertEqual(_val(s, r["A9"], h["우변 참조"]), f"mFS!E{end}-mFS!E{beg}")
        self.assertEqual(_val(s, r["A9"], D), "영업+투자+재무+환율변동효과 = 기말현금 − 기초현금")
        self.assertIsNone(_val(s, r["A9"], h["비고"]))

    def test_a10_or_form_with_fx(self):
        s, h, r, m = self._rows(self.LABS)
        beg, inc, fx, end = (m[k] for k in ("기초현금및현금성자산", "현금및현금성자산의증가", "환율변동효과", "기말현금및현금성자산"))
        cell = s.cells[(r["A10"], E)]
        self.assertEqual(cell.v, f"=OR(ROUND(('mFS'!E{beg}+'mFS'!E{inc})*cfg_unit/1000,0)=ROUND('mFS'!E{end}*cfg_unit/1000,0),"
                                 f"ROUND(('mFS'!E{beg}+'mFS'!E{inc}+'mFS'!E{fx})*cfg_unit/1000,0)=ROUND('mFS'!E{end}*cfg_unit/1000,0))")
        self.assertEqual((cell.fmt, cell.formula), ("check", True))
        self.assertTrue(cell.v.startswith(FZ.CHECK_PREFIXES)); self.assertTrue(FZ.is_check_formula(cell.v))
        self.assertEqual(s.cells[(r["A10"], h["차이 FY2024"])].v,
                         f"=ROUND(('mFS'!E{beg}+'mFS'!E{inc})*cfg_unit-'mFS'!E{end}*cfg_unit,0)")       # 좌변(환율 제외 식) − 우변
        self.assertEqual(_val(s, r["A10"], h["좌변 참조"]), f"mFS!E{beg}+mFS!E{inc} | +mFS!E{fx}")
        self.assertEqual(_val(s, r["A10"], h["비고"]), "순증가에 환율효과 포함·제외 어느 쪽이든 인정; 차이는 환율 제외 식 기준")
        self.assertEqual(_val(s, r["A10"], D), "기초현금 + 현금증가(환율 포함/제외) = 기말현금")

    def test_without_fx_row(self):
        s, h, r, m = self._rows([x for x in self.LABS if x != "환율변동효과"])
        self.assertEqual(_val(s, r["A9"], h["좌변 참조"]).count("+"), 2)
        self.assertEqual(_val(s, r["A9"], h["비고"]), "환율변동효과 행 없음")
        self.assertTrue(s.cells[(r["A9"], E)].v.startswith("=ROUND(("))
        f = s.cells[(r["A10"], E)].v
        self.assertTrue(f.startswith("=ROUND(("), f); self.assertNotIn("OR(", f)
        self.assertIsNone(_val(s, r["A10"], h["비고"]))


# ---------------------------------------------------------------- 주석 선택기 _note_row

def _blk(topic, title, col, rows, unit=1000, fs_div="OFS", note_no="11", sheet="rN_x", row0=10, periods=("FY2025",)):
    """rows: [(라벨, 값)] 한 기간(FY2025). 값이 dict 면 {기간: 값}. [Placed] — 행 번호는 row0 부터(블록 순서를 따지는 테스트는 row0 를 다르게)."""
    b = NS(fs_div=fs_div, topic=topic, periods=list(periods), col_label=col, title=title, unit=unit, note_no=note_no,
           note_title=topic)
    sec = f"주석 {note_no}. {topic} | {title}" + (f" — {col}" if col else "")
    return [C.Placed(sheet, b, sec, NS(label=lab, norm=N.norm_label(lab), values=dict(v) if isinstance(v, dict) else {"FY2025": v}),
                     row0 + i, 0) for i, (lab, v) in enumerate(rows)]


def _ctx(periods, **stds):
    """검증 ctx 흉내: stds = {"IS_매출액": mFS 행, ...} — 모든 기간이 채워진 것으로 본다."""
    rows_of = {("OFS", k.split("_", 1)[0], k.split("_", 1)[1]): r for k, r in stds.items()}
    return {"periods": list(periods), "rows_of": rows_of,
            "filled": {(fs, sj, std, i) for (fs, sj, std) in rows_of for i in range(len(periods))}}


def _items(ctx, ps):
    return {it.no: it for it in CK._b_items(ctx, "OFS", ps)}


class TestNoteRow(unittest.TestCase):
    def test_end_row_form(self):
        ps = _blk("유형자산", "유형자산의 변동", "합 계", [("기 초", 100), ("취 득", 20), ("상 각", -5), ("기 말", 115)])
        p = CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end")
        self.assertEqual(p.row.values["FY2025"], 115)
        ps = _blk("유형자산", "유형자산의 변동", None, [("기초금액", 100), ("증가", 20), ("기말금액", 120)])
        self.assertEqual(CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end").row.values["FY2025"], 120)

    def test_end_rejects_total_only_block(self):
        """가상자산 내역 — 장부가액 처럼 합계 행만 있는 표는 변동표가 아니다 → 합계 폴백 없음."""
        ps = _blk("무형자산", "가상자산 내역", "장부가액", [("비트코인", 1), ("이더리움", 2), ("합계", 3)])
        self.assertIsNone(CK._note_row(ps, "OFS", "무형자산", "FY2025", how="end"))
        ps += _blk("무형자산", "무형자산 내역", "취득원가", [("영업권", 9), ("합계", 9)])
        ps += _blk("무형자산", "무형자산 내역", "합 계", [("영업권", 7), ("합계", 7)])
        self.assertIsNone(CK._note_row(ps, "OFS", "무형자산", "FY2025", how="end"))

    def test_end_rejects_wrong_column_and_two_end_rows(self):
        ps = _blk("유형자산", "변동", "취득원가", [("기초", 1), ("기말", 2)])          # 총액 열
        self.assertIsNone(CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end"))
        ps = _blk("유형자산", "변동", None, [("기초", 1), ("기말", 2), ("기말잔액", 2)])
        r = CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end")          # 값 있는 기말 행 둘 → 해당 없음 + 사유
        self.assertFalse(r)
        self.assertEqual(r.reason, "기말 행 여러 개: 변동")

    def test_end_column_form_needs_sibling_begin_column(self):
        """열마다 블록인 변동표(기초·취득·…·기말 열): 같은 표에 기초 열이 있을 때만 기말(장부금액) 열의 합계 행."""
        rows = [("토지", 1), ("건물", 2), ("합 계", 3)]
        ps = _blk("유형자산", "유형자산 변동내역", "기초", rows) + _blk("유형자산", "유형자산 변동내역", "취득", rows)
        ps += _blk("유형자산", "유형자산 변동내역", "장부금액", [("토지", 5), ("건물", 6), ("합 계", 11)])
        self.assertEqual(CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end").row.values["FY2025"], 11)
        # 내역표(취득원가·누계액·장부가액)만 있으면 기초 열이 없어 고르지 않는다
        ps2 = _blk("유형자산", "유형자산 내역", "취득원가", rows) + _blk("유형자산", "유형자산 내역", "장부가액", rows)
        self.assertIsNone(CK._note_row(ps2, "OFS", "유형자산", "FY2025", how="end"))
        # 합계류 열은 그 표의 마지막 열일 때만(중간 소계 열은 아님)
        mid = _blk("유형자산", "변동", "기초", rows, row0=10) + _blk("유형자산", "변동", "합계", [("토지", 7), ("건물", 8), ("합 계", 15)], row0=20) \
            + _blk("유형자산", "변동", "취득", rows, row0=30)
        self.assertIsNone(CK._note_row(mid, "OFS", "유형자산", "FY2025", how="end"))
        last = _blk("유형자산", "변동", "기초", rows, row0=10) + _blk("유형자산", "변동", "취득", rows, row0=20) \
            + _blk("유형자산", "변동", "합계", [("토지", 7), ("건물", 8), ("합 계", 15)], row0=30)
        self.assertEqual(CK._note_row(last, "OFS", "유형자산", "FY2025", how="end").row.values["FY2025"], 15)
        # 다른 표의 '장부금액' 열(행이 다름)은 기초 열의 형제가 아니다
        ps3 = ps2 + _blk("유형자산", "유형자산 변동내역", "기초", rows) + _blk("유형자산", "유형자산 변동내역", "장부금액", [("리스채권", 4)])
        self.assertIsNone(CK._note_row(ps3, "OFS", "유형자산", "FY2025", how="end"))

    def test_end_where_filter(self):
        rows = [("기초", 1), ("기말", 2)]
        ps = _blk("리스", "리스부채의 변동", None, rows) + _blk("리스", "사용권자산의 변동", "합계", [("기초", 3), ("기말", 4)])
        p = CK._note_row(ps, "OFS", "리스", "FY2025", how="end", where=lambda b: CK.has(b, "사용권자산"))
        self.assertEqual(p.row.values["FY2025"], 4)

    def test_end_skips_right_of_use_subtable_in_ppe_note(self):
        """유형자산 주석 안의 '유형자산으로 인식한 사용권자산의 변동내역'(기초…기말 행)은 B3 가 아니라 B10 몫."""
        rows = [("토지", 1), ("건물", 2), ("합 계", 3)]
        ps = _blk("유형자산", "① 당기 중 유형자산으로 인식한 사용권자산의 변동내역은 다음과 같습니다.", "합계",
                  [("기초", 90), ("취득", 9), ("기말", 99)])
        ps += _blk("유형자산", "① 제 25(당) 기", "기 초", rows) + _blk("유형자산", "① 제 25(당) 기", "기 말", [("토지", 5), ("건물", 6), ("합 계", 11)])
        self.assertEqual(CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end").row.values["FY2025"], 99)      # 조건 없으면 행 변동표가 먼저
        p = CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end", where=CK.main_ppe)
        self.assertEqual((p.section[-3:], p.row.values["FY2025"]), ("기 말", 11))
        self.assertEqual(CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end", where=CK.is_rou).row.values["FY2025"], 99)
        gw = _blk("무형자산", "현금창출단위별 영업권의 변동", None, [("기초", 1), ("기말", 2)])
        self.assertIsNone(CK._note_row(gw, "OFS", "무형자산", "FY2025", how="end", where=CK.main_intangible))
        both = _blk("무형자산", "영업권 및 무형자산의 변동", None, [("기초", 1), ("기말", 2)])
        self.assertIsNotNone(CK._note_row(both, "OFS", "무형자산", "FY2025", how="end", where=CK.main_intangible))

    def test_total_exact_then_suffix(self):
        ps = _blk("성격별비용", "비용의 성격별 분류", None, [("원재료의 사용", 1), ("매출원가, 판매비와관리비 합계", 9)])
        self.assertEqual(CK._note_row(ps, "OFS", "성격별비용", "FY2025", how="total").row.values["FY2025"], 9)
        ps = _blk("판관비", "판관비", "당기", [("급여", 1), ("합계", 2)]) + _blk("판관비", "판관비", None, [("급여", 1), ("합계", 3)])
        self.assertEqual(CK._note_row(ps, "OFS", "판관비", "FY2025", how="total").row.values["FY2025"], 3)   # col_label 없는 블록만
        ps = _blk("판관비", "판관비", None, [("급여", 1), ("합계", None)])
        self.assertIsNone(CK._note_row(ps, "OFS", "판관비", "FY2025", how="total"))                         # 그 기간 값 없음

    def test_row_labels_priority_and_col_none_first(self):
        ps = _blk("리스", "리스부채 만기", "합계", [("1년 이내", 1), ("합계", 2)]) + _blk("리스", "리스부채 내역", None, [("리스부채", 7), ("합계", 8)])
        p = CK._note_row(ps, "OFS", "리스", "FY2025", how="row", labels=("리스부채", "합계"))
        self.assertEqual(p.row.values["FY2025"], 7)
        p = CK._note_row(ps, "OFS", "리스", "FY2025", how="row", labels=("합계",))
        self.assertEqual(p.row.values["FY2025"], 8)                                                           # col None 우선
        where = lambda b: CK.has(b, "리스부채")                                                               # B9: 라벨 둘 다 제목/열 조건
        bare = _blk("리스", "(단위: 천원)", None, [("리스부채", 5), ("합계", 6)])
        self.assertIsNone(CK._note_row(bare, "OFS", "리스", "FY2025", how="row", labels=("리스부채", "합계"), where=where))
        self.assertEqual(CK._note_row(ps, "OFS", "리스", "FY2025", how="row", labels=("리스부채", "합계"), where=where).row.values["FY2025"], 7)
        ps = _blk("법인세", "법인세", None, [("법인세비용", 1), ("법인세비용", 2)])
        self.assertIsNone(CK._note_row(ps, "OFS", "법인세", "FY2025", how="row", labels=("법인세비용",)))     # 둘 → 불명확

    def test_income_cost_where(self):
        ps = _blk("금융손익", "금융수익의 내역", None, [("이자수익", 1), ("합계", 1)]) \
            + _blk("금융손익", "금융원가의 내역", None, [("이자비용", 2), ("합계", 2)])
        inc = CK._note_row(ps, "OFS", "금융손익", "FY2025", how="total", where=lambda b: not CK.is_cost(b))
        cost = CK._note_row(ps, "OFS", "금융손익", "FY2025", how="total", where=CK.is_cost)
        self.assertEqual((inc.row.values["FY2025"], cost.row.values["FY2025"]), (1, 2))


# ---------------------------------------------------------------- C 그룹

class TestCGroup(unittest.TestCase):
    def test_total_sums_subtotals_only_and_skips_percent(self):
        ps = _blk("차입금", "차입금", None, [("단기", 1), ("장기", 2), ("소계", 3), ("사채", 4), ("비율", N.Percent(0.1)),
                                          ("소계", 4), ("합계", 7), ("주석", 99)])
        res = CK.c_total(ps, "FY2025")
        self.assertEqual(res.lhs, [("+", "'rN_x'!E12", 1000), ("+", "'rN_x'!E15", 1000)])
        self.assertEqual(res.rhs, [("+", "'rN_x'!E16", 1000)])
        ps = _blk("판관비", "판관비", None, [("급여", 1), ("비율", N.Percent(0.5)), ("임차료", 2), ("합계", 3)])
        res = CK.c_total(ps, "FY2025")
        self.assertEqual(res.lhs, [("+", "'rN_x'!E10", 1000), ("+", "'rN_x'!E12", 1000)])
        self.assertEqual(CK.ref_text(res.lhs, ranges=True), "rN_x!E10+rN_x!E12")

    def test_total_not_applicable_cases(self):
        self.assertIsNone(CK.c_total(_blk("판관비", "x", None, [("급여", 1)]), "FY2025"))                  # 합계 행 없음
        self.assertIsNone(CK.c_total(_blk("판관비", "x", None, [("급여", 1), ("계", 1), ("합계", 1)]), "FY2025"))   # 둘
        res = CK.c_total(_blk("판관비", "x", None, [("급여", 1), ("합계", None)]), "FY2025")
        self.assertEqual((res.lhs, res.reason), (None, "합계 값 없음"))
        res = CK.c_total(_blk("판관비", "x", None, [("급여", "텍스트"), ("합계", 1)]), "FY2025")
        self.assertEqual(res.reason, "더할 숫자 행 없음")

    def test_roll_forward(self):
        ps = _blk("충당부채", "충당부채 변동", None, [("기초금액", 10), ("증가", 5), ("감소", -2), ("기말금액", 13), ("비고", 0)])
        res = CK.c_roll(ps, "FY2025")
        self.assertEqual(res.lhs, [("+", "'rN_x'!E10", 1000), ("+", "'rN_x'!E11", 1000), ("+", "'rN_x'!E12", 1000)])
        self.assertEqual(res.rhs, [("+", "'rN_x'!E13", 1000)])
        self.assertEqual(CK.ref_text(res.lhs, ranges=True), "rN_x!E10:E12")
        res = CK.c_roll(_blk("유형자산", "x", None, [("기초", 1), ("취득원가", 2), ("감가상각누계액", -1), ("기말", 2)]), "FY2025")
        self.assertEqual((res.lhs, res.reason), (None, "총액·누계액 구분 표"))
        self.assertIsNone(CK.c_roll(_blk("유형자산", "x", None, [("기초", 1), ("증가", 1)]), "FY2025"))
        res = CK.c_roll(_blk("유형자산", "x", None, [("기초", None), ("증가", 1), ("기말", 1)]), "FY2025")
        self.assertEqual(res.reason, "기초·기말 값 없음")


# ---------------------------------------------------------------- 오비 픽스처(B·C 실제 블록)

class TestFixtureOb(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, {"recommended": True}, "default")
        cls.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        cls.wb = C.compose(cls.col, None, "p", "s")
        cls.s = cls.wb.get("검증")
        cls.h = _hdr(cls.s)

    def test_b1_sga_note_vs_is(self):
        s, h = self.s, self.h
        r = _row_of(s, "B1")
        f = s.cells[(r, E)].v
        self.assertRegex(f, r"^=ROUND\('rN_판관비'!E\d+\*1000/1000,0\)=ROUND\('mFS'!E\d+\*cfg_unit/1000,0\)$")
        self.assertRegex(_val(s, r, h["좌변 참조"]), r"^rN_판관비!E\d+$")
        self.assertRegex(_val(s, r, h["우변 참조"]), r"^mFS!E\d+$")
        rr = int(f.split("'rN_판관비'!E")[1].split("*")[0])
        self.assertEqual(self.wb.get("rN_판관비").cells[(rr, D)].v, "합 계")

    def test_b2_two_is_terms(self):
        f = self.s.cells[(_row_of(self.s, "B2"), E)].v
        self.assertRegex(f, r"^=ROUND\('rN_성격별비용'!E\d+\*1000/1000,0\)=ROUND\(\('mFS'!E\d+\+'mFS'!E\d+\)\*cfg_unit/1000,0\)$")

    def test_b3_end_column_form_picks_change_table(self):
        """오비 유형자산 변동표는 열(기초·취득·…·장부금액)마다 블록 → '장부금액' 열의 합 계 행. 내역표 '합 계' 열은 아니다."""
        s = self.s
        r = _row_of(s, "B3")
        f = s.cells[(r, E)].v
        self.assertRegex(f, r"^=ROUND\('rN_유형자산'!E\d+\*1000/1000,0\)=ROUND\('mFS'!E\d+\*cfg_unit/1000,0\)$")
        rr = int(f.split("'rN_유형자산'!E")[1].split("*")[0])
        raw = self.wb.get("rN_유형자산")
        self.assertEqual(raw.cells[(rr, D)].v, "합 계")
        sec = max(((row, c.v) for (row, cc), c in raw.cells.items() if cc == D and c.fmt == "section" and row < rr))[1]
        self.assertIn("변동내역", sec); self.assertTrue(sec.endswith("장부금액"), sec)

    def test_b5_b6_finance(self):
        s = self.s
        for no, word in (("B5", "금융수익"), ("B6", "금융비용")):
            r = _row_of(s, no)
            self.assertIn("'rN_금융손익'!E", s.cells[(r, E)].v, no)
            self.assertIn(word, _val(s, r, D))

    def test_c_rows_cover_note_blocks(self):
        s, h = self.s, self.h
        rows = [(r, c.v) for (r, cc), c in s.cells.items() if cc == 2 and str(c.v).startswith("C")]
        self.assertGreater(len(rows), 5)
        labels = [_val(s, r, D) for r, _ in rows]
        self.assertTrue(any(l.startswith("rN_판관비 · ") and l.endswith(" · 합계") for l in labels), labels[:5])
        self.assertTrue(any(l.startswith("rN_충당부채 · ") and l.endswith(" · 변동") for l in labels), labels)
        self.assertTrue(all(len(l) <= 80 for l in labels))
        self.assertFalse(any("주석 " in l and ". " in l.split(" · ")[1][:12] for l in labels), labels[:3])   # 캡션만, '주석 27. …' 접두 없음
        self.assertTrue(any(l == "rN_판관비 · 판매비, 관리비와 물류비 · 합계" for l in labels), labels[:3])      # 캡션이 단위뿐 → 주석 제목
        self.assertTrue(any(l == "rN_유형자산 · 유형자산 — 취득원가 · 합계" for l in labels), [l for l in labels if "유형자산" in l][:3])
        self.assertTrue(any(l.startswith("rN_유형자산 · (2) 당기 및 전기 중 유형자산 장부금액의 변동내역") for l in labels))   # 내용 있는 캡션은 그대로
        self.assertEqual(len(labels), len(set(labels)), [l for l in labels if labels.count(l) > 1])           # 같은 시트 중복은 #2·#3
        self.assertIn("rN_리스 · 리스 · 합계", labels); self.assertIn("rN_리스 · 리스 · 합계 #2", labels)
        self.assertFalse(any(l.startswith("rN_기타") for l in labels))
        r = next(r for r, _ in rows if _val(s, r, D).startswith("rN_판관비"))
        self.assertRegex(s.cells[(r, E)].v, r"^=ROUND\(\('rN_판관비'!E\d+(\+'rN_판관비'!E\d+)+\)\*1000/1000,0\)=ROUND\('rN_판관비'!E\d+\*1000/1000,0\)$")
        self.assertRegex(_val(s, r, h["좌변 참조"]), r"^rN_판관비!E\d+:E\d+$")

    def test_not_applicable_rows_name_missing(self):
        s, h = self.s, self.h
        r = _row_of(s, "B8")                                    # 차입금 표는 열 블록(유동·비유동)뿐 → 합계 블록 없음
        self.assertEqual(_val(s, r, E), "해당 없음")
        self.assertIn("없음: ", _val(s, r, h["비고"]))
        self.assertIn("없음: 주석 차입금 합계 행", _val(s, r, h["비고"]))
        self.assertIn("우변 = 단기차입금 (사채 제외)", _val(s, r, h["비고"]))    # Task 5(h): 우변은 있는 BS 표준계정만, 사채는 캡션에 있을 때만

    def test_all_check_cells_have_check_fmt_or_na(self):
        s, h = self.s, self.h
        for (r, c), cell in s.cells.items():
            if c in (E, E + 1) and r > 3 and cell.v not in (None, ""):
                if cell.formula:
                    self.assertTrue(str(cell.v).startswith("=ROUND("), cell.v)
                    self.assertEqual(cell.fmt, "check")
                else:
                    self.assertEqual((cell.v, cell.fmt), ("해당 없음", "italic"))


# ---------------------------------------------------------------- Task 5 — 선택기 결함(5-2 d~k)

P2 = ("FY2021", "FY2025")


class TestRevenuePick(unittest.TestCase):
    """(d) B7: 캡션에 계약자산·계약부채·이월·부문·지역·비유동자산 이 없는 첫 블록, 기간마다 첫 합계 행(값 없으면 첫 소계 행)."""
    CAP = "연결회사가 고객과의 계약에서 생기는 수익으로 인식한 금액 및 구성내역은 다음과 같습니다. (단위: 천원)"

    def _ps(self):
        a = _blk("매출구분", self.CAP, None, [("클라우드 서비스", {"FY2025": 10}), ("기타매출", {"FY2021": 7, "FY2025": 2}),
                                              ("소계", {"FY2021": 7, "FY2025": 12}), ("임대료수입", {"FY2025": 1}),
                                              ("합계", {"FY2025": 13}), ("한시점에 이행", {"FY2021": 7, "FY2025": 13}),
                                              ("소계", {"FY2021": 7}), ("합계", {"FY2025": 13})], periods=P2, row0=10)
        b = _blk("매출구분", "당기 및 전기에 인식한 수익 중 전기에서 이월된 계약부채와 관련된 금액은 다음과 같습니다. (단위: 천원)", None,
                 [("이월 계약부채", {"FY2021": 5, "FY2025": 5}), ("합계", {"FY2021": 5, "FY2025": 5})], periods=P2, row0=30)
        c = _blk("매출구분", "(2) 당기와 전기 중 부문별 손익은 다음과 같습니다. (단위: 천원)", "총부문수익",
                 [("IT서비스", {"FY2021": 9, "FY2025": 9}), ("합계", {"FY2021": 9, "FY2025": 9})], periods=P2, row0=40, note_no="35")
        return a, b, c

    def test_first_total_then_first_subtotal(self):
        a, b, c = self._ps()
        it = _items(_ctx(P2, IS_매출액=68), b + a + c)        # 이월 블록이 시트상 먼저 와도 캡션 낱말로 제외
        st = it["B7"].per["FY2025"]
        self.assertEqual(st, ("ok", [CK._term(a[4], "FY2025")], [("+", "'mFS'!F68", C.UNIT_NAME)]))
        st = it["B7"].per["FY2021"]
        self.assertEqual(st, ("ok", [CK._term(a[2], "FY2021")], [("+", "'mFS'!E68", C.UNIT_NAME)]))
        self.assertEqual(it["B7"].note, "FY2021 는 소계")

    def test_excluded_words_and_no_candidate(self):
        a, b, c = self._ps()
        it = _items(_ctx(P2, IS_매출액=68), b + c)
        self.assertEqual(it["B7"].per["FY2025"], ("na", "없음: 주석 매출 구분 합계 행"))
        self.assertIsNone(it["B7"].note)
        for word in ("계약자산", "지역", "비유동자산", "연결회사의 주요 수익은"):
            x = _blk("매출구분", f"{word} 관련 표 (단위: 천원)", None, [("x", {"FY2025": 1}), ("합계", {"FY2025": 1})], periods=P2)
            self.assertEqual(_items(_ctx(P2, IS_매출액=68), x)["B7"].per["FY2025"][0], "na", word)

    def test_block_chosen_per_period(self):
        """첫 후보 블록이 일부 연도만 다루면(연도별 양식 차이) 기간마다 값 있는 첫 블록. 뒤 블록을 골랐으면 비고에 블록 캡션."""
        a = _blk("매출구분", "수익의 구분 (단위: 천원)", None, [("상품", {"FY2021": 1}), ("합계", {"FY2021": 1})], periods=P2, row0=10)
        c = _blk("매출구분", "고객과의 계약에서 생기는 수익 (단위: 천원)", None, [("용역", {"FY2025": 2}), ("합계", {"FY2025": 2})],
                 periods=P2, row0=30)
        it = _items(_ctx(P2, IS_매출액=68), a + c)
        self.assertEqual(it["B7"].per["FY2021"][1], [CK._term(a[1], "FY2021")])
        self.assertEqual(it["B7"].per["FY2025"][1], [CK._term(c[1], "FY2025")])
        self.assertEqual(it["B7"].note, "FY2025 는 블록 '고객과의 계약에서 생기는 수익 (단위: 천원)'")

    def test_block_without_total_rows_is_skipped(self):
        """합계·소계 행이 없는 매출 구성표(총매출액·매출액(순액))는 구조상 합계 표가 아니다 → 다음 후보 블록(휴젤 픽스처)."""
        a = _blk("매출구분", "(2) 당기 및 전기 매출액의 구성내역은 다음과 같습니다.", None,
                 [("제품매출", 1), ("총매출액", 1), ("반품 예상액 조정", 0), ("매출액(순액)", 1)], row0=10)
        b = _blk("매출구분", "(3) 당기 및 전기 지역별 매출액은 다음과 같습니다.", None, [("국내", 1), ("합 계", 1)], row0=20)
        c = _blk("매출구분", "고객과의 계약에서 생기는 수익을 주요 지리적 시장과 주요 제품에 따라 구분한 내역", None,
                 [("국내", 1), ("합 계", 1), ("톡신", 1), ("합 계", 1)], row0=30)
        it = _items(_ctx(("FY2025",), IS_매출액=68), a + b + c)
        self.assertEqual(it["B7"].per["FY2025"][1], [CK._term(c[1], "FY2025")])       # a: 합계 없음, b: '지역' 제외, c: 첫 합계
        self.assertTrue(it["B7"].note.startswith("블록 '고객과의 계약에서 생기는 수익을 주요 지리적 시장과"), it["B7"].note)
        self.assertEqual(_items(_ctx(("FY2025",), IS_매출액=68), a)["B7"].per["FY2025"][0], "na")

    def test_total_without_value_and_no_subtotal(self):
        x = _blk("매출구분", self.CAP, None, [("상품매출", {"FY2021": 1}), ("합계", {"FY2025": 1})], periods=P2)
        it = _items(_ctx(P2, IS_매출액=68), x)
        self.assertEqual(it["B7"].per["FY2025"][0], "ok")
        self.assertEqual(it["B7"].per["FY2021"], ("na", "없음: 주석 매출 구분 합계 행"))


class TestSubtotalRule(unittest.TestCase):
    """(e) C-합계: 소계 행들 + 마지막 소계 뒤의 비소계 행들(합계 앞까지)."""
    def test_rows_after_last_subtotal_are_added(self):
        ps = _blk("매출구분", "부문별 손익", "총부문수익", [("IT서비스", 1), ("기타", 2), ("소계", 3), ("연결조정 등", -1), ("합계", 2)])
        res = CK.c_total(ps, "FY2025")
        self.assertEqual(res.lhs, [("+", "'rN_x'!E12", 1000), ("+", "'rN_x'!E13", 1000)])
        self.assertEqual(res.rhs, [("+", "'rN_x'!E14", 1000)])
        ps = _blk("차입금", "차입금", None, [("단기", 1), ("소계", 1), ("장기", 2), ("소계", 2), ("합계", 3)])   # 소계 사이 행은 안 더함
        self.assertEqual(CK.c_total(ps, "FY2025").lhs, [("+", "'rN_x'!E11", 1000), ("+", "'rN_x'!E13", 1000)])


class TestAdjustmentTable(unittest.TestCase):
    """(8) C-합계: 캡션에 회계이익·유효세율 이 든 법인세 조정표는 첫 행(회계이익)이 합산 대상이 아니라 검산하지 않는다 → '조정표'."""
    def test_skipped_with_reason(self):
        rows = [("법인세비용차감전순이익", 100), ("적용세율에 따른 세부담액", 22), ("조정사항", 1), ("합계", 23)]
        for cap in ("법인세비용과 회계이익의 관계는 다음과 같습니다. (단위: 천원)", "(3) 유효세율 조정내역 (단위: 천원)"):
            res = CK.c_total(_blk("법인세", cap, None, rows), "FY2025")
            self.assertEqual((res.lhs, res.reason), (None, "조정표"), cap)
        self.assertIsNotNone(CK.c_total(_blk("법인세", "법인세비용의 구성 (단위: 천원)", None, rows), "FY2025").lhs)


class TestMfsDuplicateRows(unittest.TestCase):
    """(9) mFS: 같은 기간에 정규화 라벨·값이 같은 raw 행('비지배지분'·'Ⅰ. 비지배지분' — 머리글·항목 되풀이)은 첫 행만 더한다."""
    @staticmethod
    def _p(label, vals, rownum, unit=1):
        b = NS(periods=["FY2021", "FY2022"], unit=unit, sj_div="BS")
        return C.Placed("rFS_연결", b, "재무상태표", NS(label=label, norm=N.norm_label(label), values=vals, account_id=None), rownum, 0)

    def test_same_label_and_value_counted_once(self):
        ps = [self._p("비지배지분", {"FY2021": 25726063055, "FY2022": 7}, 40), self._p("Ⅰ. 비지배지분", {"FY2021": 25726063055, "FY2022": 9}, 60)]
        warns = []
        self.assertEqual(C._sum_formula(ps, "FY2021", warns), "='rFS_연결'!E40/cfg_unit")
        self.assertEqual(warns, ["중복 행 1회만 합산: Ⅰ. 비지배지분 FY2021"])
        self.assertEqual(C._sum_formula(ps, "FY2022", warns), "=('rFS_연결'!F40+'rFS_연결'!F60)/cfg_unit")     # 값이 다르면 둘 다
        self.assertEqual(len(warns), 1)
        other = [self._p("비지배지분", {"FY2021": 5}, 40), self._p("지배기업소유주지분", {"FY2021": 5}, 41)]        # 라벨이 다르면 둘 다
        self.assertEqual(C._sum_formula(other, "FY2021"), "=('rFS_연결'!E40+'rFS_연결'!E41)/cfg_unit")


class TestEndAmbiguity(unittest.TestCase):
    """(f) how="end": 열형·합계류 후보가 없고 col_label None 행형 후보가 그 기간에 값 있는 것이 2개 이상이면 해당 없음 + 비고."""
    T1 = "(1) 손실충당부채 변동내역은 다음과 같습니다. (단위: 천원)"
    T2 = "복구충당부채 변동내역은 다음과 같습니다. (단위: 천원)"

    def _two(self):
        return (_blk("충당부채", self.T1, None, [("기초", {"FY2025": 122}), ("설정", {"FY2025": 55}), ("기말", {"FY2025": 54})], periods=P2, row0=10)
                + _blk("충당부채", self.T2, None, [("기초", {"FY2021": 1, "FY2025": 4877}), ("설정", {"FY2025": 7446}),
                                                   ("기말", {"FY2021": 2, "FY2025": 11371})], periods=P2, row0=20))

    def test_two_bare_candidates_is_not_applicable(self):
        ps = self._two()
        r = CK._note_row(ps, "OFS", "충당부채", "FY2025", how="end")
        self.assertFalse(r)
        self.assertEqual(r.reason, f"변동표 여러 개: {self.T1}, {self.T2} — 합계 표 없음")
        self.assertEqual(CK._note_row(ps, "OFS", "충당부채", "FY2021", how="end").row.values["FY2021"], 2)   # 그 기간 값 있는 것 하나뿐
        it = _items(_ctx(P2, BS_충당부채=23), ps)
        self.assertEqual(it["B12"].per["FY2025"], ("na", r.reason))
        self.assertEqual(it["B12"].per["FY2021"][0], "ok")

    def test_total_column_or_labeled_candidate_wins(self):
        ps = self._two() + _blk("충당부채", "충당부채 변동", "합계", [("기초", {"FY2025": 5}), ("기말", {"FY2025": 9})], periods=P2, row0=30)
        self.assertEqual(CK._note_row(ps, "OFS", "충당부채", "FY2025", how="end").row.values["FY2025"], 9)
        rows = [("손실충당부채", {"FY2025": 1}), ("복구충당부채", {"FY2025": 2}), ("합계", {"FY2025": 3})]
        col = (_blk("충당부채", "변동", "기초", rows, periods=P2, row0=40) + _blk("충당부채", "변동", "기말", rows, periods=P2, row0=50))
        self.assertEqual(CK._note_row(self._two() + col, "OFS", "충당부채", "FY2025", how="end").row.values["FY2025"], 3)

    def test_first_keeps_reason_when_all_fail(self):
        ps = self._two()
        sel = lambda pk: CK._note_row(ps, "OFS", "충당부채", pk, how="end")
        none = lambda pk: None
        r = CK._first(none, sel)("FY2025")
        self.assertFalse(r); self.assertIn("변동표 여러 개", r.reason)
        self.assertEqual(CK._first(sel, lambda pk: ps[-1])("FY2025"), ps[-1])         # 뒤 선택기가 맞으면 그것


class TestSubheadingPick(unittest.TestCase):
    """(g) B9·B10: 소제목 행(리스부채·사용권자산) 아래 첫 합계 행. 소제목 행에 값이 있으면 그 행."""
    CAP = "리스와 관련해 연결재무상태표에 인식된 금액은 다음과 같습니다. (단위: 천원)"
    ROWS = [("사용권자산", None), ("건물", 69), ("IDC 시설", 1), ("차량운반구", 1), ("합계", 71),
            ("리스부채(*)", None), ("유동리스부채", 7), ("비유동리스부채", 51), ("합계", 58)]

    def test_sub_total_after_heading(self):
        ps = _blk("리스", self.CAP, None, self.ROWS)
        r = CK._note_row(ps, "OFS", "리스", "FY2025", how="sub", labels=("리스부채",))
        self.assertEqual([p.row.values["FY2025"] for p in r.rows], [58]); self.assertEqual(r.note, "리스부채 소제목 아래 합계")
        r = CK._note_row(ps, "OFS", "리스", "FY2025", how="sub", labels=("사용권자산",))
        self.assertEqual([p.row.values["FY2025"] for p in r.rows], [71]); self.assertEqual(r.note, "사용권자산 소제목 아래 합계")
        ps2 = _blk("리스", self.CAP, None, [("리스부채", 60), ("유동", 7), ("합계", 58)])
        r = CK._note_row(ps2, "OFS", "리스", "FY2025", how="sub", labels=("리스부채",))
        self.assertEqual([p.row.values["FY2025"] for p in r.rows], [60]); self.assertIsNone(r.note)        # 소제목 행에 값
        self.assertIsNone(CK._note_row(_blk("리스", "만기분석", None, [("1년 이내", 1), ("합계", 2)]), "OFS", "리스", "FY2025",
                                       how="sub", labels=("리스부채",)))

    def test_b9_b10_items(self):
        ps = _blk("리스", self.CAP, None, self.ROWS)
        it = _items(_ctx(("FY2025",), BS_유동리스부채=22, BS_비유동리스부채=32, BS_사용권자산=12), ps)
        self.assertEqual(it["B9"].per["FY2025"], ("ok", [CK._term(ps[8], "FY2025")],
                                                  [("+", "'mFS'!E22", C.UNIT_NAME), ("+", "'mFS'!E32", C.UNIT_NAME)]))
        self.assertEqual(it["B9"].note, "리스부채 소제목 아래 합계")
        self.assertEqual(it["B10"].per["FY2025"], ("ok", [CK._term(ps[4], "FY2025")], [("+", "'mFS'!E12", C.UNIT_NAME)]))
        self.assertEqual(it["B10"].note, "사용권자산 소제목 아래 합계")
        roll = _blk("리스", "사용권자산의 변동", None, [("기초", 1), ("기말", 2)], row0=50)       # 변동표가 있으면 그것이 먼저
        it = _items(_ctx(("FY2025",), BS_사용권자산=12), roll + ps)
        self.assertEqual(it["B10"].per["FY2025"][1], [CK._term(roll[1], "FY2025")]); self.assertIsNone(it["B10"].note)


class TestBorrowingsPick(unittest.TestCase):
    """(h) B8: 합계 행이 여러 개면 모두 더한다. 우변은 있는 BS 표준계정만(사채는 캡션에 '사채' 가 있을 때만)."""
    CAP = "보고기간말 연결회사의 차입금 현황은 다음과 같습니다. (단위: 천원)"
    ROWS = [("단기차입금 / 원화", 5), ("단기차입금 / 소계", 5), ("유동성장기차입금 / 소계", 3), ("합계", 8),
            ("장기차입금 / 원화", 10), ("합계", 10)]

    def test_sum_all_total_rows(self):
        ps = _blk("차입금", self.CAP, None, self.ROWS)
        r = CK._note_row(ps, "OFS", "차입금", "FY2025", how="total_all", where=lambda b: CK.has(b, "차입금"))
        self.assertEqual([p.row.values["FY2025"] for p in r.rows], [8, 10]); self.assertEqual(r.note, "합계 행 2개 합산")
        one = _blk("차입금", self.CAP, None, [("단기차입금", 5), ("합계", 5)])
        r = CK._note_row(one, "OFS", "차입금", "FY2025", how="total_all")
        self.assertEqual(len(r.rows), 1); self.assertIsNone(r.note)
        self.assertIsNone(CK._note_row(_blk("차입금", self.CAP, None, [("단기차입금", 5)]), "OFS", "차입금", "FY2025", how="total_all"))

    def test_grand_total_after_group_totals_is_not_double_counted(self):
        """그룹 합계 + 총합계: 마지막 합계 앞에 숫자 행이 없으면 마지막 행만(구조 판단)."""
        ps = _blk("차입금", self.CAP, None, [("단기차입금", 5), ("합계", 5), ("장기차입금", 10), ("합계", 10), ("합계", 15)])
        r = CK._note_row(ps, "OFS", "차입금", "FY2025", how="total_all")
        self.assertEqual([p.row.values["FY2025"] for p in r.rows], [15]); self.assertIsNone(r.note)
        ps = _blk("차입금", self.CAP, None, [("단기차입금", 5), ("합계", 5), ("장기차입금", 10), ("합계", 10)])    # 사이에 숫자 행 → 둘 다
        self.assertEqual([p.row.values["FY2025"] for p in CK._note_row(ps, "OFS", "차입금", "FY2025", how="total_all").rows], [5, 10])
        ps = _blk("차입금", self.CAP, None, [("단기차입금", 5), ("합계", 5), ("장기차입금", 10), ("합계", 10), ("비고", "담보"), ("합계", 15)])
        self.assertEqual([p.row.values["FY2025"] for p in CK._note_row(ps, "OFS", "차입금", "FY2025", how="total_all").rows], [15])

    def test_b8_rhs_present_stds_only(self):
        ps = _blk("차입금", self.CAP, None, self.ROWS)
        it = _items(_ctx(("FY2025",), BS_단기차입금=20, BS_장기차입금=31, BS_사채=33), ps)
        st = it["B8"].per["FY2025"]
        self.assertEqual(st, ("ok", [CK._term(ps[3], "FY2025"), CK._term(ps[5], "FY2025")],
                              [("+", "'mFS'!E20", C.UNIT_NAME), ("+", "'mFS'!E31", C.UNIT_NAME)]))
        self.assertIn("합계 행 2개 합산", it["B8"].note); self.assertIn("우변 = 단기차입금 + 장기차입금", it["B8"].note)
        self.assertNotIn("사채", it["B8"].note.replace("사채 제외", ""))
        bond = _blk("차입금", "차입금 및 사채의 내역 (단위: 천원)", None, self.ROWS)
        it = _items(_ctx(("FY2025",), BS_단기차입금=20, BS_장기차입금=31, BS_사채=33), bond)
        self.assertEqual(it["B8"].per["FY2025"][2][-1], ("+", "'mFS'!E33", C.UNIT_NAME))
        self.assertIn("우변 = 단기차입금 + 장기차입금 + 사채", it["B8"].note)
        it = _items(_ctx(("FY2025",), IS_매출액=68), ps)
        self.assertEqual(it["B8"].per["FY2025"], ("na", "없음: 단기차입금, 유동성장기부채, 장기차입금"))


class TestTaxPick(unittest.TestCase):
    """(i) B13: `^법인세(비용|수익)` 로 시작하고 당기법인세·이연법인세·총 이 없는 행 중 마지막 행. '(A+B)' 꼬리 무시."""
    ROWS = [("당기법인세", None), ("당기손익에 대한 당기법인세", 2500), ("총당기법인세(A)", 2604), ("이연법인세", None),
            ("일시적차이의 증감", 4585), ("총이연법인세(B)", -2733), ("법인세수익(A+B)", -128)]

    def test_last_matching_row(self):
        ps = _blk("법인세", "(1) 법인세비용(수익)의 내역 (단위: 천원)", None, self.ROWS)
        it = _items(_ctx(("FY2025",), IS_법인세비용=80), ps)
        self.assertEqual(it["B13"].per["FY2025"], ("ok", [CK._term(ps[6], "FY2025")], [("+", "'mFS'!E80", C.UNIT_NAME)]))
        ps = _blk("법인세", "x", None, [("법인세비용차감전순손실", -9), ("법인세비용", 1), ("기타", 0), ("법인세비용(수익)", 2)])
        self.assertEqual(_items(_ctx(("FY2025",), IS_법인세비용=80), ps)["B13"].per["FY2025"][1], [CK._term(ps[3], "FY2025")])
        ps = _blk("법인세", "x", None, [("총법인세비용", 1), ("당기법인세비용", 2)])      # 둘 다 제외 낱말 → 없음
        self.assertEqual(_items(_ctx(("FY2025",), IS_법인세비용=80), ps)["B13"].per["FY2025"][0], "na")
        ps = _blk("법인세", "x", None, [("법인세비용", 1), ("법인세비용차감전순이익", 9)])   # '차감전' 은 법인세 행이 아니다(마지막이어도)
        self.assertEqual(_items(_ctx(("FY2025",), IS_법인세비용=80), ps)["B13"].per["FY2025"][1], [CK._term(ps[0], "FY2025")])


class TestCapitalPick(unittest.TestCase):
    """(j) B14: 열 라벨 '자본금(천원)'(단위 표기는 비교에서 뗀다) 블록의 합계 행. 합계 행이 없으면 해당 없음 '합계 행 없음'
    (연도별 표가 합쳐져 보통주 행이 중복될 수 있어 행 합산은 하지 않는다)."""
    CAP = "(1) 자본금의 내역은 다음과 같습니다."

    def test_total_row_or_not_applicable(self):
        ps = _blk("자본", self.CAP, "자본금(천원)", [("보통주", 200000), ("우선주(*1)", None), ("보통주", 200000), ("우선주", 161086)])
        it = _items(_ctx(("FY2025",), BS_자본금=40), ps)
        self.assertEqual(it["B14"].per["FY2025"], ("na", "합계 행 없음"))
        tot = _blk("자본", self.CAP, "자본금(천원)", [("보통주", 200000), ("우선주", 161086), ("합계", 361086)])
        it = _items(_ctx(("FY2025",), BS_자본금=40), tot)
        self.assertEqual(it["B14"].per["FY2025"], ("ok", [CK._term(tot[2], "FY2025")], [("+", "'mFS'!E40", C.UNIT_NAME)]))
        self.assertIsNone(it["B14"].note)
        it = _items(_ctx(("FY2025",), BS_자본금=40), _blk("자본", self.CAP, "발행주식수(주)", [("보통주", 1)]))
        self.assertEqual(it["B14"].per["FY2025"][0], "na")


class TestSceNetIncomeByPeriod(unittest.TestCase):
    """(k) A14: 총계 블록에 당기순손실(FY2024)·당기순이익(FY2025) 행이 따로 있어도 기간마다 값 있는 행이 하나면 그 행."""
    def test_a14_per_period(self):
        pers = [Period("FY2024", date(2024, 12, 31), "annual", 26), Period("FY2025", date(2025, 12, 31), "annual", 27),
                Period.half(2026, 28)]
        cfg = P.BuildConfig("X", "X", date(2026, 10, 26), 2024, 2025, True, ["OFS"], 1, None, "default")
        sce = [("2025.1.1 (당기초)", {"FY2024": 100, "FY2025": 95, "2026H1": 103}), ("당기순손실", {"FY2024": -5, "FY2025": None, "2026H1": None}),
               ("당기순이익", {"FY2024": None, "FY2025": 8, "2026H1": 3}), ("2025.12.31 (당기말)", {"FY2024": 95, "FY2025": 103, "2026H1": 106})]
        recs = [FsRecord(pk, "OFS", "SCE", i, lab, N.sce_row_norm(lab), v[pk], 1, "총계", None, "r", "XML", "cur")
                for pk in ("FY2024", "FY2025", "2026H1") for i, (lab, v) in enumerate(sce)]
        recs += [FsRecord(pk, "OFS", "IS", 0, "당기순이익(손실)", "당기순이익(손실)", v, 1, None, None, "r", "XML", "cur")
                 for pk, v in (("FY2024", -5), ("FY2025", 8), ("2026H1", 3))]
        col = P.Collected(cfg, NS(periods=pers, reports={}), ["FY2024", "FY2025", "2026H1"], fs_records=recs, mapping=M.load())
        wb = C.compose(col, None, "p", "s")
        s = wb.get("검증")
        h = _hdr(s)
        r = _row_of(s, "A14")
        self.assertEqual(s.cells[(r, E + 2)].v, "해당 없음")                      # 반기 열은 연차 자본변동표와 다른 행 → 제외
        self.assertIn("반기 자본변동표 제외(2026H1)", _val(s, r, h["비고"]))
        self.assertNotIn("반기", _val(s, _row_of(s, "A13"), h["비고"]) or "")                  # A13 은 반기 제외 규칙 없음
        e, f = s.cells[(r, E)].v, s.cells[(r, E + 1)].v
        self.assertRegex(e, r"^=ROUND\('rFS_별도'!E\d+/1000,0\)=ROUND\('mFS'!E\d+\*cfg_unit/1000,0\)$")
        self.assertRegex(f, r"^=ROUND\('rFS_별도'!F\d+/1000,0\)=ROUND\('mFS'!F\d+\*cfg_unit/1000,0\)$")
        raw = wb.get("rFS_별도")
        re_, rf = (int(x.split("'rFS_별도'!")[1][1:].split("/")[0]) for x in (e, f))
        self.assertEqual((raw.cells[(re_, D)].v, raw.cells[(rf, D)].v), ("당기순손실", "당기순이익"))


class TestBeginEndLabels(unittest.TestCase):
    """당기초·당기말 도 기초·기말 행. 연도 병합 블록(FY2024 '기말' + FY2025 '당기말')은 기간마다 값 있는 행을 고르고,
    값 있는 기초/기말 행이 여럿이면 '해당 없음'(비고 기초/기말 행 여러 개) — 조용히 빼지 않는다."""
    P = ("FY2024", "FY2025")
    ROWS = [("기초", {"FY2024": 100, "FY2025": 110}), ("취득", {"FY2024": 15, "FY2025": 5}), ("기타 (*1)", {"FY2024": -5, "FY2025": 1}),
            ("기말", {"FY2024": 110}), ("당기말", {"FY2025": 116})]

    def test_current_end_label_end_and_roll(self):
        ps = _blk("유형자산", "사용권자산의 변동", None, self.ROWS, periods=self.P)
        self.assertEqual(CK._note_row(ps, "OFS", "유형자산", "FY2025", how="end").row.values["FY2025"], 116)
        self.assertEqual(CK._note_row(ps, "OFS", "유형자산", "FY2024", how="end").row.values["FY2024"], 110)
        res = CK.c_roll(ps, "FY2025")
        self.assertEqual(res.lhs, [CK._term(ps[0], "FY2025"), CK._term(ps[1], "FY2025"), CK._term(ps[2], "FY2025")])
        self.assertEqual(res.rhs, [CK._term(ps[4], "FY2025")])
        res = CK.c_roll(ps, "FY2024")                                   # 사이의 '기말'(FY2025 값 없음)은 FY2025 식에서 빠진다
        self.assertEqual(res.rhs, [CK._term(ps[3], "FY2024")])
        only = _blk("충당부채", "변동", None, [("당기초", 1), ("증가", 2), ("당기말", 3)])
        self.assertEqual(CK.c_roll(only, "FY2025").rhs, [CK._term(only[2], "FY2025")])
        self.assertEqual(CK._note_row(only, "OFS", "충당부채", "FY2025", how="end").row.values["FY2025"], 3)

    def test_roll_with_subtotals_sums_subtotals_only(self):
        """퇴직급여 변동표(오비): 당기초 + 당기손익 소계 + 재측정요소 소계 + 기여금·지급액 = 당기말 — 항목과 소계를 둘 다 더하지 않는다."""
        ps = _blk("퇴직급여", "확정급여채무 변동", None, [("당기초", 100), ("근무원가", 10), ("이자비용", 2), ("소 계", 12),
                                                     ("보험수리적손실", 3), ("소 계", 3), ("지급액", -5), ("당기말", 110)])
        res = CK.c_roll(ps, "FY2025")
        self.assertEqual(res.lhs, [CK._term(ps[i], "FY2025") for i in (0, 3, 5, 6)])
        self.assertEqual(res.rhs, [CK._term(ps[7], "FY2025")])

    def test_multiple_valued_rows_are_not_applicable(self):
        ps = _blk("충당부채", "변동", None, [("기초", 1), ("증가", 2), ("기말", 3), ("당기말", 3)])
        self.assertEqual(CK.c_roll(ps, "FY2025"), CK.CRes(None, None, "기말 행 여러 개"))
        ps = _blk("충당부채", "변동", None, [("기초", 1), ("당기초", 1), ("증가", 2), ("기말", 3)])
        self.assertEqual(CK.c_roll(ps, "FY2025"), CK.CRes(None, None, "기초 행 여러 개"))
        ctx = _ctx(("FY2025",), BS_충당부채=23)
        items = [it for it in CK._c_items(ctx, "OFS", ps) if it.label.endswith("변동")]
        self.assertEqual(len(items), 1)                                 # 검산 항목은 남고 사유가 비고로
        self.assertEqual(items[0].per["FY2025"], ("na", "기초 행 여러 개"))
        it = _items(ctx, _blk("충당부채", "변동", None, [("기초", 1), ("기말", 3), ("당기말", 3)]))
        self.assertEqual(it["B12"].per["FY2025"], ("na", "기말 행 여러 개: 변동"))


class TestLeaseMaturityExcluded(unittest.TestCase):
    """B9: '합계' 라벨 후보에서 캡션에 만기·할인·현금흐름 이 든 표(리스부채 만기분석)는 뺀다. '리스부채' 행은 그대로."""
    def test_maturity_total_not_picked(self):
        ctx = _ctx(("FY2025",), BS_유동리스부채=22, BS_비유동리스부채=32)
        for cap in ("리스부채의 만기분석 (단위: 천원)", "리스부채의 할인되지 않은 계약상 현금흐름 (단위: 천원)"):
            mat = _blk("리스", cap, None, [("1년 이내", 10), ("1년 초과", 50), ("합계", 60)])
            self.assertEqual(_items(ctx, mat)["B9"].per["FY2025"][0], "na", cap)
            bal = _blk("리스", "리스부채 내역 (단위: 천원)", None, [("유동", 7), ("비유동", 51), ("합계", 58)], row0=40)
            self.assertEqual(_items(ctx, mat + bal)["B9"].per["FY2025"][1], [CK._term(bal[2], "FY2025")], cap)
        disc = _blk("리스", "리스부채의 할인 전 현금흐름 (단위: 천원)", None, [("합계", 60), ("현재가치할인차금", -2), ("리스부채", 58)])
        self.assertEqual(_items(ctx, disc)["B9"].per["FY2025"][1], [CK._term(disc[2], "FY2025")])


class TestSceTotalRow(unittest.TestCase):
    """A13: API 원천 상장사는 연차 기말이 '자본총계' 행, 반기 기말이 '기말자본' 행에 갈려 있다 — 기간마다 값 있는 행이 하나면 그 행."""
    def _wb(self, extra=()):
        pers = [Period("FY2024", date(2024, 12, 31), "annual", 26), Period("FY2025", date(2025, 12, 31), "annual", 27),
                Period.half(2026, 28)]
        cfg = P.BuildConfig("X", "X", date(2026, 10, 26), 2024, 2025, True, ["OFS"], 1, None, "default")
        sce = [("기초자본", {"FY2024": 90, "FY2025": 100, "2026H1": 120}), ("당기순이익", {"FY2024": 10, "FY2025": 20, "2026H1": 5}),
               ("기말자본", {"FY2024": None, "FY2025": None, "2026H1": 125}),
               ("자본총계", {"FY2024": 100, "FY2025": 120, "2026H1": None})] + list(extra)
        recs = [FsRecord(pk, "OFS", "SCE", i, lab, N.sce_row_norm(lab), v[pk], 1, "총계", None, "r", "API", "cur")
                for pk in ("FY2024", "FY2025", "2026H1") for i, (lab, v) in enumerate(sce)]
        recs += [FsRecord(pk, "OFS", "BS", 0, "자본총계", "자본총계", v, 1, None, None, "r", "API", "cur")
                 for pk, v in (("FY2024", 100), ("FY2025", 120), ("2026H1", 125))]
        col = P.Collected(cfg, NS(periods=pers, reports={}), ["FY2024", "FY2025", "2026H1"], fs_records=recs, mapping=M.load())
        return C.compose(col, None, "p", "s")

    def test_disjoint_period_rows(self):
        wb = self._wb()
        s, raw = wb.get("검증"), wb.get("rFS_별도")
        h, r = _hdr(s), _row_of(s, "A13")
        labels = []
        for k in range(3):
            f = s.cells[(r, E + k)].v
            self.assertRegex(f, r"^=ROUND\('rFS_별도'!\w\d+/1000,0\)=ROUND\('mFS'!\w\d+\*cfg_unit/1000,0\)$")
            rr = int("".join(ch for ch in f.split("'rFS_별도'!")[1].split("/")[0] if ch.isdigit()))
            labels.append(raw.cells[(rr, D)].v)
        self.assertEqual(labels, ["자본총계", "자본총계", "기말자본"])
        self.assertEqual(_val(s, r, h["비고"]), "FY2024·FY2025 는 자본총계 행")

    def test_two_valued_rows_not_applicable(self):
        wb = self._wb(extra=[("2025.12.31 (당기말)", {"FY2024": None, "FY2025": 120, "2026H1": None})])
        s = wb.get("검증")
        h, r = _hdr(s), _row_of(s, "A13")
        self.assertEqual(s.cells[(r, E + 1)].v, "해당 없음")
        self.assertIn("자본변동표 기말 행 여러 개(FY2025)", _val(s, r, h["비고"]))
        self.assertTrue(s.cells[(r, E)].v.startswith("=ROUND("))


# ---------------------------------------------------------------- finalize.verify: 검증 시트만, bool 값만

class _FakeRange:
    def __init__(self, vals, forms):
        self.Row, self.Column, self.Value2, self.Formula = 1, 1, vals, forms

    def SpecialCells(self, *a):
        import pywintypes
        raise pywintypes.com_error(FZ.NO_CELLS_HRESULT, "없음", None, None)


class TestVerify(unittest.TestCase):
    def test_is_check_formula_by_shape(self):
        self.assertTrue(FZ.is_check_formula("=ROUND('mFS'!E5*cfg_unit/1000,0)=ROUND('mFS'!E6*cfg_unit/1000,0)"))
        self.assertTrue(FZ.is_check_formula("=OR(ROUND((E1+E2)/1000,0)=ROUND(E3/1000,0),ROUND((E1+E2+E4)/1000,0)=ROUND(E3/1000,0))"))
        self.assertFalse(FZ.is_check_formula("=ABS(ROUND(E1-E2,0))<=1"))                                   # 옛 ±1 모양은 더 안 받음
        self.assertFalse(FZ.is_check_formula("=ROUND('mFS'!E5*cfg_unit-'mFS'!E6*cfg_unit,0)"))     # 차이 열
        self.assertFalse(FZ.is_check_formula("=COUNTIF('검증'!E6:E89,TRUE)"))
        self.assertFalse(FZ.is_check_formula(None)); self.assertFalse(FZ.is_check_formula(5))

    def test_scan_counts_errors_as_false(self):
        vals = ((True, -5, "해당 없음"), (False, 3, None), (-2146826281, True, None))        # 3행 A: #N/A 오류 코드
        forms = (("=ROUND(1,0)=ROUND(1,0)", "=ROUND(1-6,0)", "해당 없음"),
                 ("=ROUND(1,0)=ROUND(2,0)", "=COUNTIF(A1:A2,TRUE)", ""),
                 ("=ROUND(X/1000,0)=ROUND(1,0)", "=OR(ROUND(1,0)=ROUND(1,0),ROUND(2,0)=ROUND(1,0))", ""))
        res = {"errors": 0, "checks": {"TRUE": 0, "FALSE": 0, "해당 없음": 0}, "false_cells": []}
        FZ.scan_checks([list(r) for r in vals], [list(r) for r in forms], 1, 1, "검증", res)
        self.assertEqual(res["checks"], {"TRUE": 2, "FALSE": 2, "해당 없음": 1})
        self.assertEqual(res["false_cells"], ["검증!A2", "검증!A3"])

    def test_counts_only_check_sheet(self):
        chk = NS(Name="검증", UsedRange=_FakeRange(((True, -5, "해당 없음"), (False, 3, None)),
                                                   (("=ROUND(1,0)=ROUND(1,0)", "=ROUND(1-6,0)", "해당 없음"),
                                                    ("=ROUND(1,0)=ROUND(2,0)", "=COUNTIF(A1:A2,TRUE)", ""))))
        mfs = NS(Name="mFS", UsedRange=_FakeRange(((3, "해당 없음"),), (("=COUNTIF('검증'!E5:E9,TRUE)", "해당 없음"),)))
        res = FZ.verify(NS(Worksheets=[mfs, chk]))
        self.assertEqual(res["checks"], {"TRUE": 1, "FALSE": 1, "해당 없음": 1})
        self.assertEqual(res["false_cells"], ["검증!A2"])
        self.assertEqual(res["errors"], 0)


if __name__ == "__main__":
    unittest.main()
