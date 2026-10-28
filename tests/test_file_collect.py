"""core.file_collect: 회사 제시 엑셀(가상 '샘플전자' 픽스처 2개 연도) → Collected → compose 까지(COM 없이)."""
import json
import os
import tempfile
import unittest
from datetime import date

from openpyxl import Workbook

from core import file_collect as FC, mapping as M, pipeline as P
from excel import compose as C, write_openpyxl as W
from tests.fixtures import xlsx_fixture as F

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out", "file_collect")


def _file(path, period_end, kind="annual", fs_div="OFS", unit=None):
    return {"path": path, "period_end": period_end, "kind": kind, "fs_div": fs_div, "unit": unit}


def _cfg(files, layout=None, **kw):
    args = dict(corp_code="FILE:샘플전자", corp_name="샘플전자", as_of=date(2026, 10, 27), start_year=2024, end_year=2025,
                include_half=False, fs_divs=["OFS"], unit_display=1, note_include=None, preset="default",
                source="file", files=files, layout=layout)
    args.update(kw)
    return P.BuildConfig(**args)


def _wb(path, sheets):
    """{시트명: [행...]} → xlsx."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = Workbook()
    first = True
    for name, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet(name)
        ws.title = name
        first = False
        for row in rows:
            ws.append(list(row))
    wb.save(path)
    return path


class TestBuildConfig(unittest.TestCase):
    def test_source_missing_means_dart(self):
        cfg = P.BuildConfig.from_dict({"corp_code": "OB", "as_of": "2026-10-26", "start_year": 2024, "end_year": 2025})
        self.assertEqual((cfg.source, cfg.files, cfg.layout), ("dart", None, None))
        self.assertEqual(cfg.to_dict()["source"], "dart")

    def test_file_roundtrip(self):
        files = [_file(r"C:\x\a.xlsx", "2025-12-31"), _file(r"C:\x\b.xlsx", date(2024, 12, 31))]
        layout = {"a.xlsx|재무상태표|0": {"periods": {"3": "FY2025", "4": None}, "unit": 1000}}
        cfg = _cfg(files, layout)
        js = json.dumps(cfg.to_dict(), ensure_ascii=False)          # date 가 섞여도 JSON 으로 나간다
        back = P.BuildConfig.from_dict(json.loads(js))
        self.assertEqual(back.source, "file")
        self.assertEqual(back.files[1]["period_end"], "2024-12-31")
        self.assertEqual(back.layout, layout)
        self.assertEqual(back.corp_code, "FILE:샘플전자")


class TestTwoYears(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.makedirs(OUT, exist_ok=True)
        cls.st25 = F.build_statements(os.path.join(OUT, "statements_2025.xlsx"))
        cls.st24 = F.build_statements(os.path.join(OUT, "statements_2024.xlsx"), year=2024)
        cls.nt25 = F.build_notes(os.path.join(OUT, "notes_2025.xlsx"))
        cls.nt24 = F.build_notes(os.path.join(OUT, "notes_2024.xlsx"), year=2024)
        cls.cfg = _cfg([_file(cls.st24, "2024-12-31"), _file(cls.nt24, "2024-12-31"),
                        _file(cls.st25, "2025-12-31"), _file(cls.nt25, "2025-12-31")])
        cls.logs = []
        cls.col = FC.collect_files(cls.cfg, M.load(), cls.logs.append)
        cls.rc25, cls.rc24 = "FILE:statements_2025.xlsx", "FILE:statements_2024.xlsx"

    def _fs(self, pk, sj):
        return [r for r in self.col.fs_records if r.period_key == pk and r.sj_div == sj]

    def test_periods_and_plan(self):
        self.assertEqual(self.col.periods, ["FY2024", "FY2025"])
        self.assertEqual([p.key for p in self.col.plan.periods], ["FY2024", "FY2025"])
        self.assertEqual(self.col.plan.sources["FY2024"], [(self.rc25, "prev"), (self.rc24, "cur")])
        self.assertEqual(self.col.plan.sources["FY2025"], [(self.rc25, "cur")])
        self.assertEqual(self.col.plan.base_report.rcept_no, self.rc25)
        rep = self.col.plan.reports[self.rc25]
        self.assertEqual((rep.report_nm, rep.kind, rep.rtype, rep.period_end, rep.fs_divs),
                         ("statements_2025.xlsx", "annual", "파일", date(2025, 12, 31), ("OFS",)))
        self.assertIsInstance(rep.rcept_dt, date)

    def test_fs_records_restated(self):
        self.assertEqual({r.sj_div for r in self.col.fs_records}, {"BS", "IS", "CF", "SCE"})
        self.assertTrue(all(r.source == "XLSX" and r.fs_div == "OFS" and r.account_id is None for r in self.col.fs_records))
        bs25 = self._fs("FY2025", "BS")
        self.assertTrue(bs25 and all(r.rcept_no == self.rc25 and r.role == "cur" for r in bs25))
        bs24 = self._fs("FY2024", "BS")
        self.assertTrue(bs24 and all(r.rcept_no == self.rc25 and r.role == "prev" for r in bs24))
        ar = next(r for r in bs24 if r.label == "매출채권")
        self.assertEqual((ar.amount, ar.unit, ar.ord), (220 * F.M, 1, 4))      # 재작성(2025 파일 전기) 값, ord = 표 안 행 순서(자산·유동자산·현금 다음)
        self.assertEqual(next(r for r in bs24 if r.label == "자산").amount, None)   # 구분 행도 원문 그대로
        d = [x for x in self.col.restate_diffs if "매출채권" in x["key"]]
        self.assertEqual(len(d), 1)
        self.assertEqual((d[0]["채택"], d[0]["원보고서"], d[0]["채택출처"], d[0]["원보고서출처"]),
                         (220 * F.M, 221 * F.M, self.rc25, self.rc24))
        cf = self._fs("FY2025", "CF")
        self.assertIn("1. 당기순이익", [r.label for r in cf])            # 2단 라벨(A·B열) 중 비어 있지 않은 쪽
        self.assertEqual(next(r for r in cf if r.label == "Ⅱ. 투자활동현금흐름").amount, -150 * F.M)

    def test_sce(self):
        sce25 = self._fs("FY2025", "SCE")
        self.assertEqual({r.detail for r in sce25}, {"자본금", "이익잉여금", "총계"})
        end = [r for r in sce25 if r.detail == "총계" and r.label_norm == "기말"]
        self.assertEqual([(r.label, r.amount, r.role) for r in end], [("2025.12.31(당기말)", 900 * F.M, "cur")])
        sce24 = self._fs("FY2024", "SCE")
        self.assertTrue(all(r.rcept_no == self.rc25 and r.role == "prev" for r in sce24))
        self.assertEqual(next(r.amount for r in sce24 if r.detail == "총계" and r.label_norm == "기말"), 700 * F.M)
        self.assertEqual(next(r.amount for r in sce24 if r.detail == "이익잉여금" and r.label_norm == "당기순이익"), 120 * F.M)

    def test_notes_and_toc(self):
        notes = self.col.note_records
        self.assertEqual({r.topic for r in notes}, {"판관비", "유형자산"})
        self.assertTrue(all(r.raw_cell is None and r.fs_div == "OFS" and r.unit == 1 and r.unit_known for r in notes))
        sga = [r for r in notes if r.topic == "판관비"]
        self.assertEqual({(r.note_no, r.note_title, r.caption) for r in sga}, {("27", "판매비와관리비", "27. 판매비와관리비")})
        tot24 = next(r for r in sga if r.period_key == "FY2024" and r.row_norm == "합계")
        self.assertEqual((tot24.value, tot24.rcept_no, tot24.role, tot24.col_label), (330 * F.M, "FILE:notes_2025.xlsx", "prev", None))
        ppe = [r for r in notes if r.topic == "유형자산"]
        self.assertEqual({r.period_key for r in ppe}, {"FY2024", "FY2025"})       # 2024 파일 당기표 + 2025 파일 당기표
        self.assertEqual({r.col_label for r in ppe}, {"토지", "건물", "기계장치", "합계"})
        end25 = next(r for r in ppe if r.period_key == "FY2025" and r.col_label == "합계" and r.row_norm == "기말")
        self.assertEqual((end25.value, end25.note_no, end25.note_title), (800 * F.M, "12", "유형자산"))
        self.assertEqual(end25.caption, "(1) 당기 중 유형자산의 변동내역은 다음과 같습니다.")
        end24 = next(r for r in ppe if r.period_key == "FY2024" and r.col_label == "합계" and r.row_norm == "기말")
        self.assertEqual((end24.value, end24.rcept_no), (750 * F.M, "FILE:notes_2024.xlsx"))
        toc = self.col.toc
        self.assertEqual([(e.no, e.title, e.topic, e.recommended, e.include) for e in toc],
                         [("12", "유형자산", "유형자산", True, True), ("27", "판매비와관리비", "판관비", True, True)])
        self.assertEqual([e.table_count for e in toc], [2, 2])

    def test_units_and_logs(self):
        u = self.col.units
        self.assertEqual((u[("OFS", "BS")], u[("OFS", "IS")], u[("OFS", "CF")], u[("OFS", "SCE")]), (1, 1, 1, 1))
        self.assertEqual((u[("fiscal", "FY2025")], u[("fiscal", "FY2024")]), (12, 11))
        self.assertIn("파일 statements_2025.xlsx 시트 재무상태표 표 1: BS FY2025(당기)·FY2024(전기) 단위 원", self.col.logs)
        self.assertIn("파일 notes_2025.xlsx 시트 주석 표 2: NOTE FY2025(당기) 단위 원", self.col.logs)
        self.assertEqual(self.logs, self.col.logs)
        self.assertFalse(self.col.auditor_rows)
        self.assertEqual(self.col.plan.inferred, [])

    def test_compose_to_openpyxl(self):
        wb = C.compose(self.col, None, r"C:\py\python.exe", r"C:\app\excel\refresh.py")
        names = [s.name for s in wb.sheets]
        for n in ("설정", "목차", "mFS", "검증", "rFS_별도", "rN_판관비", "rN_유형자산", "데이터", "매핑표", "로그"):
            self.assertIn(n, names)
        vals = wb.get("검증").values_matrix()
        nos = {r[1]: r for r in vals if r[1]}
        for no in ("A1", "A5", "B1", "B3"):
            f = str(nos[no][C.RAW_FIRST_VALUE_COL])          # FY2025 열
            self.assertTrue(f.startswith("=ROUND("), (no, f))
        self.assertIn("'rN_판관비'!", str(nos["B1"][C.RAW_FIRST_VALUE_COL]))
        self.assertIn("'rN_유형자산'!", str(nos["B3"][C.RAW_FIRST_VALUE_COL]))
        self.assertTrue(str(nos["B1"][C.RAW_FIRST_VALUE_COL - 1]).startswith("=ROUND("))   # FY2024 도(2025 파일 전기 열)
        log = wb.get("로그").values_matrix()
        self.assertIn(["기간말", "파일", "시트(표)", "수정일", "단위"], [r[:5] for r in log])
        row = next(r for r in log if r[1] == "statements_2025.xlsx")
        self.assertEqual(row[0], date(2025, 12, 31))
        self.assertIn("재무상태표(표 1)", row[2])
        self.assertIsInstance(row[3], date)
        self.assertEqual(row[4], "원")
        self.assertFalse(any(r[0] == "감사인·의견" for r in log))               # 파일 입력엔 감사인 구역 없음
        st = wb.get("설정")
        got = {st.cells[(r, 2)].v: (r, st.cells.get((r, 3))) for (r, c) in st.cells if c == 2}
        self.assertEqual(got["corp_code"][1].v, "파일 입력")
        r_files, cell = got["입력 파일"]
        self.assertEqual(cell.v.split("\n"), [self.st24, self.nt24, self.st25, self.nt25])
        self.assertEqual(cell.fmt, "wrap")
        self.assertEqual(r_files, got["연결/별도"][0] + 1)
        self.assertIn("설정 JSON", st.cells[(r_files, 4)].v)
        self.assertEqual(st.names["cfg_json"], f"C{got['설정 JSON'][0]}")
        self.assertEqual(st.names["cfg_status"], f"C{got['상태'][0]}")
        js = json.loads(got["설정 JSON"][1].v)
        self.assertEqual((js["source"], [d["path"] for d in js["files"]]), ("file", [self.st24, self.nt24, self.st25, self.nt25]))
        with tempfile.TemporaryDirectory() as d:
            p = W.write(wb, os.path.join(d, "file.xlsx"))
            self.assertTrue(os.path.exists(p))


class TestVariants(unittest.TestCase):
    def test_other_table_logged_not_converted(self):
        tb = F.build_trial_balance(os.path.join(OUT, "tb.xlsx"))
        col = FC.collect_files(_cfg([_file(tb, "2025-12-31")]), M.load(), None)
        self.assertEqual(col.fs_records, [])
        self.assertIn("파일 tb.xlsx 시트 시산표 표 1: 표 인식 못 함(기타)", col.logs)
        self.assertEqual(col.periods, ["FY2025"])                     # 파일의 결산일은 항상 기간이 된다

    def test_missing_file(self):
        with self.assertRaises(FileNotFoundError) as cm:
            FC.collect_files(_cfg([_file(os.path.join(OUT, "없음.xlsx"), "2025-12-31")]), M.load(), None)
        self.assertIn("입력 파일을 찾을 수 없습니다", str(cm.exception))

    def test_layout_override(self):
        st = F.build_statements(os.path.join(OUT, "layout.xlsx"))
        layout = {"layout.xlsx|재무상태표|0": {"periods": {"4": None}, "unit": 1000},
                  "layout.xlsx|손익계산서|0": {"periods": {"2": "FY2025", "3": "FY2023"}},
                  "layout.xlsx|현금흐름표|0": {"kind": "OTHER"}}
        col = FC.collect_files(_cfg([_file(st, "2025-12-31")], layout, start_year=2023), M.load(), None)
        bs = [r for r in col.fs_records if r.sj_div == "BS"]
        self.assertEqual({r.period_key for r in bs}, {"FY2025"})           # 열 4 제외
        self.assertTrue(all(r.unit == 1000 for r in bs))
        self.assertEqual(col.units[("OFS", "BS")], 1000)
        is_ = [r for r in col.fs_records if r.sj_div == "IS"]
        self.assertEqual({r.period_key for r in is_}, {"FY2025", "FY2023"})
        self.assertEqual({r.role for r in is_ if r.period_key == "FY2023"}, {"prev2"})
        self.assertFalse([r for r in col.fs_records if r.sj_div == "CF"])
        self.assertIn("파일 layout.xlsx 시트 현금흐름표 표 1: 표 인식 못 함(기타)", col.logs)
        self.assertEqual(col.periods, ["FY2023", "FY2024", "FY2025"])        # FY2024 는 자본변동표 전기 행에서

    def test_fiscal_only_columns_inferred_and_flagged(self):
        p = _wb(os.path.join(OUT, "fiscal_only.xlsx"), {"재무상태표": [
            ["재무상태표"], ["(단위: 천원)"], ["과목", "제 11 기", "제 10 기"],
            ["자산총계", 1500, 1340], ["부채총계", 600, 640], ["자본총계", 900, 700]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        got = {(r.period_key, r.role) for r in col.fs_records}
        self.assertEqual(got, {("FY2025", "cur"), ("FY2024", "prev")})
        self.assertEqual(col.plan.inferred, ["fiscal_only.xlsx|재무상태표|0"])
        self.assertTrue(any("기간 역할을 회차로 추정 — 화면에서 확인" in m for m in col.logs))
        self.assertEqual((col.units[("fiscal", "FY2025")], col.units[("OFS", "BS")]), (11, 1000))

    def test_file_unit_when_sheet_has_none(self):
        p = _wb(os.path.join(OUT, "nounit.xlsx"), {"BS": [
            ["과목", "당기", "전기"], ["자산총계", 1500, 1340], ["부채총계", 600, 640], ["자본총계", 900, 700]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31", unit=1000000)]), M.load(), None)
        self.assertTrue(col.fs_records and all(r.unit == 1000000 for r in col.fs_records))
        self.assertIn("파일 nounit.xlsx 시트 BS 표 1: BS FY2025(당기)·FY2024(전기) 단위 백만원", col.logs)
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        self.assertTrue(all(r.unit == 1 for r in col.fs_records))
        self.assertIn("파일 nounit.xlsx 시트 BS 표 1: BS FY2025(당기)·FY2024(전기) 단위 미표기", col.logs)

    def test_half_file(self):
        p = _wb(os.path.join(OUT, "half.xlsx"), {
            "재무상태표": [["재무상태표"], ["(단위: 원)"], ["과목", "당반기말", "전기말"],
                           ["자산총계", 1600, 1500], ["부채총계", 650, 600], ["자본총계", 950, 900]],
            "손익계산서": [["손익계산서"], ["(단위: 원)"], ["과목", "당반기", "전반기"],
                           ["매출액", 1100, 1000], ["영업이익", 150, 130], ["반기순이익", 110, 100]],
            "주석": [["27. 판매비와관리비"], ["(단위: 원)"], ["구분", "제 12 기", "제 11 기"], ["급여", 80, 70], ["합계", 80, 70]]})
        st = F.build_statements(os.path.join(OUT, "half_base.xlsx"))
        cfg = _cfg([_file(st, "2025-12-31"), _file(p, "2026-06-30", kind="half")], include_half=True, end_year=2026)
        col = FC.collect_files(cfg, M.load(), None)
        self.assertEqual(col.periods, ["FY2024", "FY2025", "2025H1", "2026H1"])
        bs = {(r.period_key, r.role, r.rcept_no) for r in col.fs_records if r.sj_div == "BS" and r.label == "자산총계"}
        self.assertIn(("2026H1", "cur", "FILE:half.xlsx"), bs)
        # 연간 키의 출처는 연간 파일뿐(DART 와 같음): 반기 BS 전기말 열은 FY2025 를 덮어쓰지 않는다
        self.assertEqual(col.plan.sources["FY2025"], [("FILE:half_base.xlsx", "cur")])
        self.assertEqual(col.plan.sources["2026H1"], [("FILE:half.xlsx", "cur")])
        self.assertEqual(col.plan.sources["2025H1"], [("FILE:half.xlsx", "prev")])
        fy25 = [r for r in col.fs_records if r.sj_div == "BS" and r.period_key == "FY2025"]
        self.assertEqual(len(fy25), len(F.BS_ROWS))                               # 연간 파일의 전체 행
        self.assertTrue(all(r.rcept_no == "FILE:half_base.xlsx" for r in fy25))
        self.assertEqual(next(r.amount for r in fy25 if r.label == "자산총계"), 1500 * F.M)
        is_ = {(r.period_key, r.role) for r in col.fs_records if r.sj_div == "IS" and r.rcept_no == "FILE:half.xlsx"}
        self.assertEqual(is_, {("2026H1", "cur"), ("2025H1", "prev")})
        self.assertFalse([r for r in col.note_records if r.rcept_no == "FILE:half.xlsx"])
        self.assertIn("파일 half.xlsx 시트 주석 표 1: 반기 파일의 주석 표는 쓰지 않음(반기는 재무제표만)", col.logs)
        self.assertEqual(col.plan.inferred, [])                                    # 반기 주석의 '제 N 기' 열은 해석하지 않는다
        self.assertNotIn(("fiscal", "2025H1"), col.units)
        self.assertEqual(col.plan.base_report.rcept_no, "FILE:half_base.xlsx")
        self.assertNotIn("2026H1", col.plan.note_sources)

    def test_half_file_alone_has_no_annual_period(self):
        p = _wb(os.path.join(OUT, "half_only.xlsx"), {"재무상태표": [
            ["재무상태표"], ["(단위: 원)"], ["과목", "당반기말", "전기말"],
            ["자산총계", 1600, 1500], ["부채총계", 650, 600], ["자본총계", 950, 900]]})
        col = FC.collect_files(_cfg([_file(p, "2026-06-30", kind="half")], include_half=True, end_year=2026), M.load(), None)
        self.assertEqual(col.periods, ["2026H1"])                                   # 전기말 열만으로 FY2025 를 만들지 않는다
        self.assertEqual(list(col.plan.sources), ["2026H1"])
        self.assertEqual({r.period_key for r in col.fs_records}, {"2026H1"})
        self.assertIsNone(col.plan.base_report)

    def test_fiscal_only_column_with_date_anchor(self):
        p = _wb(os.path.join(OUT, "half_anchor.xlsx"), {"재무상태표": [
            ["재무상태표"], ["(단위: 원)"], ["과목", "2026.06.30 제 12 기", "제 11 기"],
            ["자산총계", 1600, 1500], ["부채총계", 650, 600], ["자본총계", 950, 900]]})
        col = FC.collect_files(_cfg([_file(p, "2026-06-30", kind="half")], include_half=True, end_year=2026), M.load(), None)
        # 반기 BS: 날짜 열이 기준 → 당기 열은 2026H1, 회차 −1 열은 직전 연차 기말(FY2025, 반기 파일이라 FY 열은 만들지 않음)
        self.assertIn("파일 half_anchor.xlsx 시트 재무상태표 표 1: BS 2026H1(당기)·FY2025(전기)(반기 파일 — FY 출처 아님) 단위 원",
                      col.logs)
        self.assertEqual({(r.period_key, r.role) for r in col.fs_records}, {("2026H1", "cur")})
        self.assertEqual(col.plan.inferred, ["half_anchor.xlsx|재무상태표|0"])
        self.assertIn("파일 half_anchor.xlsx 시트 재무상태표 표 1: 기간 역할을 회차로 추정 — 화면에서 확인", col.logs)

    def test_half_file_does_not_overwrite_annual_fiscal(self):
        st = F.build_statements(os.path.join(OUT, "fiscal_base.xlsx"))               # FY2025 = 제 12 기
        p = _wb(os.path.join(OUT, "half_fiscal.xlsx"), {"재무상태표": [
            ["재무상태표"], ["(단위: 원)"], ["과목", "2026.06.30 제 12 기", "제 11 기"],
            ["자산총계", 1600, 1500], ["부채총계", 650, 600], ["자본총계", 950, 900]]})
        cfg = _cfg([_file(st, "2025-12-31"), _file(p, "2026-06-30", kind="half")], include_half=True, end_year=2026)
        col = FC.collect_files(cfg, M.load(), None)
        self.assertEqual(col.units[("fiscal", "FY2025")], 12)                       # 반기 파일의 '제 11 기' 가 덮지 않음
        self.assertEqual(col.units[("fiscal", "2026H1")], 12)

    def test_file_meta_sheets_and_unit(self):
        st = F.build_statements(os.path.join(OUT, "meta.xlsx"))
        col = FC.collect_files(_cfg([_file(st, "2025-12-31")]), M.load(), None)
        meta = col.plan.file_meta["FILE:meta.xlsx"]
        self.assertIn("재무상태표(표 1)", meta["sheets"])
        self.assertEqual(meta["unit"], "원")

    def test_toc_keyed_by_title_keeps_latest_number(self):
        old = _wb(os.path.join(OUT, "notes_old_no.xlsx"), {"주석": [
            ["26. 판매비와관리비"], ["(단위: 원)"], ["구분", "당기", "전기"], ["급여", 140, 130], ["합계", 140, 130]]})
        new = F.build_notes(os.path.join(OUT, "notes_new_no.xlsx"))                 # '27. 판매비와관리비'
        col = FC.collect_files(_cfg([_file(old, "2024-12-31"), _file(new, "2025-12-31")]), M.load(), None)
        sga = [e for e in col.toc if e.title_norm == "판매비와관리비"]
        self.assertEqual([(e.no, e.title, e.table_count) for e in sga], [("27", "판매비와관리비", 2)])
        self.assertEqual({r.note_no for r in col.note_records if r.topic == "판관비"}, {"27"})   # FY2024 도 2025 파일 전기 열

    def test_duplicate_basename_layout_key(self):
        os.makedirs(os.path.join(OUT, "dupA"), exist_ok=True); os.makedirs(os.path.join(OUT, "dupB"), exist_ok=True)
        a = F.build_statements(os.path.join(OUT, "dupA", "same.xlsx"))
        b = F.build_statements(os.path.join(OUT, "dupB", "same.xlsx"), year=2024)
        layout = {"same.xlsx#2|재무상태표|0": {"unit": 1000}}
        col = FC.collect_files(_cfg([_file(a, "2025-12-31"), _file(b, "2024-12-31")], layout), M.load(), None)
        self.assertEqual(set(col.plan.reports), {"FILE:same.xlsx", "FILE:same.xlsx#2"})
        self.assertIn("파일 same.xlsx 시트 재무상태표 표 1: BS FY2024(당기)·FY2023(전기) 단위 천원", col.logs)
        self.assertIn("파일 same.xlsx 시트 재무상태표 표 1: BS FY2025(당기)·FY2024(전기) 단위 원", col.logs)

    def test_note_without_period_marks_becomes_raw(self):
        p = _wb(os.path.join(OUT, "rawnote.xlsx"), {"주석": [
            ["12. 유형자산"], ["(1) 유형자산의 내역은 다음과 같습니다."], ["(단위: 원)"],
            ["구분", "취득원가", "감가상각누계액", "장부금액"],
            ["토지", 300, "-", 300], ["건물", 400, -80, 320], ["합계", 700, -80, 620]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        raw = [r for r in col.note_records if r.raw_cell is not None]
        self.assertTrue(raw and len(raw) == len(col.note_records))
        self.assertEqual({r.period_key for r in raw}, {"FY2025"})
        self.assertEqual(raw[0].raw_cell, (0, 0))
        self.assertEqual({r.col_label for r in raw}, {"구분", "취득원가", "감가상각누계액", "장부금액"})
        self.assertEqual(next(r.value for r in raw if r.row_label == "합계" and r.col_label == "장부금액"), 620)
        self.assertTrue(any("기간 표시 없음 → rN_기타" in m for m in col.logs))
        self.assertEqual([(e.no, e.title) for e in col.toc], [("12", "유형자산")])

    def test_hidden_sheet_and_orphan_rows_logged(self):
        p = _wb(os.path.join(OUT, "hidden.xlsx"), {
            "재무상태표": [["재무상태표"], ["(단위: 원)"], ["과목", "당기", "전기"],
                           ["자산총계", 1500, 1340], ["부채총계", 600, 640], ["자본총계", 900, 700],
                           [], [], ["참고 합계", 3000]],
            "메모": [["메모", 1]]})
        from openpyxl import load_workbook
        wb = load_workbook(p); wb["메모"].sheet_state = "hidden"; wb.save(p)
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        self.assertIn("파일 hidden.xlsx 시트 메모: 숨김 시트 → 건너뜀", col.logs)
        self.assertIn("파일 hidden.xlsx 시트 재무상태표: 표에 속하지 않은 숫자 행 1개(행 9)", col.logs)
        self.assertEqual(len([r for r in col.fs_records if r.period_key == "FY2025"]), 3)

    def test_note_include_filter(self):
        nt = F.build_notes(os.path.join(OUT, "notes_inc.xlsx"))
        col = FC.collect_files(_cfg([_file(nt, "2025-12-31")], note_include={"판매비와관리비": True}), M.load(), None)
        self.assertEqual({r.topic for r in col.note_records}, {"판관비"})
        self.assertEqual([(e.no, e.include) for e in col.toc], [("12", False), ("27", True)])

    def test_half_year_sub_columns_use_cumulative_and_expose_all(self):
        """반기 손익 '당반기 | 전반기' 아래 '3개월 | 누적 | 3개월 | 누적' → 누적 열이 기간 열. 화면(auto_periods)엔 네 열 전부."""
        p = _wb(os.path.join(OUT, "half_sub.xlsx"), {"손익계산서": [
            ["손익계산서"], ["(단위: 원)"], ["과목", "당반기", None, "전반기", None], [None, "3개월", "누적", "3개월", "누적"],
            ["매출액", 50, 100, 40, 80], ["영업이익", 10, 20, 8, 16], ["반기순이익", 5, 10, 4, 8]]})
        col = FC.collect_files(_cfg([_file(p, "2026-06-30", kind="half")], include_half=True, end_year=2026), M.load(), None)
        sales = {r.period_key: r.amount for r in col.fs_records if r.label == "매출액"}
        self.assertEqual(sales, {"2026H1": 100, "2025H1": 80})
        pv = FC.preview(p, period_end="2026-06-30", kind="half")
        t = pv["scan"]["sheets"][0]["tables"][0]
        self.assertEqual(t["auto_periods"], {"2": None, "3": "2026H1", "4": None, "5": "2025H1"})
        self.assertEqual(t["col_texts"]["2"], "당반기 3개월")
        # 사용자가 3개월 열로 바꿀 수 있다(값 열이면 어느 열이든 기간을 매길 수 있음)
        layout = {"half_sub.xlsx|손익계산서|0": {"periods": {"2": "2026H1", "3": None}}}
        col = FC.collect_files(_cfg([_file(p, "2026-06-30", kind="half")], layout, include_half=True, end_year=2026), M.load(), None)
        self.assertEqual({r.period_key: r.amount for r in col.fs_records if r.label == "매출액"}, {"2026H1": 50, "2025H1": 80})

    def test_label_cols_override_keeps_two_row_header_periods(self):
        """2행 머리글(당기·전기 / 금액·구성비) 표에 라벨 열만 고쳐도 기간 열이 남아 레코드가 나온다."""
        p = _wb(os.path.join(OUT, "lc_2row.xlsx"), {"손익계산서": [
            ["손익계산서"], ["(단위: 원)"], ["코드", "과목", "당기", None, "전기", None],
            [None, None, "금액", "구성비", "금액", "구성비"],
            ["A01", "매출액", 1000, 1.0, 900, 1.0], ["A02", "영업이익", 400, 0.4, 400, 0.44]]})
        layout = {"lc_2row.xlsx|손익계산서|0": {"label_cols": [2]}}
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")], layout), M.load(), None)
        self.assertEqual({(r.period_key, r.amount) for r in col.fs_records if r.label == "매출액"}, {("FY2025", 1000), ("FY2024", 900)})
        pv = FC.preview(p, period_end="2025-12-31", kind="annual", layout=layout)
        t = pv["scan"]["sheets"][0]["tables"][0]
        self.assertEqual((t["auto_periods"]["3"], t["auto_periods"]["5"]), ("FY2025", "FY2024"))

    def test_two_columns_same_period_keep_user_mapped(self):
        """사용자가 3개월 열을 2025H1 로 매기고 누적 열도 자동으로 2025H1 → 사용자 열 하나만 쓰고 누적 열은 제외(로그)."""
        p = _wb(os.path.join(OUT, "half_dup.xlsx"), {"손익계산서": [
            ["손익계산서"], ["(단위: 원)"], ["과목", "당반기", None, "전반기", None], [None, "3개월", "누적", "3개월", "누적"],
            ["매출액", 50, 100, 40, 80], ["영업이익", 10, 20, 8, 16], ["반기순이익", 5, 10, 4, 8]]})
        layout = {"half_dup.xlsx|손익계산서|0": {"periods": {"2": "2025H1"}}}
        col = FC.collect_files(_cfg([_file(p, "2025-06-30", kind="half")], layout, include_half=True), M.load(), None)
        sales = [(r.period_key, r.amount) for r in col.fs_records if r.label == "매출액"]
        self.assertEqual([a for k, a in sales if k == "2025H1"], [50])
        self.assertIn("파일 half_dup.xlsx 시트 손익계산서 표 1: 열 3 은 기간 2025H1 와 겹쳐 제외", col.logs)

    def test_sce_prefix_switches_only_on_boundary_words(self):
        """'전기 기초잔액 … 전기 기말잔액 / 당기 기초잔액 … 당기 기말잔액': 구간 안의 '당기순이익'·'전기오류수정이익'은 구간을 바꾸지 않는다."""
        p = _wb(os.path.join(OUT, "sce_prefix.xlsx"), {"자본변동표": [
            ["자본변동표"], ["(단위: 원)"], ["구분", "자본금", "이익잉여금", "총계"],
            ["전기 기초잔액", 500, 80, 580], ["당기순이익", "-", 120, 120], ["배당금", "-", -10, -10],
            ["전기 기말잔액", 500, 190, 690], ["당기 기초잔액", 500, 190, 690], ["전기오류수정이익", "-", 5, 5],
            ["당기순이익", "-", 200, 200], ["당기 기말잔액", 500, 395, 895]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        tot = [(r.period_key, r.label) for r in col.fs_records if r.detail == "총계"]
        self.assertEqual(tot, [("FY2024", "전기 기초잔액"), ("FY2024", "당기순이익"), ("FY2024", "배당금"), ("FY2024", "전기 기말잔액"),
                               ("FY2025", "당기 기초잔액"), ("FY2025", "전기오류수정이익"), ("FY2025", "당기순이익"),
                               ("FY2025", "당기 기말잔액")])

    def test_sce_mixed_boundary_styles(self):
        """한 표에 '전기초 … 전기말' 과 '당기 기초잔액 … 당기 기말잔액' 이 섞여도 앞 구간 FY2024, 뒤 구간 FY2025."""
        p = _wb(os.path.join(OUT, "sce_mixed.xlsx"), {"자본변동표": [
            ["자본변동표"], ["(단위: 원)"], ["구분", "자본금", "이익잉여금", "총계"],
            ["전기초", 500, 80, 580], ["당기순이익", "-", 120, 120], ["전기말", 500, 200, 700],
            ["당기 기초잔액", 500, 200, 700], ["당기순이익", "-", 200, 200], ["당기 기말잔액", 500, 400, 900]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        tot = [(r.period_key, r.label) for r in col.fs_records if r.detail == "총계"]
        self.assertEqual(tot, [("FY2024", "전기초"), ("FY2024", "당기순이익"), ("FY2024", "전기말"),
                               ("FY2025", "당기 기초잔액"), ("FY2025", "당기순이익"), ("FY2025", "당기 기말잔액")])

    def test_from_to_header_end_date_and_guess(self):
        """'제 12 기 2025년 1월 1일부터 2025년 12월 31일까지': 기말 = 마지막 날짜. 결산일 추정도 12-31."""
        p = _wb(os.path.join(OUT, "fromto.xlsx"), {"손익계산서": [
            ["손익계산서"], ["(단위: 원)"],
            ["과목", "제 12 기 2025년 1월 1일부터 2025년 12월 31일까지", "제 11 기 2024년 1월 1일부터 2024년 12월 31일까지"],
            ["매출액", 100, 90], ["영업이익", 40, 30], ["당기순이익", 30, 20]]})
        pv = FC.preview(p)
        self.assertEqual((pv["period_end_guess"], pv["kind_guess"]), ("2025-12-31", "annual"))
        self.assertEqual(pv["scan"]["sheets"][0]["tables"][0]["auto_periods"], {"2": "FY2025", "3": "FY2024"})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        self.assertEqual({(r.period_key, r.role) for r in col.fs_records}, {("FY2025", "cur"), ("FY2024", "prev")})
        self.assertEqual(col.plan.inferred, [])
        self.assertEqual((col.units[("fiscal", "FY2025")], col.units[("fiscal", "FY2024")]), (12, 11))

    def test_non_period_end_date_falls_back_to_fiscal(self):
        """날짜가 6월·12월 말이 아니어도 회차가 있으면 제외하지 않고 회차로 역할을 추정한다(확인 필요 표시)."""
        p = _wb(os.path.join(OUT, "oddate.xlsx"), {"재무상태표": [
            ["재무상태표"], ["(단위: 원)"], ["과목", "제 12 기 (2025.01.01)", "제 11 기 (2024.01.01)"],
            ["자산총계", 1500, 1340], ["부채총계", 600, 640], ["자본총계", 900, 700]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        self.assertEqual({(r.period_key, r.role) for r in col.fs_records}, {("FY2025", "cur"), ("FY2024", "prev")})
        self.assertEqual(col.plan.inferred, ["oddate.xlsx|재무상태표|0"])
        self.assertFalse(any("제외" in m for m in col.logs))

    def test_sce_rows_without_year_split_by_prefix(self):
        """연도 없는 자본변동표 행(전기초/당기순이익/전기말/당기초/당기순이익/당기말): 경계 행으로 전기·당기 구간을 나눈다."""
        p = _wb(os.path.join(OUT, "sce_noyear.xlsx"), {"자본변동표": [
            ["자본변동표"], ["(단위: 원)"], ["구분", "자본금", "이익잉여금", "총계"],
            ["전기초", 500, 80, 580], ["당기순이익", "-", 120, 120], ["전기말", 500, 200, 700],
            ["당기초", 500, 200, 700], ["당기순이익", "-", 200, 200], ["당기말", 500, 400, 900]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        tot = [(r.period_key, r.role, r.label_norm, r.amount) for r in col.fs_records if r.detail == "총계"]
        self.assertEqual(tot, [("FY2024", "prev", "기초", 580), ("FY2024", "prev", "당기순이익", 120), ("FY2024", "prev", "기말", 700),
                               ("FY2025", "cur", "기초", 700), ("FY2025", "cur", "당기순이익", 200), ("FY2025", "cur", "기말", 900)])
        self.assertEqual(col.periods, ["FY2024", "FY2025"])
        self.assertFalse(any("전부 당기" in m for m in col.logs))
        # 아무 표기도 없으면 전부 당기 + 로그
        q = _wb(os.path.join(OUT, "sce_nomark.xlsx"), {"자본변동표": [
            ["자본변동표"], ["(단위: 원)"], ["구분", "자본금", "이익잉여금", "총계"],
            ["기초", 500, 200, 700], ["순이익", "-", 200, 200], ["기말", 500, 400, 900]]})
        col = FC.collect_files(_cfg([_file(q, "2025-12-31")]), M.load(), None)
        self.assertEqual({(r.period_key, r.role) for r in col.fs_records}, {("FY2025", "cur")})
        self.assertTrue(any("전부 당기로" in m for m in col.logs))

    def test_formula_only_file_logged(self):
        """수식만 있는 파일(계산값 없음): 표가 없어 조용히 비는 대신 로그에 이유를 남긴다. preview 에도 note."""
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "재무상태표"
        ws.append(["재무상태표"]); ws.append(["(단위: 원)"]); ws.append(["과목", "당기", "전기"])
        for lab, a, b in (("자산총계", 1500, 1340), ("부채총계", 600, 640), ("자본총계", 900, 700)):
            ws.append([lab, f"={a}", f"={b}"])
        p = os.path.join(OUT, "formula_only.xlsx")
        wb.save(p)
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        self.assertEqual(col.fs_records, [])
        self.assertIn(f"파일 formula_only.xlsx 시트 재무상태표: {FC.XI.FORMULA_NOTE}", col.logs)
        self.assertIn("파일 formula_only.xlsx: 인식된 표 없음(수식만 있는 파일이면 Excel 에서 한 번 저장한 뒤 다시 넣으세요)", col.logs)
        pv = FC.preview(p)
        self.assertEqual(pv["note"], FC.NO_TABLE_NOTE)
        self.assertEqual(pv["scan"]["sheets"][0]["note"], FC.XI.FORMULA_NOTE)
        self.assertIsNone(FC.preview(F.build_statements(os.path.join(OUT, "has_tables.xlsx")))["note"])

    def test_layout_header_row_and_label_cols_rebuild(self):
        """머리글 행·라벨 열 수정값은 표를 그 행·열로 다시 인식한다 — 기간 열·값 열이 수정값과 맞게."""
        p = _wb(os.path.join(OUT, "rebuild.xlsx"), {"재무상태표": [
            ["재무상태표"], ["(단위: 천원)"], ["코드", "과목", "당기", "전기"], [None, None, "제 12 기", "제 11 기"],
            ["A01", "자산총계", 1500, 1340], ["B01", "부채총계", 600, 640], ["C01", "자본총계", 900, 700]]})
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")]), M.load(), None)
        self.assertEqual(col.plan.inferred, [])
        self.assertEqual({r.label for r in col.fs_records}, {"A01 / 자산총계", "B01 / 부채총계", "C01 / 자본총계"})
        layout = {"rebuild.xlsx|재무상태표|0": {"header_row": 4, "label_cols": [2]}}
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")], layout), M.load(), None)
        self.assertEqual({r.label for r in col.fs_records}, {"자산총계", "부채총계", "자본총계"})
        self.assertEqual({(r.period_key, r.role, r.unit) for r in col.fs_records}, {("FY2025", "cur", 1000), ("FY2024", "prev", 1000)})
        # 머리글 행 4 = 자동 인식 값(2행 머리글의 아래 행) → 머리글 첫 행(3행 '당기|전기')부터 다시 인식해 역할 유지, 추정 아님
        self.assertEqual(col.plan.inferred, [])
        pv = FC.preview(p, period_end="2025-12-31", layout=layout)
        t = pv["scan"]["sheets"][0]["tables"][0]
        self.assertEqual((t["header_row"], t["first_row"], t["label_cols"], t["value_cols"]), (4, 5, [2], [3, 4]))
        self.assertEqual({c: g["role"] for c, g in t["period_cols"].items()}, {"3": "cur", "4": "prev"})
        self.assertFalse(t["inferred"])
        # 지정한 행부터 머리글이 2행이면(3행 '당기|전기' + 4행 '제 N 기') 본문은 그 아래부터 — 머리글 행이 본문으로 새지 않는다
        col = FC.collect_files(_cfg([_file(p, "2025-12-31")], {"rebuild.xlsx|재무상태표|0": {"header_row": 3}}), M.load(), None)
        self.assertEqual({r.label for r in col.fs_records}, {"A01 / 자산총계", "B01 / 부채총계", "C01 / 자본총계"})
        self.assertEqual({(r.period_key, r.role) for r in col.fs_records}, {("FY2025", "cur"), ("FY2024", "prev")})

    def test_missing_period_end_message(self):
        st = F.build_statements(os.path.join(OUT, "nope.xlsx"))
        for pe in (None, "", "2025-13-99"):
            with self.assertRaises(ValueError) as cm:
                FC.collect_files(_cfg([{"path": st, "period_end": pe, "kind": "annual"}]), M.load(), None)
            self.assertEqual(str(cm.exception), "파일 nope.xlsx: 결산일이 없습니다")


if __name__ == "__main__":
    unittest.main()
