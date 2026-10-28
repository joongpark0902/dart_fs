"""core.xlsx_input: 회사 제시 엑셀 재무제표 표 인식(가상 '샘플전자' 픽스처)."""
import json
import unittest
from core import xlsx_input as X
from tests.fixtures import xlsx_fixture as F


def _sheet(scan, name):
    return next(s for s in scan.sheets if s.name == name)


class TestStatements(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.scan = X.scan(F.build_statements())

    def test_bs(self):
        ts = _sheet(self.scan, "재무상태표").tables
        self.assertEqual(len(ts), 1)
        t = ts[0]
        self.assertEqual(t.kind, "BS")
        self.assertEqual(t.title, "재무상태표")
        self.assertEqual((t.header_row, t.first_row, t.last_row), (5, 6, 5 + len(F.BS_ROWS)))
        self.assertEqual(t.label_cols, [1])
        self.assertEqual(t.value_cols, [3, 4])           # 주석 열(B)은 값 열이 아니다
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {3: "cur", 4: "prev"})
        self.assertEqual(t.period_cols[3].end_date, "2025-12-31")
        self.assertEqual(t.period_cols[4].end_date, "2024-12-31")
        self.assertEqual((t.period_cols[3].fiscal, t.period_cols[4].fiscal), (12, 11))
        self.assertEqual(t.unit, 1)
        self.assertEqual(t.unit_text, "(단위: 원)")
        self.assertEqual(t.confidence, "high")
        self.assertIn("자산총계", t.reason)

    def test_is(self):
        t, = _sheet(self.scan, "손익계산서").tables
        self.assertEqual(t.kind, "IS")
        self.assertEqual((t.header_row, t.first_row), (5, 6))
        self.assertEqual(t.label_cols, [1])
        self.assertEqual(t.value_cols, [2, 3])
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {2: "cur", 3: "prev"})
        self.assertEqual(t.unit, 1)
        self.assertEqual(t.confidence, "high")

    def test_cf_two_label_cols(self):
        t, = _sheet(self.scan, "현금흐름표").tables
        self.assertEqual(t.kind, "CF")
        self.assertEqual(t.header_row, 5)
        self.assertEqual(t.label_cols, [1, 2])
        self.assertEqual(t.value_cols, [3, 4])
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {3: "cur", 4: "prev"})
        self.assertEqual(t.last_row, 5 + len(F.CF_ROWS))

    def test_sce_component_header(self):
        t, = _sheet(self.scan, "자본변동표").tables
        self.assertEqual(t.kind, "SCE")
        self.assertEqual((t.header_row, t.first_row, t.last_row), (4, 5, 4 + len(F.SCE_ROWS)))
        self.assertEqual(t.label_cols, [1])
        self.assertEqual(t.value_cols, [2, 3, 4])
        self.assertEqual(t.period_cols, {})
        self.assertEqual(t.unit, 1)

    def test_to_dict_json(self):
        d = self.scan.to_dict()
        s = json.dumps(d, ensure_ascii=False)
        self.assertIn("재무상태표", s)
        self.assertEqual(d["sheets"][0]["tables"][0]["kind"], "BS")
        self.assertEqual(d["sheets"][0]["tables"][0]["period_cols"]["3"]["role"], "cur")

    def test_max_sheets(self):
        self.assertEqual(len(X.scan(F.build_statements(), max_sheets=2).sheets), 2)


class TestIsCisSameSheet(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ts = _sheet(X.scan(F.build_is_cis()), "손익계산서").tables

    def test_two_tables(self):
        self.assertEqual([t.kind for t in self.ts], ["IS", "CIS"])
        self.assertEqual([t.idx for t in self.ts], [0, 1])

    def test_is_two_row_header_amount_ratio(self):
        t = self.ts[0]
        self.assertEqual((t.header_row, t.first_row, t.last_row), (4, 5, 14))
        self.assertEqual(t.value_cols, [2, 3, 4, 5])
        # 같은 기간의 금액·구성비 2열 → 기간은 첫 숫자 열만
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {2: "cur", 4: "prev"})
        self.assertEqual(t.unit, 1)

    def test_cis(self):
        t = self.ts[1]
        self.assertEqual(t.title, "포괄손익계산서")
        self.assertEqual((t.header_row, t.first_row, t.last_row), (19, 20, 23))
        self.assertEqual(t.value_cols, [2, 3])
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {2: "cur", 3: "prev"})


class TestNotes(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ts = _sheet(X.scan(F.build_notes()), "주석").tables

    def test_sga(self):
        t = self.ts[0]
        self.assertEqual(t.kind, "NOTE")
        self.assertEqual(t.title, "27. 판매비와관리비")
        self.assertEqual((t.header_row, t.first_row, t.last_row), (3, 4, 8))
        self.assertEqual(t.label_cols, [1])
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {2: "cur", 3: "prev"})
        self.assertEqual(t.unit, 1)
        self.assertEqual(t.confidence, "high")

    def test_ppe_movement(self):
        t = self.ts[1]
        self.assertEqual(t.kind, "NOTE")
        self.assertEqual(t.title, "12. 유형자산")
        self.assertEqual((t.header_row, t.first_row, t.last_row), (14, 15, 18))
        self.assertEqual(t.label_cols, [1])
        self.assertEqual(t.value_cols, [2, 3, 4, 5])      # '토지'(대부분 '-')도 숫자 열
        self.assertEqual(t.period_cols, {})
        self.assertEqual(t.unit, 1)


class TestTrialBalance(unittest.TestCase):
    def test_other(self):
        t, = _sheet(X.scan(F.build_trial_balance()), "시산표").tables
        self.assertEqual(t.kind, "OTHER")
        self.assertEqual(t.header_row, 3)
        self.assertEqual(t.label_cols, [2])
        self.assertEqual(t.value_cols, [3, 4])
        self.assertIsNone(t.unit)
        self.assertEqual(t.confidence, "low")


class TestHelpers(unittest.TestCase):
    def test_unit_of(self):
        self.assertEqual(X.unit_of("(단위 : 천원)"), 1000)
        self.assertEqual(X.unit_of("단위: 백만원"), 10**6)
        self.assertEqual(X.unit_of("(천원)"), 1000)
        self.assertIsNone(X.unit_of("매출액"))
        from core import normalize as N
        self.assertEqual(N.unit_multiplier("백만 원"), 10**6)
        self.assertIsNone(N.unit_multiplier("주"))

    def test_period_guess(self):
        g = X.period_guess("제 12(당) 기")
        self.assertEqual((g.role, g.fiscal, g.end_date), ("cur", 12, None))
        self.assertEqual(X.period_guess("2025-12-31").end_date, "2025-12-31")
        self.assertEqual(X.period_guess("FY2025").end_date, "2025")
        self.assertEqual(X.period_guess("2025년").end_date, "2025")
        self.assertEqual(X.period_guess("전기말").role, "prev")
        self.assertEqual(X.period_guess("당반기").role, "cur")
        self.assertIsNone(X.period_guess("매출액"))
        self.assertIsNone(X.period_guess("(1) 당기 중 유형자산의 변동내역은 다음과 같습니다."))

    def test_period_guess_from_to_dates_use_last_date(self):
        """'…부터 …까지' 머리글: 기말은 마지막 날짜. 날짜를 뗀 뒤 길이 제한·역할·회차를 보므로 긴 DART 식 머리글도 읽는다."""
        g = X.period_guess("제 12 기 2025년 1월 1일부터 2025년 12월 31일까지")
        self.assertEqual((g.role, g.fiscal, g.end_date), (None, 12, "2025-12-31"))
        g = X.period_guess("제 12(당) 기 2025년 1월 1일부터 2025년 12월 31일까지")
        self.assertEqual((g.role, g.fiscal, g.end_date), ("cur", 12, "2025-12-31"))
        g = X.period_guess("제 11(전) 기 2024.01.01 ~ 2024.12.31")
        self.assertEqual((g.role, g.fiscal, g.end_date), ("prev", 11, "2024-12-31"))
        self.assertEqual(X.period_guess("2025년 12월 31일 현재").end_date, "2025-12-31")
        self.assertIsNone(X.period_guess("2025년 1월 1일부터 2025년 12월 31일까지의 회계기간에 대한 재무제표의 주석입니다"))



class TestScanRows(unittest.TestCase):
    """시트 값 목록을 직접 넣는 변형 머리글."""

    def test_role_above_fiscal_two_rows(self):
        rows = [["(단위 : 천원)"], ["과목", "당기", "전기"], [None, "제 12 기", "제 11 기"],
                ["자산총계", 100, 90], ["부채총계", 40, 30], ["자본총계", 60, 60]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.kind, t.header_row, t.first_row, t.unit), ("BS", 3, 4, 1000))
        self.assertEqual({c: (g.role, g.fiscal) for c, g in t.period_cols.items()},
                         {2: ("cur", 12), 3: ("prev", 11)})

    def test_fiscal_roles_from_anchor_only(self):
        """역할+회차가 함께 적힌 열('제 12(당) 기')이 있으면 나머지 '제 N 기' 열의 역할을 회차 차이로 채운다."""
        rows = [["과목", "제 12(당) 기", "제 11 기", "제 10 기"],
                ["자산총계", 100, 90, 80], ["부채총계", 40, 30, 20], ["자본총계", 60, 60, 60]]
        t, = X.scan_rows(rows)
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {2: "cur", 3: "prev", 4: "prev2"})

    def test_fiscal_only_no_anchor_role_none(self):
        """'제 11 기'·'제 10 기' 만 있으면 최대 회차를 당기로 단정하지 않는다(상위에서 결산일·기수로 정함)."""
        rows = [["과목", "제 11 기", "제 10 기"],
                ["자산총계", 100, 90], ["부채총계", 40, 30], ["자본총계", 60, 60]]
        t, = X.scan_rows(rows)
        self.assertEqual({c: (g.role, g.fiscal) for c, g in t.period_cols.items()},
                         {2: (None, 11), 3: (None, 10)})

    def test_text_row_after_blank_not_in_body(self):
        rows = [["27. 판매비와관리비"], ["(단위: 원)"], ["구분", "당기", "전기"],
                ["급여", 150, 140], ["감가상각비", 60, 70], ["합계", 210, 210], [],
                ["상기 금액은 모두 비용으로 인식하였습니다."], ["구분", "토지", "건물"]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.first_row, t.last_row), (4, 6))

    def test_bs_section_row_after_blank_continues(self):
        """자산 → 빈 행 → '부채' 구분 행 → 부채·자본: 하나의 표(빈 행 뒤 구분 행에서 끝내면 부채·자본이 사라짐)."""
        rows = [["재무상태표"], ["(단위: 원)"], ["과목", "당기", "전기"],          # 1~3
                ["자산", None, None], ["유동자산", 600, 500], ["비유동자산", 900, 840],  # 4~6
                ["자산총계", 1500, 1340], [],                                        # 7~8
                ["부채", None, None], ["부채총계", 600, 640], ["자본", None, None],    # 9~11
                ["자본금", 500, 500], ["이익잉여금", 400, 200], ["자본총계", 900, 700],  # 12~14
                ["부채와자본총계", 1500, 1340]]                                       # 15
        ts = X.scan_rows(rows)
        self.assertEqual(len(ts), 1)
        t = ts[0]
        self.assertEqual((t.kind, t.first_row, t.last_row, t.confidence), ("BS", 4, 15, "high"))

    def test_numbered_section_row_after_blank_continues(self):
        """빈 행 뒤 '2. 비유동자산'(번호 붙은 구분 행, 숫자 없음)은 주석 제목이 아니다 — 표를 끊지 않는다."""
        rows = [["과목", "당기", "전기"], ["자산", None, None], ["1. 유동자산", None, None],   # 1~3
                ["현금", 600, 500], [], ["2. 비유동자산", None, None],                      # 4~6
                ["유형자산", 900, 840], ["자산총계", 1500, 1340],                          # 7~8
                ["부채총계", 600, 640], ["자본총계", 900, 700]]                            # 9~10
        ts = X.scan_rows(rows)
        self.assertEqual(len(ts), 1)
        self.assertEqual((ts[0].kind, ts[0].last_row), ("BS", 10))

    def test_orphan_rows(self):
        from tests.fixtures.xlsx_fixture import _out
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "잡표"
        for row in (["메모", 123], [], [], ["과목", "당기", "전기"], ["자산총계", 100, 90],
                    ["부채총계", 40, 30], ["자본총계", 60, 60], [], [], ["참고 금액", 5]):
            ws.append(row)
        path = _out("sample_orphans.xlsx")
        wb.save(path)
        sh = X.scan(path).sheets[0]
        self.assertEqual(len(sh.tables), 1)
        self.assertEqual(sh.orphan_rows, [1, 10])
        self.assertEqual(sh.to_dict()["orphan_rows"], [1, 10])

    def test_next_table_without_blank_row(self):
        rows = [["27. 판매비와관리비"], ["(단위: 원)"], ["구분", "당기", "전기"],
                ["급여", 150, 140], ["합계", 150, 140],
                ["12. 유형자산"], ["구분", "토지", "건물", "합계"],
                ["기초", 300, 300, 600], ["기말", 300, 320, 620]]
        ts = X.scan_rows(rows)
        self.assertEqual(len(ts), 2)
        self.assertEqual([(t.header_row, t.last_row) for t in ts], [(3, 5), (7, 9)])
        self.assertEqual(ts[1].title, "12. 유형자산")
        self.assertEqual(ts[1].unit, 1)

    def test_bare_years_must_be_consecutive(self):
        rows = [["과목", 2100, 1995], ["매출액", 100, 90], ["영업이익", 40, 30], ["당기순이익", 30, 20]]
        self.assertTrue(all(t.header_row != 1 for t in X.scan_rows(rows)))
        rows[0] = ["과목", 2023, 2024, 2025]
        rows[1:] = [["매출액", 1, 2, 3], ["영업이익", 1, 2, 3], ["당기순이익", 1, 2, 3]]
        t, = X.scan_rows(rows)
        self.assertEqual({c: g.end_date for c, g in t.period_cols.items()}, {2: "2023", 3: "2024", 4: "2025"})

    def test_trailing_text_row_not_last_row(self):
        rows = [["구분", "당기", "전기"], ["급여", 150, 140], ["감가상각비", "-", 70],
                ["상기 금액은 모두 비용으로 인식하였습니다."]]
        t, = X.scan_rows(rows)
        self.assertEqual(t.last_row, 3)              # 대시만 있는 행은 본문, 설명 문장은 아님

    def test_string_year_headers(self):
        rows = [["과목", "2025", "2024"], ["매출액", 100, 90], ["영업이익", 40, 30], ["당기순이익", 30, 20]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.kind, t.header_row), ("IS", 1))
        self.assertEqual({c: g.end_date for c, g in t.period_cols.items()}, {2: "2025", 3: "2024"})

    def test_half_year_month_headers(self):
        rows = [["과목", "2026.06", "2025-06"], ["매출액", 100, 90], ["영업이익", 40, 30], ["반기순이익", 30, 20]]
        t, = X.scan_rows(rows)
        self.assertEqual({c: g.end_date for c, g in t.period_cols.items()}, {2: "2026-06", 3: "2025-06"})

    def test_unit_carried_note_in_reason(self):
        rows = [["(단위: 천원)"], ["과목", "당기", "전기"], ["자산총계", 100, 90], ["부채총계", 40, 30],
                ["자본총계", 60, 60], [], [], ["과목", "당기", "전기"],
                ["매출액", 100, 90], ["영업이익", 40, 30], ["당기순이익", 30, 20]]
        t1, t2 = X.scan_rows(rows)
        self.assertNotIn("이월", t1.reason)
        self.assertEqual(t2.unit, 1000)
        self.assertIn("단위 시트 상단에서 이월", t2.reason)

    def test_other_always_low(self):
        rows = [["구분", "당기", "전기"], ["가", 1, 2], ["나", 3, 4]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.kind, t.confidence), ("OTHER", "low"))

    def test_datetime_and_year_number_headers(self):
        import datetime as dt
        rows = [["과목", dt.datetime(2025, 12, 31), dt.datetime(2024, 12, 31)],
                ["자산총계", 100, 90], ["부채총계", 40, 30], ["자본총계", 60, 60]]
        t, = X.scan_rows(rows)
        self.assertEqual({c: g.end_date for c, g in t.period_cols.items()}, {2: "2025-12-31", 3: "2024-12-31"})
        rows = [["과목", 2025, 2024], ["매출액", 100, 90], ["영업이익", 40, 30], ["당기순이익", 30, 20]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.kind, t.header_row), ("IS", 1))
        self.assertEqual({c: g.end_date for c, g in t.period_cols.items()}, {2: "2025", 3: "2024"})

    def test_no_numeric_columns_not_reported(self):
        rows = [["구분", "당기", "전기"], ["설명", "해당 없음", "해당 없음"], ["비고", "없음", "없음"]]
        self.assertEqual(X.scan_rows(rows), [])

    def test_half_year_sub_columns_prefer_cumulative(self):
        """'당반기 | 전반기' 아래 '3개월 | 누적 | 3개월 | 누적': 같은 기간의 여러 열 중 '누적' 열이 기간 열. 값 열은 넷 다."""
        rows = [["손익계산서"], ["(단위: 원)"],
                ["과목", "당반기", None, "전반기", None], [None, "3개월", "누적", "3개월", "누적"],
                ["매출액", 50, 100, 40, 80], ["영업이익", 10, 20, 8, 16], ["반기순이익", 5, 10, 4, 8]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.kind, t.header_row, t.first_row), ("IS", 4, 5))
        self.assertEqual(t.value_cols, [2, 3, 4, 5])
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {3: "cur", 5: "prev"})
        self.assertEqual(t.period_cols[3].text, "당반기 누적")
        self.assertEqual(t.col_texts, {2: "당반기 3개월", 3: "당반기 누적", 4: "전반기 3개월", 5: "전반기 누적"})
        self.assertEqual(t.to_dict()["col_texts"]["2"], "당반기 3개월")
        # 금액·구성비처럼 '누적' 이 없으면 첫 숫자 열(기존 규칙)
        rows[3] = [None, "금액", "구성비", "금액", "구성비"]
        t, = X.scan_rows(rows)
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {2: "cur", 4: "prev"})

    def test_long_from_to_headers_give_period_cols(self):
        rows = [["손익계산서"], ["(단위: 원)"],
                ["과목", "제 12(당) 기 2025년 1월 1일부터 2025년 12월 31일까지", "제 11(전) 기 2024년 1월 1일부터 2024년 12월 31일까지"],
                ["매출액", 100, 90], ["영업이익", 40, 30], ["당기순이익", 30, 20]]
        t, = X.scan_rows(rows)
        self.assertEqual({c: (g.role, g.fiscal, g.end_date) for c, g in t.period_cols.items()},
                         {2: ("cur", 12, "2025-12-31"), 3: ("prev", 11, "2024-12-31")})

    def test_headerless_note_column_is_not_a_label(self):
        """머리글 없는 주석 번호 열('주석 4'·'4,5'): 라벨 열도 값 열도 아니다."""
        rows = [["재무상태표"], ["(단위: 원)"], ["과목", None, "당기", "전기"],
                ["현금및현금성자산", "주석 4", 100, 90], ["매출채권", "4,5", 50, 40], ["자산총계", None, 150, 130],
                ["부채총계", "주석 7", 60, 50], ["자본총계", None, 90, 80]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.kind, t.label_cols, t.value_cols), ("BS", [1], [3, 4]))
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {3: "cur", 4: "prev"})
        from core import file_collect as FC
        self.assertEqual(FC._label(rows[3], t.label_cols), "현금및현금성자산")

    def test_formula_only_sheet_has_note(self):
        """openpyxl 로 수식만 쓴 파일(계산값 없음): 글자만 있고 숫자가 없어 표를 못 찾고, 그 이유를 SheetScan.note 에 남긴다."""
        from tests.fixtures.xlsx_fixture import _out
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "재무상태표"
        ws.append(["재무상태표"]); ws.append(["(단위: 원)"]); ws.append(["과목", "당기", "전기"])
        for lab, a, b in (("자산총계", 100, 90), ("부채총계", 40, 30), ("자본총계", 60, 60)):
            ws.append([lab, f"={a}", f"={b}"])
        ws2 = wb.create_sheet("메모")
        ws2.append(["검토 메모"])
        path = _out("sample_formula_only.xlsx")
        wb.save(path)
        sc = X.scan(path)
        self.assertEqual([s.tables for s in sc.sheets], [[], []])
        self.assertEqual(sc.sheets[0].note, X.FORMULA_NOTE)
        self.assertEqual(sc.sheets[1].note, X.FORMULA_NOTE)        # 글자만 있는 시트 — 같은 파일이 수식 파일이므로 같은 안내
        self.assertEqual(sc.to_dict()["sheets"][0]["note"], X.FORMULA_NOTE)
        self.assertIsNone(X.scan(F.build_statements()).sheets[0].note)

    def test_rebuild_with_header_row_and_label_cols(self):
        rows = [["재무상태표"], ["(단위: 천원)"], ["코드", "과목", "당기", "전기"], [None, None, "제 12 기", "제 11 기"],
                ["A01", "자산총계", 100, 90], ["B01", "부채총계", 40, 30], ["C01", "자본총계", 60, 60]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.header_row, t.label_cols, t.unit), (4, [1, 2], 1000))
        self.assertEqual({c: (g.role, g.fiscal) for c, g in t.period_cols.items()}, {3: ("cur", 12), 4: ("prev", 11)})
        r = X.rebuild(rows, t, header_row=4, label_cols=[2])
        self.assertEqual((r.idx, r.kind, r.header_row, r.first_row, r.last_row, r.label_cols, r.unit), (0, "BS", 4, 5, 7, [2], 1000))
        # 머리글 행을 그대로 두고 라벨 열만 바꾸면 2행 머리글 위 행의 당기·전기 표식을 그대로 쓴다
        self.assertEqual({c: (g.role, g.fiscal) for c, g in r.period_cols.items()}, {3: ("cur", 12), 4: ("prev", 11)})
        self.assertEqual(r.header_start, 3)
        r = X.rebuild(rows, t, label_cols=[2])
        self.assertEqual({c: (g.role, g.fiscal) for c, g in r.period_cols.items()}, {3: ("cur", 12), 4: ("prev", 11)})
        self.assertIsNone(X.rebuild(rows, t, header_row=99))

    def test_rebuild_label_cols_only_keeps_upper_header_marks(self):
        """'코드|과목|당기||전기|' 위 '금액|구성비|금액|구성비' 아래: 라벨 열만 고쳐도 기간 열 {3: 당기, 5: 전기} 유지."""
        rows = [["손익계산서"], ["(단위: 원)"], ["코드", "과목", "당기", None, "전기", None],
                [None, None, "금액", "구성비", "금액", "구성비"],
                ["A01", "매출액", 1000, 1.0, 900, 1.0], ["A02", "매출원가", 600, 0.6, 500, 0.56],
                ["A03", "영업이익", 400, 0.4, 400, 0.44]]
        t, = X.scan_rows(rows)
        self.assertEqual((t.header_start, t.header_row), (3, 4))
        self.assertEqual({c: g.role for c, g in t.period_cols.items()}, {3: "cur", 5: "prev"})
        r = X.rebuild(rows, t, label_cols=[2])
        self.assertEqual((r.header_row, r.label_cols), (4, [2]))
        self.assertEqual({c: g.role for c, g in r.period_cols.items()}, {3: "cur", 5: "prev"})


if __name__ == "__main__":
    unittest.main()
