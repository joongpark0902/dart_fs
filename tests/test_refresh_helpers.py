import os, tempfile, unittest
from excel import refresh as R
from excel import write_com as WC


class TestHelpers(unittest.TestCase):
    def test_com_value_dates_are_excel_serials(self):
        from datetime import date, datetime
        self.assertEqual(WC._com_value(date(2025, 12, 31)), 46022.0)
        self.assertEqual(WC._com_value(datetime(2025, 12, 31, 12, 0)), 46022.5)
        self.assertEqual(WC._com_value(3_000_000_000), 3e9)

    def test_parse_toc_rows(self):
        rows = [["번호", "제목", "표준 주제", "포함", "추천", "배치 시트", "표 수", "정규화 제목"],
                ["11", "유형자산", "유형자산", True, "★", "rN_유형자산", 3, "유형자산"],
                ["1", "회사의 개요", "", False, "", "rN_기타", 1, "회사의개요"],
                ["27", "판관비", "판관비", "TRUE", "★", "rN_판관비", 1, "판관비"],
                ["28", "기타", "", "예", "", "rN_기타", 1, "기타"],
                ["29", "빈행", "", None, "", "rN_기타", 1, "빈행"],
                [None] * 8]
        self.assertEqual(R.parse_toc_rows(rows), {"유형자산": True, "회사의개요": False, "판관비": True,
                                                  "기타": True, "빈행": False})

    def test_backup_name(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "a.xlsm")
            with open(p, "wb") as f:
                f.write(b"PK")
            b = R.backup(p)
            self.assertTrue(os.path.basename(b).startswith("backup_") and b.endswith("_update.xlsm"))
            self.assertTrue(os.path.exists(b))
            b2 = R.backup(p)                      # 같은 mtime 으로 다시 → 덮어쓰지 않고 새 이름
            self.assertNotEqual(b, b2)
            self.assertTrue(os.path.exists(b2))

    def test_blocks_from_columns(self):
        # C열 마커(§) + D열: 마커 행 다음 3행(단위·헤더 2행)은 건너뛰고 다음 마커까지 비어 있지 않은 D 값
        colC = [None, "§", None, None, None, None, None, "§", None, None, None, None]
        colD = [None, "재무상태표", "(단위 : 원)", "", "", "자 산", "I. 유동자산", "포괄손익계산서", "(단위 : 원)", "", "", "매출액"]
        self.assertEqual(WC.blocks_from_columns(colC, colD),
                         {"재무상태표": ["자 산", "I. 유동자산"], "포괄손익계산서": ["매출액"]})
        self.assertIs(R.blocks_from_columns, WC.blocks_from_columns)

    def test_blocks_both_header_shapes(self):
        # 재무제표 블록(마커·단위·헤더 2행·데이터=마커+4)과 rN_기타 원표(마커·단위·헤더 1행·데이터=마커+3)
        colC = ["§", None, None, None, None, None, None, None, "§", None, None, None, None]
        colD = ["재무상태표", "(단위 : 원)", "", None, "자 산", "부채", None, None,
                "주석 1. 기타 | 표 [2025]", "(단위 : 천원)", None, "첫 데이터", "둘째"]
        colE = [None, None, "2025", "2025-12-31", 1, 2, None, None, None, "2025 · 주석 1", "금액", 3, 4]
        self.assertEqual(len(colE), len(colD))
        self.assertEqual(WC.blocks_from_columns(colC, colD),
                         {"재무상태표": ["자 산", "부채"], "주석 1. 기타 | 표 [2025]": ["첫 데이터", "둘째"]})

    def test_blocks_ignore_rows_before_first_marker_and_marker_without_title(self):
        colC = ["§", None, "§", None, None, None, None]
        colD = ["메모", "", None, "x", "y", "z", "w"]
        self.assertEqual(WC.blocks_from_columns(colC, colD), {"메모": []})   # 두 번째 § 는 D 가 비어 블록 아님
        self.assertEqual(WC.blocks_from_columns([], []), {})


    def test_status_text_mentions_shifts(self):
        res = {"errors": 0, "checks": {"TRUE": 3, "FALSE": 0}}
        self.assertNotIn("밀림", R._status_text(res, ""))
        sh = [f"rFS_별도 'b{i}': 블록이 6행 밀림 — 모델 참조 확인 필요" for i in range(5)]
        t = R._status_text(res, " (저장은 Ctrl+S)", sh)
        self.assertIn("경고 블록 밀림 5건", t)
        self.assertIn("'b0'", t); self.assertNotIn("'b3'", t)
        self.assertIn("외 2건", t)
        self.assertTrue(t.endswith("(저장은 Ctrl+S)"))

    def test_write_model_restores_original_calculation(self):
        """사용자가 수동 계산이었으면 수동으로 돌려준다(자동으로 바꾸지 않음)."""
        from excel.grid import WorkbookModel

        class App:
            def __init__(self, calc):
                self.Calculation = calc
                self.seen = []

            def __setattr__(self, k, v):
                if k == "Calculation" and hasattr(self, "seen"):
                    self.seen.append(v)
                object.__setattr__(self, k, v)

        class Sheet:
            Name = "설정"

            def Activate(self):
                pass

        class Wb:
            def __init__(self, app):
                self.Application, self.ActiveSheet, self.Worksheets = app, Sheet(), []

        for calc in (WC.XL_CALC_MANUAL, WC.XL_CALC_AUTOMATIC):
            app = App(calc)
            WC.write_model(Wb(app), WorkbookModel())
            self.assertEqual(app.Calculation, calc)
            self.assertEqual(app.seen[0], WC.XL_CALC_MANUAL)          # 쓰는 동안은 수동


if __name__ == "__main__":
    unittest.main()
