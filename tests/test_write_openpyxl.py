import os, tempfile, unittest, openpyxl
from datetime import date
from core import pipeline as P, mapping as M
from excel import grid as G, write_openpyxl as W, compose as C, build as B
from tests.test_pipeline import FakeClient


class TestWrite(unittest.TestCase):
    def test_formats_and_formula_guard(self):
        wb = G.WorkbookModel()
        s = wb.add(G.Sheet("설정", is_data=False))
        s.set(1, 2, "제목", "title"); s.set(2, 2, "= Rf + β", "italic"); s.names = {"cfg_status": "C3"}; s.set(3, 3, "ok")
        r = wb.add(G.Sheet("rFS_별도"))
        r.set(1, 4, "재무상태표", "section"); r.set(2, 4, "(단위 : 원)", "unit"); r.set(3, 5, "제 28 기", "hdr"); r.set(4, 5, date(2025, 12, 31), "date")
        r.set(5, 4, "자산총계", "bold"); r.set(5, 5, -1234567, "num"); r.col_widths = {4: 44}
        m = wb.add(G.Sheet("mFS")); m.set(1, 5, "='rFS_별도'!E5", "num", formula=True); m.set(2, 5, "=ROUND(E1-E1,0)=0", "check", formula=True)
        m.freeze = "E2"
        with tempfile.TemporaryDirectory() as d:
            p = W.write(wb, os.path.join(d, "t.xlsx"))
            x = openpyxl.load_workbook(p)
            self.assertEqual(x.sheetnames, ["설정", "rFS_별도", "mFS"])
            ws = x["rFS_별도"]
            self.assertEqual(ws["E5"].number_format, '#,##0;[Red](#,##0);"-"')
            self.assertEqual(ws["E5"].alignment.horizontal, "right")
            self.assertEqual(ws["E4"].number_format, "yyyy-mm-dd")
            self.assertEqual(ws["E3"].alignment.horizontal, "center")
            self.assertEqual(ws["E3"].border.top.style, "thin"); self.assertEqual(ws["E3"].border.bottom.style, "thin")
            self.assertTrue(ws["D5"].font.bold); self.assertEqual(ws["D5"].font.name, "맑은 고딕"); self.assertEqual(ws["D5"].font.size, 10)
            self.assertEqual(x["설정"]["B1"].font.size, 12); self.assertTrue(x["설정"]["B1"].font.bold)
            self.assertEqual(ws["D1"].font.size, 11)
            self.assertTrue(ws["D2"].font.italic)
            self.assertFalse(ws.sheet_view.showGridLines)
            self.assertEqual(ws.sheet_view.zoomScale, 85)
            self.assertEqual(x["mFS"]["E1"].value, "='rFS_별도'!E5")
            self.assertEqual(x["mFS"]["E1"].data_type, "f")
            self.assertEqual(x["mFS"]["E1"].number_format, W.NUM_FMT)
            self.assertEqual(x["mFS"]["E2"].number_format, "General")
            self.assertEqual(x["mFS"].freeze_panes, "E2")
            self.assertEqual(x["설정"]["B2"].value, "= Rf + β")          # 텍스트로 보존(수식 아님)
            self.assertTrue(x["설정"]["B2"].quotePrefix)
            self.assertFalse(x["mFS"]["E1"].quotePrefix)
            self.assertEqual(x["설정"]["B2"].data_type, "s")
            self.assertIn("cfg_status", x.defined_names)
            self.assertEqual(x.defined_names["cfg_status"].attr_text, "'설정'!$C$3")
            self.assertEqual(ws.column_dimensions["D"].width, 44)

    def test_bool_not_numeric_format(self):
        wb = G.WorkbookModel()
        s = wb.add(G.Sheet("목차")); s.set(1, 4, True, "check")
        with tempfile.TemporaryDirectory() as d:
            x = openpyxl.load_workbook(W.write(wb, os.path.join(d, "b.xlsx")))
            self.assertIs(x["목차"]["D1"].value, True)
            self.assertNotEqual(x["목차"]["D1"].number_format, W.NUM_FMT)

    def test_composed_model_writes_and_only_flagged_formulas(self):
        cfg = P.BuildConfig("OB", "테스트회사", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, {"recommended": True}, "default")
        col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        model = C.compose(col, None, r"C:\py\python.exe", B.REFRESH)
        flagged = {(sh.name, r, c) for sh in model.sheets for (r, c), cell in sh.cells.items() if cell.formula}
        with tempfile.TemporaryDirectory() as d:
            p = W.write(model, os.path.join(d, "e2e.xlsx"))
            x = openpyxl.load_workbook(p)
            self.assertEqual(x.sheetnames, [s.name for s in model.sheets])
            found = {(ws.title, c.row, c.column) for ws in x.worksheets for row in ws.iter_rows() for c in row if c.data_type == "f"}
            self.assertEqual(found, flagged)
            self.assertTrue(flagged)
            for n in ("cfg_python", "cfg_script", "cfg_status"):
                self.assertIn(n, x.defined_names)

            # 검증 시트: FALSE 파랑 조건부 서식이 grid 의 check 셀 전부를 덮는다
            chk = next(sh for sh in model.sheets if sh.name == "검증")
            want = {(r, c) for (r, c), cell in chk.cells.items() if cell.fmt == "check" and cell.formula}
            self.assertTrue(want)
            ws = x["검증"]
            covered, rules = set(), []
            for cf in ws.conditional_formatting:
                for rule in cf.rules:
                    rules.append(rule)
                    if rule.type == "cellIs" and rule.operator == "equal" and rule.formula == ["FALSE"]:
                        self.assertEqual(rule.dxf.font.color.rgb[-6:], "0000FF")
                        for cr in cf.sqref.ranges:
                            for r in range(cr.min_row, cr.max_row + 1):
                                for c in range(cr.min_col, cr.max_col + 1):
                                    covered.add((r, c))
            self.assertEqual(covered, want)
            e = min(want)
            self.assertEqual(ws.cell(row=e[0], column=e[1]).alignment.horizontal, "center")
            self.assertEqual(ws.cell(row=e[0], column=e[1]).number_format, "General")
            # 값 열 폭 15: mFS·검증·rFS·rN(기타 포함). 라벨 열 D 는 그대로
            ncol = len(col.periods)
            for name in [n for n in x.sheetnames if n in ("mFS", "검증") or n.startswith(("rFS_", "rN_"))]:
                if name == C.ETC_SHEET and x[name].max_column < 5:
                    continue
                for i in range(ncol):
                    self.assertEqual(x[name].column_dimensions[chr(ord("E") + i)].width, 15, (name, i))
                if name != "검증":
                    self.assertEqual(x[name].column_dimensions["D"].width, 44, name)
            for i in range(ncol):                                  # 검증 차이 열(값 열 + 빈 열 1 뒤)
                self.assertEqual(x["검증"].column_dimensions[chr(ord("E") + ncol + 1 + i)].width, 15)
            self.assertEqual(x["검증"].column_dimensions["D"].width, 52)

    def test_check_ranges_contiguous_runs(self):
        s = G.Sheet("검증")
        for r in (5, 6, 7, 9):
            s.set(r, 5, "=ROUND(1,0)=ROUND(1,0)", "check", formula=True)
        s.set(5, 6, "=ROUND(1,0)=ROUND(1,0)", "check", formula=True)
        s.set(8, 5, "해당 없음", "italic")
        s.set(10, 4, True, "check")                                  # 수식 아닌 check(목차) 제외
        self.assertEqual(W.check_ranges(s), ["E5:E7", "E9", "F5"])


class TestBas(unittest.TestCase):
    def test_bas_cp949_crlf(self):
        with open(B.BAS, "rb") as f:
            raw = f.read()
        self.assertTrue(raw.startswith(b'Attribute VB_Name = "DartFs"'))
        self.assertIn(b"\r\n", raw); self.assertNotIn(b"\n", raw.replace(b"\r\n", b""))
        txt = raw.decode("cp949")
        self.assertIn("업데이트 실행 중", txt)
        self.assertIn('Names("cfg_python")', txt)
        with self.assertRaises(UnicodeDecodeError):
            raw.decode("utf-8")

    def test_safe(self):
        self.assertEqual(B.safe('a/b:c* d'), "a_b_c_d")


if __name__ == "__main__":
    unittest.main()
