"""Excel COM 마무리 통합 테스트. 느리고 Excel 이 필요하므로 DART_FS_COM_TESTS=1 일 때만 실행."""
import os, shutil, tempfile, unittest
from datetime import date
import openpyxl
from core import pipeline as P, mapping as M
from excel import grid as G, write_openpyxl as W, build as B, compose as C
from tests.test_pipeline import FakeClient

COM = os.environ.get("DART_FS_COM_TESTS") == "1"


def _model():
    wb = G.WorkbookModel()
    s = wb.add(G.Sheet("설정", is_data=False))
    s.set(1, 2, "dart_fs 테스트", "title")
    for i, (k, v, n) in enumerate([("파이썬 경로", r"C:\py\python.exe", "cfg_python"),
                                   ("refresh.py 경로", r"C:\app\excel\refresh.py", "cfg_script"),
                                   ("상태", "생성 완료", "cfg_status")]):
        s.set(4 + i, 2, k, "bold"); s.set(4 + i, 3, v); s.names[n] = f"C{4 + i}"
    r = wb.add(G.Sheet("rFS_별도"))
    r.set(1, 4, "재무상태표", "section"); r.set(2, 4, "(단위 : 원)", "unit")
    r.set(5, 4, "자산총계", "bold"); r.set(5, 5, 1234567, "num")
    r.set(6, 4, "부채와자본총계", "bold"); r.set(6, 5, 1234567, "num")
    m = wb.add(G.Sheet("mFS"))
    m.set(1, 4, "자산총계"); m.set(1, 5, "='rFS_별도'!E5", "num", formula=True)
    m.set(2, 4, "부채와자본총계"); m.set(2, 5, "='rFS_별도'!E6", "num", formula=True)
    m.set(3, 4, "검증 TRUE 수"); m.set(3, 5, "=COUNTIF('검증'!E1:E2,TRUE)", "num", formula=True)
    k = wb.add(G.Sheet("검증"))
    k.set(1, 4, "자산 = 부채와자본"); k.set(1, 5, "=ROUND('mFS'!E1/1000,0)=ROUND('mFS'!E2/1000,0)", "check", formula=True)
    k.set(1, 8, "=ROUND('mFS'!E1-'mFS'!E2,0)", "num", formula=True)            # 차이 열은 검증 셀로 세지 않는다
    k.set(2, 5, "해당 없음", "italic")
    return wb


@unittest.skipUnless(COM, "DART_FS_COM_TESTS=1 일 때만(Excel COM 필요)")
class TestFinalizeCom(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_finalize_to_xlsm(self):
        from excel import finalize as FZ
        xlsx = W.write(_model(), os.path.join(self.d, "t.xlsx"))
        xlsm = os.path.join(self.d, "t.xlsm")
        res = FZ.finalize(xlsx, xlsm, B.BAS)
        print("\nverify:", res)
        self.assertTrue(os.path.exists(xlsm))
        self.assertFalse(os.path.exists(xlsx))
        self.assertIsNotNone(openpyxl.load_workbook(xlsm, keep_vba=True).vba_archive)
        self.assertEqual(res["errors"], 0)
        self.assertGreaterEqual(res["checks"]["TRUE"], 1)
        self.assertEqual(res["checks"]["FALSE"], 0)
        self.assertEqual(res["checks"]["해당 없음"], 1)

        res2 = FZ.finalize_inplace(xlsm, B.BAS)      # 재실행해도 모듈·버튼 중복 없이
        print("verify(inplace):", res2)
        self.assertEqual(res2["checks"]["TRUE"], res["checks"]["TRUE"])
        x = openpyxl.load_workbook(xlsm, keep_vba=True)
        names = [n for n in x.vba_archive.namelist() if n.lower().startswith("xl/vbaproject")]
        self.assertTrue(names)

        import pythoncom, win32com.client as win32
        pythoncom.CoInitialize()
        app = wb = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(os.path.abspath(xlsm))
            comps = [c.Name for c in wb.VBProject.VBComponents]
            self.assertEqual(comps.count("DartFs"), 1, comps)
            btns = wb.Worksheets(1).Buttons()
            self.assertEqual(btns.Count, 1)
            self.assertEqual(btns(1).Caption, "업데이트")
            self.assertTrue(btns(1).OnAction.endswith("DartFs.UpdateData"), btns(1).OnAction)
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if app is not None:
                app.Quit()
            app = wb = None
            pythoncom.CoUninitialize()

    def test_hugel_fixture_false_cells(self):
        """hugel 픽스처(연결+별도, 백만원, 2025H1 포함) 검증 결과 기대값. FALSE 4(수정 전) = B3·B4·B13 반올림 경계 단수차 3셀
        (E125·F126·E135, −675·1,275·−1,314원) + F28(연결 B7 FY2025: 캡션에 '지역' 이 든 지역별 매출액 표 대신 '주요 지리적 시장·주요 제품'
        표의 첫 합계 425,109,983천원을 고르는데 원문 그 합계가 IS 매출액 425,109,984천원과 1천원 차이 — 원문 표 자체의 단수차).
        Task 5 수정 전 FALSE 였던 법인세 '법인세비용과 회계이익의 관계' 조정표 C-합계 4셀(E85·F85·E181·F181, 첫 행 회계이익이
        합산 대상이 아님)은 캡션 '회계이익' 규칙으로 '해당 없음(조정표)'.
        최종 리뷰 2차: ① 각주 (*1)·(*2) 정규화로 무형자산 '기 타 (*1)'/'기 타 (*2)' 열 블록과 관계기업투자 연도별 블록이 하나로
        묶여 연결 C 항목이 2개 줄고 그 아래 행이 2행씩 올라감(E125·F126·E135 → E123·F124·E133, 같은 B3·B4·B13 단수차).
        ② '당기말' 을 기말로 인식해 B12 FY2025(연결 F33·별도 F132)가 충당부채 변동표 '당기말' 열 합계 1,801,283천원을 고름 —
        BS 유동 1,064,283,502 + 비유동 737,000,000 = 1,801,283,502원 → 1,801,284천원, 반올림 경계 단수차(−502원).
        ③ A13 FY2024·FY2025 가 '자본총계' 행으로, C14~C16 사용권자산 변동표 FY2025 가 '당기말' 행으로 TRUE.
        ①~③ 합쳐 TRUE 290→295, 해당 없음 258→245."""
        from excel import finalize as FZ
        cfg = P.BuildConfig("HG", "휴젤", date(2026, 10, 26), 2024, 2025, True, ["CFS", "OFS"], 1000000,
                            {"recommended": True}, "default")
        col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        xlsx = W.write(C.compose(col, None, "p", "s"), os.path.join(self.d, "hg.xlsx"))
        res = FZ.finalize(xlsx, os.path.join(self.d, "hg.xlsm"), B.BAS)
        print("\nverify(hugel):", res)
        self.assertEqual(res["errors"], 0)
        self.assertEqual(res["false_cells"], ["검증!F28", "검증!F33", "검증!E123", "검증!F124", "검증!F132", "검증!E133"])
        self.assertEqual(res["checks"], {"TRUE": 295, "FALSE": 6, "해당 없음": 245})

    def test_rebuild_backs_up_existing(self):
        from excel import finalize as FZ
        xlsm = os.path.join(self.d, "t.xlsm")
        FZ.finalize(W.write(_model(), os.path.join(self.d, "a.xlsx")), xlsm, B.BAS)
        res = FZ.finalize(W.write(_model(), os.path.join(self.d, "b.xlsx")), xlsm, B.BAS)
        self.assertTrue(res["backup"] and os.path.exists(res["backup"]))
        self.assertTrue(os.path.exists(xlsm))


if __name__ == "__main__":
    unittest.main()
