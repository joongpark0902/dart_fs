"""refresh COM 통합 테스트(run_open·main 열린 경로·run_closed). DART_FS_COM_TESTS=1 일 때만.

사용자 Excel 에는 붙지 않는다: 우리 DispatchEx 인스턴스(Visible False)에서 파일을 열고 그 (app, wb) 를
run_open 에 직접 넘긴다. 시작 전부터 있던 EXCEL.EXE PID 와 우리 인스턴스 PID 가 다른지 확인한다.
"""
import glob, os, shutil, subprocess, tempfile, time, unittest
from datetime import date
from unittest import mock
import openpyxl
from core import pipeline as P, mapping as M
from excel import compose as C, write_openpyxl as W, build as B
from tests.test_pipeline import FakeClient

COM = os.environ.get("DART_FS_COM_TESTS") == "1"


def _excel_pids():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True).stdout
    pids = set()
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def _wait_pids_within(allowed, timeout=15.0):
    """EXCEL.EXE PID 가 allowed 의 부분집합이 될 때까지(최대 timeout 초) 기다려 현재 PID 집합 반환."""
    end = time.time() + timeout
    while True:
        cur = _excel_pids()
        if cur <= allowed or time.time() > end:
            return cur
        time.sleep(0.5)


# ob 픽스처(FY2024~FY2025, 별도)의 검증 결과 기대값 — 전부 주석 천원 합계와 BS 원 단위의 반올림 경계 단수차(허용폭 없음 규칙):
# E23·F23 B3 유형자산(차이 −784·−1,333원), F24 B4 무형자산(577원), F32 B12 충당부채(895원). 바뀌면 선택 규칙·행 배치가 바뀐 것.
OB_FALSE_CELLS = ["검증!E23", "검증!F23", "검증!F24", "검증!F32"]
# Task 5: TRUE 130→135, 해당 없음 26→21 — B9·B10(소제목 '리스부채'·'사용권자산' 아래 합계 규칙, 2기간 ×2) +
# C44 FY2024(각주 P 가 다음 캡션에 섞이지 않아 이연법인세 증감표 당기·전기 블록이 하나로 묶임). FALSE 셀은 그대로.
# 최종 리뷰 2차: TRUE 135→138, 해당 없음 21→24 — C47~C49 퇴직급여 순확정급여부채 변동표(열 블록 3개)가 '당기초'·'당기말' 행을
# 기초·기말로 인식해 검산 항목이 됨(예전엔 '기초…' 로 시작하는 행이 없어 빠짐). FY2025 TRUE(소계 행 규칙), FY2024 는 '전기초'·'전기말'
# 이라 해당 없음. 그 뒤 C 번호가 3씩 밀리지만 FALSE 셀(A·B 행)은 그대로.
OB_CHECKS = {"TRUE": 138, "FALSE": 4, "해당 없음": 24}



@unittest.skipUnless(COM, "DART_FS_COM_TESTS=1 일 때만(Excel COM 필요)")
class TestRunOpenCom(unittest.TestCase):
    def setUp(self):
        self.user_pids = _excel_pids()           # finalize 인스턴스를 띄우기 전(사용자 Excel 등)
        self.d = tempfile.mkdtemp()
        cfg = P.BuildConfig("OB", "테스트", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1,
                            {"recommended": True}, "default")
        self.cfg = cfg
        self.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        model = C.compose(self.col, None, r"C:\py\python.exe", B.REFRESH)
        xlsx = W.write(model, os.path.join(self.d, "t.xlsx"))
        wb = openpyxl.load_workbook(xlsx)
        user = wb.create_sheet("모델", 1)
        user["B2"] = "내 메모"
        user["C3"] = 777
        user["C4"] = "='rFS_별도'!E5"
        wb["rFS_별도"]["Z300"] = "stale"
        wb.remove(wb["로그"])                                  # 없는 데이터 시트 → 새로 만들어져야 함
        wb.save(xlsx)
        from excel import finalize as FZ
        self.xlsm = os.path.join(self.d, "t.xlsm")
        FZ.finalize(xlsx, self.xlsm, B.BAS)
        _wait_pids_within(self.user_pids)        # finalize 인스턴스가 완전히 끝날 때까지

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_run_open_writes_data_sheets(self):
        import pythoncom, win32com.client as win32, win32process
        from excel import refresh as R
        before = _excel_pids()
        pythoncom.CoInitialize()
        app = wb = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
            print(f"\n기존 EXCEL PID {sorted(before)} / 테스트 인스턴스 PID {pid}")
            self.assertNotIn(pid, before)
            wb = app.Workbooks.Open(self.xlsm)
            active_before = wb.ActiveSheet.Name
            logs = []
            with mock.patch.object(R, "_collect_for",
                                   return_value=(self.cfg, self.col, {"python": r"C:\py\python.exe"})):
                res = R.run_open(app, wb, lambda m: (logs.append(m), print(m)))
            print("verify:", res)
            status = wb.Names("cfg_status").RefersToRange.Value2
            print("status:", status)
            self.assertTrue(str(status).startswith("업데이트 완료"))
            self.assertIn("(저장은 Ctrl+S)", status)
            self.assertEqual(res["errors"], 0)
            self.assertEqual(res["false_cells"], OB_FALSE_CELLS)
            self.assertEqual(res["checks"], OB_CHECKS)
            self.assertEqual(wb.ActiveSheet.Name, active_before)
            self.assertEqual(app.Calculation, -4105)

            names = [ws.Name for ws in wb.Worksheets]
            print("sheets:", names)
            self.assertEqual(names[1], "모델")
            self.assertEqual(names[-1], "로그")                 # 매핑표 뒤에 새로 생성
            self.assertEqual(names[-2], "매핑표")
            ws_user = wb.Worksheets("모델")
            self.assertEqual(ws_user.Range("B2").Value2, "내 메모")
            self.assertEqual(ws_user.Range("C3").Value2, 777)

            rfs = wb.Worksheets("rFS_별도")
            self.assertIsNone(rfs.Range("Z300").Value2)
            cd = rfs.Range("C1:D10").Value2
            self.assertIn(("§", "재무상태표"), [tuple(r) for r in cd])
            row = next(i + 1 for i, r in enumerate(cd) if r[1] == "재무상태표")
            self.assertTrue(rfs.Cells(row, 4).Font.Bold)
            self.assertEqual(rfs.Cells(row, 4).Font.Name, "맑은 고딕")
            self.assertEqual(rfs.Columns(4).ColumnWidth, 44)
            ecol = [r[0] for r in rfs.Range(f"E{row + 4}:E{row + 60}").Value2]
            num_row = row + 4 + next(i for i, v in enumerate(ecol) if isinstance(v, float))
            self.assertIn(rfs.Cells(num_row, 5).NumberFormat, (W.NUM_FMT, W.NUM_FMT.replace("[Red]", "[빨강]")))

            self.assertEqual(rfs.Columns(5).ColumnWidth, 15)          # 값 열 폭

            chk = wb.Worksheets("검증")                              # FALSE → 파랑 글꼴 조건부 서식(업데이트 후에도)
            fcs = chk.Cells.FormatConditions
            self.assertGreaterEqual(fcs.Count, 1)
            fc = fcs(1)
            self.assertEqual((fc.Type, fc.Operator), (1, 3))
            self.assertIn("FALSE", fc.Formula1.upper())
            self.assertEqual(fc.Font.Color, 0xFF0000)                # RGB 0000FF (BGR)
            fr = int(OB_FALSE_CELLS[0][len("검증!E"):])               # 검증!E23 = FALSE → 파랑
            self.assertIs(chk.Cells(fr, 5).Value2, False)
            self.assertEqual(chk.Cells(fr, 5).DisplayFormat.Font.Color, 0xFF0000)
            true_row = next(r for r in range(5, 60) if chk.Cells(r, 5).Value2 is True)
            self.assertEqual(chk.Cells(true_row, 5).DisplayFormat.Font.Color, 0)   # TRUE 는 검정
            self.assertEqual(chk.Columns(5).ColumnWidth, 15)
            n_rules = fcs.Count
            with mock.patch.object(R, "_collect_for",                 # 두 번째 업데이트: 규칙이 쌓이지 않음
                                   return_value=(self.cfg, self.col, {"python": r"C:\py\python.exe"})):
                R.run_open(app, wb, lambda m: None)
            self.assertEqual(wb.Worksheets("검증").Cells.FormatConditions.Count, n_rules)

            mfs = wb.Worksheets("mFS")
            forms = [f for r in mfs.UsedRange.Formula for f in r if isinstance(f, str) and f.startswith("=")]
            self.assertTrue(forms)
            toc = wb.Worksheets("목차")
            self.assertEqual(toc.Range("H1").Value2, "정규화 제목")
            self.assertIsInstance(toc.Range("A2").Value2, str)       # '11' 같은 문자열이 숫자로 바뀌지 않음

            ex = R.WC.read_existing(wb)                      # 다시 읽어도 블록 구조 유지(COM 날짜 일련번호 → 기간 키)
            bs = ex["rFS_별도"]["blocks"]["재무상태표"]
            self.assertEqual(bs["start"], 1)
            self.assertTrue(bs["labels"])
            self.assertEqual(ex["rFS_별도"]["periods"], ["FY2024", "FY2025"])
            self.assertEqual(ex["mFS"]["periods"], ["FY2024", "FY2025"])
            self.assertEqual(ex, C.existing_from_model(C.compose(self.col, ex, "p", "s")))   # 같은 입력이면 같은 구조
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if app is not None:
                app.Quit()
            app = wb = None
            pythoncom.CoUninitialize()

    def _with_instance(self, body):
        """우리 DispatchEx 인스턴스에서 self.xlsm 을 열어 body(app, wb) 실행. 사용자 Excel 은 건드리지 않음."""
        import pythoncom, win32com.client as win32, win32process
        before = _excel_pids()
        pythoncom.CoInitialize()
        app = wb = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
            print(f"\n기존 EXCEL PID {sorted(before)} / 테스트 인스턴스 PID {pid}")
            self.assertNotIn(pid, before)
            wb = app.Workbooks.Open(self.xlsm)
            return body(app, wb)
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if app is not None:
                app.Quit()
            app = wb = None
            pythoncom.CoUninitialize()

    def test_main_open_path_via_attach(self):
        from excel import refresh as R

        def body(app, wb):
            with mock.patch.object(R, "_attach", return_value=(app, wb)), \
                 mock.patch.object(R, "_collect_for",
                                   return_value=(self.cfg, self.col, {"python": r"C:\py\python.exe"})):
                rc = R.main(["--workbook", self.xlsm])
            status = wb.Names("cfg_status").RefersToRange.Value2
            print("rc:", rc, "status:", status)
            self.assertEqual(rc, 0)
            self.assertTrue(str(status).startswith("업데이트 완료"))
            self.assertIn("(저장은 Ctrl+S)", status)
            self.assertEqual(app.Calculation, -4105)
        self._with_instance(body)
        self.assertTrue(glob.glob(os.path.join(self.d, "backup_*_update.xlsm")))


@unittest.skipUnless(COM, "DART_FS_COM_TESTS=1 일 때만(Excel COM 필요)")
class TestRunClosedCom(TestRunOpenCom):
    """닫힌 파일: run_closed 가 우리 DispatchEx 인스턴스로 열어 쓰고 저장·Quit."""
    test_run_open_writes_data_sheets = None
    test_main_open_path_via_attach = None

    def test_run_closed_preserves_user_sheet_and_saves(self):
        from excel import refresh as R
        before = _excel_pids()
        bk = R.backup(self.xlsm)                   # main 이 하는 백업과 같은 것
        logs = []
        with mock.patch.object(R, "_collect_for",
                               return_value=(self.cfg, self.col, {"python": r"C:\py\python.exe"})):
            res = R.run_closed(self.xlsm, lambda m: (logs.append(m), print(m)))
        print("verify:", res)
        self.assertFalse(any("openpyxl 경로" in m for m in logs))       # COM 경로를 탔다
        self.assertEqual(res["errors"], 0)
        self.assertEqual(res["false_cells"], OB_FALSE_CELLS)
        self.assertEqual(res["checks"], OB_CHECKS)
        after = _wait_pids_within(before)          # 우리 인스턴스 종료 확인(사용자 Excel 은 그대로)
        print(f"EXCEL PID 실행 전 {sorted(before)} / 후 {sorted(after)}")
        self.assertLessEqual(after, before)
        self.assertLessEqual(self.user_pids & before, after)
        self.assertTrue(os.path.exists(bk))
        self.assertFalse(os.path.exists(R.lock_path(self.xlsm)))

        wb = openpyxl.load_workbook(self.xlsm, keep_vba=True)
        self.assertIsNotNone(wb.vba_archive)
        user = wb["모델"]
        self.assertEqual(user["B2"].value, "내 메모")
        self.assertEqual(user["C3"].value, 777)
        self.assertEqual(user["C4"].value.replace("'", ""), "=rFS_별도!E5")   # Excel 저장 시 따옴표 정규화
        rfs = wb["rFS_별도"]
        self.assertIsNone(rfs["Z300"].value)
        rows = [(rfs.cell(row=r, column=3).value, rfs.cell(row=r, column=4).value) for r in range(1, rfs.max_row + 1)]
        self.assertIn(("§", "재무상태표"), rows)
        self.assertIn("로그", wb.sheetnames)
        status = str(wb["설정"]["C17"].value)
        print("status:", status)
        self.assertTrue(status.startswith("업데이트 완료"))
        self.assertNotIn("Ctrl+S", status)


if __name__ == "__main__":
    unittest.main()
