"""닫힌 파일 경로 중 COM 없는 부분 — openpyxl 대체 경로(run_closed_openpyxl)·잠금 파일 감지.
네트워크·COM 없이: _collect_for·finalize_inplace 를 패치하고 main 의 run_closed 를 대체 경로로 돌린다.
COM 기반 run_closed 는 tests/test_refresh_com.py(DART_FS_COM_TESTS=1)."""
import glob, os, shutil, tempfile, unittest
from datetime import date
from unittest import mock
import openpyxl
from core import pipeline as P, mapping as M
from excel import compose as C, write_openpyxl as W, refresh as R
from tests.test_pipeline import FakeClient

FAKE_VERIFY = {"errors": 0, "checks": {"TRUE": 1, "FALSE": 0, "해당 없음": 0}, "false_cells": []}


class TestRunClosed(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cfg = P.BuildConfig("OB", "테스트", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1,
                            {"recommended": True}, "default")
        cls.cfg = cfg
        cls.col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)

    def setUp(self):
        self.d = tempfile.mkdtemp()
        model = C.compose(self.col, None, r"C:\py\python.exe", r"C:\app\excel\refresh.py")
        tmp = W.write(model, os.path.join(self.d, "src.xlsx"))
        wb = openpyxl.load_workbook(tmp)
        user = wb.create_sheet("모델", 1)                         # 사용자 시트(설정 뒤)
        user["B2"] = "내 메모"
        user["C3"] = 12345
        user["C4"] = "='rFS_별도'!E5"
        rfs = wb["rFS_별도"]
        rfs["Z300"] = "stale"                                     # 업데이트 후 사라져야 함
        self.path = os.path.join(self.d, "wb.xlsm")
        wb.save(self.path)
        self.order_before = wb.sheetnames

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def _run_main(self, cfg=None):
        with mock.patch.object(R, "_collect_for", return_value=(cfg or self.cfg, self.col, {"python": r"C:\py\python.exe"})), \
             mock.patch.object(R, "_attach", return_value=(None, None)), \
             mock.patch.object(R, "run_closed", side_effect=R.run_closed_openpyxl), \
             mock.patch.object(R.FZ, "finalize_inplace", return_value=dict(FAKE_VERIFY)) as fin:
            rc = R.main(["--workbook", self.path])
        return rc, fin

    def test_closed_update_rewrites_data_sheets_and_keeps_user_sheet(self):
        rc, fin = self._run_main()
        self.assertEqual(rc, 0)
        fin.assert_not_called()                 # Excel 없는 대체 경로에서 finalize_inplace 를 부르지 않는다
        self.assertTrue(glob.glob(os.path.join(self.d, "backup_*_update.xlsm")))
        with open(os.path.join(self.d, R.LOG_NAME), encoding="utf-8") as f:
            logtxt = f.read()
        self.assertIn("dart_fs 업데이트:", logtxt)
        self.assertIn(R.OPENPYXL_NOTE, logtxt)

        wb = openpyxl.load_workbook(self.path, keep_vba=True)
        self.assertEqual(wb.sheetnames, self.order_before)        # 시트 순서 보존
        user = wb["모델"]
        self.assertEqual(user["B2"].value, "내 메모")
        self.assertEqual(user["C3"].value, 12345)
        self.assertEqual(user["C4"].value, "='rFS_별도'!E5")

        rfs = wb["rFS_별도"]
        self.assertIsNone(rfs["Z300"].value)
        rows = [(rfs.cell(row=r, column=3).value, rfs.cell(row=r, column=4).value) for r in range(1, rfs.max_row + 1)]
        self.assertIn(("§", "재무상태표"), rows)
        self.assertFalse(rfs.sheet_view.showGridLines)
        self.assertEqual(rfs.column_dimensions["D"].width, 44)
        title = next(r for r in range(1, rfs.max_row + 1) if rfs.cell(row=r, column=4).value == "재무상태표")
        self.assertTrue(rfs.cell(row=title, column=4).font.b)
        self.assertEqual(rfs.cell(row=title, column=4).font.name, "맑은 고딕")
        num = next(c for row in rfs.iter_rows(min_col=5) for c in row if isinstance(c.value, (int, float)))
        self.assertEqual(num.number_format, W.NUM_FMT)

        mfs = wb["mFS"]
        forms = [c.value for row in mfs.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")]
        self.assertTrue(forms)
        self.assertEqual(wb["목차"].freeze_panes, "A2")
        self.assertIn("cfg_status", wb.defined_names)
        self.assertTrue(str(wb["설정"]["C17"].value).startswith("업데이트 완료"))
        self.assertIn("버튼·VBA 복원·재계산·검증 안 함", str(wb["설정"]["C17"].value))

    def test_settings_cells_get_new_as_of_and_years(self):
        import dataclasses, json
        cfg2 = dataclasses.replace(self.cfg, as_of=date(2027, 1, 5), end_year=2027)
        rc, _ = self._run_main(cfg2)
        self.assertEqual(rc, 0)
        ws = openpyxl.load_workbook(self.path, keep_vba=True)["설정"]
        got = {ws.cell(row=r, column=2).value: ws.cell(row=r, column=3).value for r in range(1, 30)}
        self.assertEqual(got["기준일"].date(), date(2027, 1, 5))
        self.assertEqual(got["연도 범위"], "2024~2027")
        d = json.loads(got["설정 JSON"])
        self.assertEqual((d["as_of"], d["end_year"], d["start_year"]), ("2027-01-05", 2027, 2024))
        self.assertEqual(d["note_include"], {"recommended": True})        # 다른 키는 그대로

    def test_existing_row_order_is_read_from_file(self):
        captured = {}
        real = C.compose

        def spy(col, existing, py, script):
            captured["existing"] = existing
            return real(col, existing, py, script)

        with mock.patch.object(R.C, "compose", side_effect=spy):
            rc, _ = self._run_main()
        self.assertEqual(rc, 0)
        ex = captured["existing"]["rFS_별도"]
        self.assertIn("재무상태표", ex["blocks"])
        bs = ex["blocks"]["재무상태표"]
        self.assertEqual(bs["start"], 1)
        self.assertTrue(bs["labels"])
        self.assertNotIn("(단위 : 원)", bs["labels"])
        self.assertEqual(ex["periods"], ["FY2024", "FY2025"])           # 헤더 2행째 기말일 → 기간 키
        self.assertIn("mFS", captured["existing"])
        self.assertTrue(captured["existing"]["mFS"]["blocks"])

    def test_failure_returns_1(self):
        with mock.patch.object(R, "_collect_for", side_effect=RuntimeError("boom")), \
             mock.patch.object(R, "_attach", return_value=(None, None)), \
             mock.patch.object(R, "run_closed", side_effect=R.run_closed_openpyxl):
            self.assertEqual(R.main(["--workbook", self.path]), 1)
        with open(os.path.join(self.d, R.LOG_NAME), encoding="utf-8") as f:
            logtxt = f.read()
        self.assertIn("업데이트 실패: boom", logtxt)
        self.assertIn("Traceback", logtxt)

    def test_locked_file_failure_still_writes_log(self):
        """워크북 객체가 생기기 전 실패(잠금 파일)도 로그 파일에 남는다 — 버튼 실행은 창이 안 보이므로."""
        with open(os.path.join(self.d, "~$wb.xlsm"), "wb") as f:
            f.write(b"x")
        with mock.patch.object(R, "_attach", return_value=(None, None)):
            self.assertEqual(R.main(["--workbook", self.path]), 1)
            self.assertEqual(R.main(["--workbook", self.path]), 1)       # 두 번째 실행은 덧붙인다
        with open(os.path.join(self.d, R.LOG_NAME), encoding="utf-8") as f:
            logtxt = f.read()
        self.assertIn(R.LOCKED_MSG, logtxt)
        self.assertEqual(logtxt.count("dart_fs 업데이트:"), 2)

    def test_lock_file_refuses_closed_path_without_com(self):
        lock = os.path.join(self.d, "~$wb.xlsm")
        with open(lock, "wb") as f:
            f.write(b"x")
        import win32com.client as win32
        with mock.patch.object(win32, "DispatchEx", side_effect=AssertionError("COM 호출 금지")) as dx, \
             mock.patch.object(R, "run_closed_openpyxl", side_effect=AssertionError("대체 경로 금지")):
            with self.assertRaises(RuntimeError) as cm:
                R.run_closed(self.path, lambda m: None)
        self.assertIn("다른 Excel 인스턴스", str(cm.exception))
        dx.assert_not_called()

    def test_com_unavailable_falls_back_with_warning(self):
        import win32com.client as win32
        logs = []
        with mock.patch.object(win32, "DispatchEx", side_effect=OSError("no excel")), \
             mock.patch.object(R, "run_closed_openpyxl", return_value=dict(FAKE_VERIFY)) as fb:
            res = R.run_closed(self.path, logs.append)
        self.assertEqual(res, FAKE_VERIFY)
        fb.assert_called_once()
        self.assertTrue(any("openpyxl 경로" in m and "차트" in m for m in logs))


OLD_CHECK = "별도 검증 (TRUE 여야 함)"          # 검증 시트 신설 전 mFS 맨 아래 검증 구역 제목
OLD_A, OLD_B = "주석 12. 리스 | 옛 캡션 가", "주석 12. 리스 | 옛 캡션 나"


def _starts(ws):
    ex = C.existing_from_rows([list(r) for r in ws.iter_rows(values_only=True)])
    return {t: b["start"] for t, b in ex["blocks"].items()}


class TestOldStructureUpdate(unittest.TestCase):
    """예전 구조로 만든 파일 업데이트(openpyxl 경로): mFS 에 옛 '검증 (TRUE 여야 함)' 구역, rN 에 지금은 한 블록으로
    합쳐지는 표가 두 블록(옛 캡션 둘)으로, `검증` 시트 없음. 옛 자리는 비우고 경고, 다른 블록 시작 행은 그대로,
    새 블록은 뒤에, `검증` 시트는 새로 만든다."""
    @classmethod
    def setUpClass(cls):
        cls.cfg = P.BuildConfig("OB", "테스트", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1,
                                {"recommended": True}, "default")
        cls.col = P.collect(cls.cfg, FakeClient(), M.load(), lambda s: None)

    def setUp(self):
        self.d = tempfile.mkdtemp()
        model = C.compose(self.col, None, r"C:\py\python.exe", r"C:\app\excel\refresh.py")
        wb = openpyxl.load_workbook(W.write(model, os.path.join(self.d, "src.xlsx")))
        wb.remove(wb["검증"])
        mfs = wb["mFS"]
        r = next(r for r in range(1, mfs.max_row + 1) if mfs.cell(r, 3).value == "§" and mfs.cell(r, 4).value == "별도 검증 요약")
        mfs.cell(r, 4).value = OLD_CHECK
        self.old_check = r
        self.mfs_before = _starts(mfs)
        lease = wb["rN_리스"]
        self.merged = "주석 12. 리스 | (단위 : 천원)"
        before = _starts(lease)
        a = before.pop(self.merged)
        lease.cell(a, 4).value = OLD_A                            # 옛 파일: 같은 표가 캡션 둘로 갈려 있었다
        lease.cell(a + 6, 3).value = "§"
        lease.cell(a + 6, 4).value = OLD_B
        self.old_a, self.old_b = a, a + 6
        self.lease_before = before                                # 나머지 블록 시작 행
        self.path = os.path.join(self.d, "old.xlsm")
        wb.save(self.path)

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_old_slots_left_empty_and_new_blocks_after(self):
        with mock.patch.object(R, "_collect_for", return_value=(self.cfg, self.col, {"python": r"C:\py\python.exe"})), \
             mock.patch.object(R, "_attach", return_value=(None, None)), \
             mock.patch.object(R, "run_closed", side_effect=R.run_closed_openpyxl):
            self.assertEqual(R.main(["--workbook", self.path]), 0)
        wb = openpyxl.load_workbook(self.path, keep_vba=True)
        # 검증 시트: 새로 만들어 mFS 바로 뒤
        self.assertIn("검증", wb.sheetnames)
        self.assertEqual(wb.sheetnames.index("검증"), wb.sheetnames.index("mFS") + 1)
        chk = wb["검증"]
        a1 = next(r for r in range(1, chk.max_row + 1) if chk.cell(r, 2).value == "A1")
        self.assertTrue(str(chk.cell(a1, C.RAW_FIRST_VALUE_COL).value).startswith("=ROUND("))
        log = [c for row in wb["로그"].iter_rows(values_only=True) for c in row if isinstance(c, str)]
        # mFS: 옛 검증 구역 자리는 비우고 경고, 다른 구역은 제자리, 새 '검증 요약' 은 맨 뒤
        mfs = wb["mFS"]
        after = _starts(mfs)
        self.assertNotIn(OLD_CHECK, after)
        others = {t: s for t, s in self.mfs_before.items() if t != OLD_CHECK}
        self.assertEqual({t: after[t] for t in others}, others)
        self.assertGreater(after["별도 검증 요약"], max(others.values()))
        self.assertTrue(all(mfs.cell(r, c).value is None for r in range(self.old_check, self.old_check + 5)
                            for c in range(3, C.RAW_FIRST_VALUE_COL + 3)))
        self.assertIn(f"mFS '{OLD_CHECK}': 이번 수집에 없는 블록 — 그 자리는 비워 둠", log)
        self.assertIn(f"=COUNTIF('검증'!E", str(mfs.cell(after["별도 검증 요약"] + 2, C.RAW_FIRST_VALUE_COL).value))
        # rN_리스: 옛 두 블록 자리는 비우고 경고, 나머지 블록 제자리, 합쳐진 블록은 마지막 블록 뒤
        lease = wb["rN_리스"]
        after = _starts(lease)
        self.assertNotIn(OLD_A, after); self.assertNotIn(OLD_B, after)
        self.assertEqual({t: after[t] for t in self.lease_before}, self.lease_before)
        self.assertGreater(after[self.merged], max(self.lease_before.values()))
        self.assertTrue(all(lease.cell(r, 4).value is None for r in range(self.old_a, self.old_b + 2)))
        for t in (OLD_A, OLD_B):
            self.assertIn(f"rN_리스 '{t}': 이번 수집에 없는 블록 — 그 자리는 비워 둠", log)


class TestCollectForReader(unittest.TestCase):
    """_collect_for: 설정·목차·매핑표 읽기(네트워크 없이 — collect·DartClient 패치)."""

    def test_reads_settings_toc_and_mapping_fallback(self):
        cfg = P.BuildConfig("OB", "테스트", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, None, "default")
        import json
        sheets = {
            "설정": [[None, "dart_fs", None], [None, "mFS 표시 단위", 1000.0],
                     [None, "설정 JSON", json.dumps(cfg.to_dict(), ensure_ascii=False)]],
            "목차": [["번호", "제목", "표준 주제", "포함", "추천", "배치 시트", "표 수", "정규화 제목"],
                     ["1", "a", "", True, "", "rN_기타", 1, "a"]],
        }
        logs = []
        with mock.patch.object(R.CFG, "load", return_value={"api_key": "k", "python": "py"}), \
             mock.patch.object(R, "DartClient"), \
             mock.patch.object(R.P, "collect", return_value="COL") as coll:
            got_cfg, col, conf = R._collect_for(lambda n: sheets.get(n), logs.append)
        self.assertEqual(col, "COL")
        today = date.today()
        self.assertEqual(got_cfg.as_of, today)                       # 새 공시를 받도록 기준일 = 오늘
        self.assertEqual(got_cfg.end_year, max(2025, today.year))
        self.assertEqual(got_cfg.start_year, 2024)
        self.assertIn(f"업데이트 기준일 {today.isoformat()}, 연도 2024~{max(2025, today.year)}", logs)
        self.assertEqual(coll.call_args[0][0].as_of, today)
        self.assertEqual(got_cfg.unit_display, 1000)
        self.assertEqual(got_cfg.note_include, {"a": True})
        self.assertTrue(any("매핑표" in m for m in logs))         # 매핑표 없음 → 기본 매핑 사용 로그
        self.assertIsInstance(coll.call_args[0][2], M.Mapping)


class TestSettingsUpdates(unittest.TestCase):
    def test_rows_to_cells(self):
        import json
        cfg = P.BuildConfig("OB", "t", date(2027, 3, 2), 2023, 2027, False, ["OFS"], 1, None, "default")
        old = P.BuildConfig("OB", "t", date(2026, 10, 26), 2023, 2025, False, ["OFS"], 1000, {"a": True}, "x")
        rows = [[None, "dart_fs", None], [None, None, None], [None, "기준일", date(2026, 10, 26)],
                [None, "연도 범위", "2023~2025"], [None, "설정 JSON", json.dumps(old.to_dict(), ensure_ascii=False)],
                [None, "상태", "x"]]
        got = R.settings_updates(rows, cfg)
        self.assertEqual(got[:2], [(3, 3, date(2027, 3, 2)), (4, 3, "2023~2027")])
        r, c, js = got[2]
        d = json.loads(js)
        self.assertEqual((r, c, d["as_of"], d["end_year"], d["unit_display"], d["note_include"]),
                         (5, 3, "2027-03-02", 2027, 1000, {"a": True}))
        self.assertEqual(R.settings_updates([[None, "설정 JSON", "{broken"]], cfg), [])


if __name__ == "__main__":
    unittest.main()
