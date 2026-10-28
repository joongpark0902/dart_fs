"""파일 입력 build·refresh(F3): 가상 '샘플전자' 엑셀 → tests/_out/file_build/샘플전자/…
openpyxl 단계까지는 COM 없이, finalize(xlsm) 는 DART_FS_COM_TESTS=1 일 때만."""
import json
import os
import shutil
import tempfile
import unittest
from datetime import date
from unittest import mock

import openpyxl

from core import pipeline as P
from excel import build as B, finalize as FZ, refresh as R
from tests.fixtures import xlsx_fixture as F

COM = os.environ.get("DART_FS_COM_TESTS") == "1"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_out", "file_build")
INPUTS = os.path.join(OUT, "inputs")
FAKE_VERIFY = {"errors": 0, "checks": {"TRUE": 0, "FALSE": 0, "해당 없음": 0}, "false_cells": []}
NO_DART = mock.Mock(side_effect=AssertionError("파일 입력에서 DART 클라이언트를 만들면 안 됨"))


def _cfg_dict(files, **kw):
    d = P.BuildConfig(corp_code="FILE:샘플전자", corp_name="샘플전자", as_of=date(2026, 10, 27), start_year=2024,
                      end_year=2025, include_half=False, fs_divs=["OFS"], unit_display=1, note_include=None,
                      preset="default", source="file", files=files, layout=None).to_dict()
    d.update(kw)
    return json.loads(json.dumps(d, ensure_ascii=False))


def _file(path, period_end):
    return {"path": path, "period_end": period_end, "kind": "annual", "fs_div": "OFS", "unit": None}


def _inputs(d):
    os.makedirs(d, exist_ok=True)
    st24 = F.build_statements(os.path.join(d, "statements_2024.xlsx"), year=2024)
    st25 = F.build_statements(os.path.join(d, "statements_2025.xlsx"))
    nt25 = F.build_notes(os.path.join(d, "notes_2025.xlsx"))
    return [_file(st24, "2024-12-31"), _file(st25, "2025-12-31"), _file(nt25, "2025-12-31")]


def _settings(ws):
    return {ws.cell(row=r, column=2).value: ws.cell(row=r, column=3) for r in range(1, 40)}


class TestFileBuild(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        shutil.rmtree(OUT, ignore_errors=True)
        cls.files = _inputs(INPUTS)
        cls.logs = []
        with mock.patch.object(B.CFG, "OUTPUT_DIR", OUT), \
                mock.patch.object(B.CFG, "load", return_value={"python": r"C:\py\python.exe", "api_key": ""}), \
                mock.patch("core.fetch.DartClient", NO_DART), \
                mock.patch.object(B.P, "collect", side_effect=AssertionError("DART collect 금지")), \
                mock.patch.object(FZ, "finalize", return_value=dict(FAKE_VERIFY)) as fin:
            cls.res = B.build(_cfg_dict(cls.files), cls.logs.append)
        cls.fin_args = fin.call_args.args

    def test_company_folder_and_paths(self):
        folder = os.path.join(OUT, "샘플전자")
        self.assertEqual(self.res["path"], os.path.join(folder, "샘플전자_2026-10-27.xlsm"))
        self.assertEqual(self.fin_args[:2], (os.path.join(folder, "샘플전자_2026-10-27.xlsx"), self.res["path"]))
        self.assertTrue(os.path.isfile(self.fin_args[0]))
        self.assertTrue(any("파일 statements_2025.xlsx 시트 재무상태표 표 1: BS" in m for m in self.res["logs"]))

    def test_settings_and_log_sheets(self):
        wb = openpyxl.load_workbook(self.fin_args[0])
        st = _settings(wb["설정"])
        self.assertEqual(st["corp_code"].value, "파일 입력")
        self.assertEqual(st["입력 파일"].value.split("\n"), [f["path"] for f in self.files])
        self.assertTrue(st["입력 파일"].alignment.wrap_text)
        self.assertEqual(st["입력 파일"].row, st["연결/별도"].row + 1)
        js = json.loads(st["설정 JSON"].value)
        self.assertEqual(js["source"], "file")
        for name in ("cfg_json", "cfg_status", "cfg_unit"):
            self.assertIn(name, wb.defined_names)
        dest = dict(wb.defined_names["cfg_json"].destinations)
        self.assertEqual(dest["설정"].replace("$", ""), st["설정 JSON"].coordinate)
        log = [list(r) for r in wb["로그"].iter_rows(values_only=True)]
        self.assertIn(["기간말", "파일", "시트(표)", "수정일", "단위"], [r[:5] for r in log])
        self.assertFalse(any(r[0] == "감사인·의견" for r in log))
        row = next(r for r in log if r[1] == "notes_2025.xlsx")
        self.assertIn("주석(표", row[2])
        self.assertEqual(row[4], "원")

    def test_missing_input_file_fails_before_output(self):
        with tempfile.TemporaryDirectory() as root:
            files = [_file(os.path.join(root, "없음.xlsx"), "2025-12-31")]
            with mock.patch.object(B.CFG, "OUTPUT_DIR", root), \
                    mock.patch.object(B.CFG, "load", return_value={"python": "py"}), \
                    mock.patch("excel.build.M.load") as ml, \
                    mock.patch.object(FZ, "finalize") as fin:
                with self.assertRaises(RuntimeError) as cm:
                    B.build(_cfg_dict(files), lambda m: None)
            self.assertIn(f"입력 파일을 찾을 수 없습니다: {files[0]['path']}", str(cm.exception))
            self.assertEqual(os.listdir(root), [])                  # 회사 폴더도 만들지 않음
            ml.assert_not_called(); fin.assert_not_called()

    def test_empty_path_entry_fails(self):
        with mock.patch.object(B.CFG, "load", return_value={"python": "py"}):
            with self.assertRaisesRegex(RuntimeError, "입력 파일 경로가 비어 있습니다"):
                B.build(_cfg_dict([_file("", "2025-12-31")]), lambda m: None)

    def test_empty_file_list_fails(self):
        with mock.patch.object(B.CFG, "load", return_value={"python": "py"}):
            with self.assertRaisesRegex(RuntimeError, "입력 파일이 지정되지 않았습니다"):
                B.build(_cfg_dict([]), lambda m: None)


@unittest.skipUnless(COM, "DART_FS_COM_TESTS=1 일 때만(Excel COM 필요)")
class TestFileBuildCom(unittest.TestCase):
    def test_finalize_xlsm(self):
        out = os.path.join(OUT, "com")
        shutil.rmtree(out, ignore_errors=True)
        files = _inputs(os.path.join(out, "inputs"))
        with mock.patch.object(B.CFG, "OUTPUT_DIR", out), \
                mock.patch.object(B.CFG, "load", return_value={"python": r"C:\py\python.exe", "api_key": ""}), \
                mock.patch("core.fetch.DartClient", NO_DART):
            res = B.build(_cfg_dict(files), lambda m: None)
        self.assertEqual(os.path.dirname(res["path"]), os.path.join(out, "샘플전자"))
        self.assertTrue(os.path.isfile(res["path"]))
        self.assertEqual(res["verify"]["errors"], 0)
        self.assertEqual(res["verify"]["checks"]["FALSE"], 0)


class TestFileRefreshClosed(unittest.TestCase):
    """닫힌 파일 openpyxl 대체 경로로 파일 기반 통합문서를 다시 읽는다(api_key 없음, DART 호출 금지)."""

    def setUp(self):
        self.d = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.d, True)
        self.files = _inputs(os.path.join(self.d, "inputs"))
        out = os.path.join(self.d, "output")
        with mock.patch.object(B.CFG, "OUTPUT_DIR", out), \
                mock.patch.object(B.CFG, "load", return_value={"python": "py", "api_key": ""}), \
                mock.patch.object(FZ, "finalize", return_value=dict(FAKE_VERIFY)) as fin:
            B.build(_cfg_dict(self.files), lambda m: None)
        wb = openpyxl.load_workbook(fin.call_args.args[0])
        self.path = os.path.join(out, "샘플전자", "wb.xlsm")
        wb.save(self.path)
        self.mapdir = os.path.join(self.d, "mapping_user")

    def _refresh(self):
        logs = []
        with mock.patch.object(R.CFG, "load", return_value={"python": "py", "api_key": ""}), \
                mock.patch.object(R, "DartClient", NO_DART), \
                mock.patch.object(R.P, "collect", side_effect=AssertionError("DART collect 금지")), \
                mock.patch.object(R, "MAPPING_USER_DIR", self.mapdir):
            R.run_closed_openpyxl(self.path, logs.append)
        return logs

    def test_rereads_input_files(self):
        src = self.files[1]["path"]                                      # 2025 재무제표: 자산총계 1500M → 1501M
        wb = openpyxl.load_workbook(src)
        ws = wb["재무상태표"]
        cell = next(c for row in ws.iter_rows() for c in row if c.value == 1500 * F.M)
        cell.value = 1501 * F.M
        wb.save(src)
        logs = self._refresh()
        self.assertIn("파일 입력 3개 다시 읽는 중…", logs)
        wb = openpyxl.load_workbook(self.path, keep_vba=True)
        vals = [c for row in wb["rFS_별도"].iter_rows(values_only=True) for c in row]
        self.assertIn(1501 * F.M, vals)
        st = _settings(wb["설정"])
        self.assertEqual(st["corp_code"].value, "파일 입력")
        self.assertTrue(str(st["상태"].value).startswith("업데이트 완료"))
        js = json.loads(st["설정 JSON"].value)
        self.assertEqual(js["source"], "file")
        self.assertEqual(st["연도 범위"].value, "2024~2025")                  # 파일 입력은 끝 연도를 올리지 않음
        self.assertEqual((js["start_year"], js["end_year"], js["as_of"]), (2024, 2025, date.today().isoformat()))
        self.assertTrue(any(f"연도 2024~2025" in m for m in logs))
        log =[list(r) for r in wb["로그"].iter_rows(values_only=True)]
        self.assertIn(["기간말", "파일", "시트(표)", "수정일", "단위"], [r[:5] for r in log])

    def test_missing_input_file_no_write(self):
        os.remove(self.files[2]["path"])
        def data():
            with open(self.path, "rb") as fh:
                return fh.read()
        before = data()
        with self.assertRaises(R.InputFileError) as cm:
            self._refresh()
        self.assertIn("입력 파일을 찾을 수 없습니다", str(cm.exception))
        self.assertEqual(data(), before)                                 # 부분 쓰기 없음

    def test_missing_input_file_com_path_sets_status(self):
        """COM 경로(_refresh_workbook): 데이터 시트는 쓰지 않고 상태 셀 기록, 닫힌 파일이면 저장."""
        os.remove(self.files[2]["path"])
        app, wb = mock.Mock(), mock.Mock()
        xl = openpyxl.load_workbook(self.path)

        def reader(_wb):
            return lambda n: [list(r) for r in xl[n].iter_rows(values_only=True)] if n in xl.sheetnames else None
        with mock.patch.object(R.WC, "read_existing", return_value={}), \
                mock.patch.object(R, "_com_reader", side_effect=reader), \
                mock.patch.object(R.WC, "write_model") as wm, \
                mock.patch.object(R, "set_status_com") as ss, \
                mock.patch.object(R.CFG, "load", return_value={"python": "py", "api_key": ""}), \
                mock.patch.object(R, "MAPPING_USER_DIR", self.mapdir):
            with self.assertRaises(R.InputFileError):
                R._refresh_workbook(app, wb, lambda m: None, save=True)
        wm.assert_not_called()
        self.assertIn("입력 파일을 찾을 수 없습니다", ss.call_args.args[1])
        self.assertIn("데이터 시트는 바뀌지 않았습니다", ss.call_args.args[1])
        wb.Save.assert_called_once()

    def test_open_path_input_error_keeps_status_message(self):
        """열린 통합문서: InputFileError 면 main 이 '저장하지 말고 닫으세요' 로 상태 셀을 덮지 않는다."""
        app, wb = mock.Mock(), mock.Mock()
        with mock.patch.object(R, "_attach", return_value=(app, wb)), \
                mock.patch.object(R, "run_open", side_effect=R.InputFileError("입력 파일을 찾을 수 없습니다: x")), \
                mock.patch.object(R, "set_status_com") as ss:
            self.assertEqual(R.main(["--workbook", self.path]), 1)
        ss.assert_not_called()
        with mock.patch.object(R, "_attach", return_value=(app, wb)), \
                mock.patch.object(R, "run_open", side_effect=RuntimeError("boom")), \
                mock.patch.object(R, "set_status_com") as ss:
            self.assertEqual(R.main(["--workbook", self.path]), 1)
        self.assertIn("저장하지 말고 닫으세요", ss.call_args.args[1])           # 다른 오류는 그대로


if __name__ == "__main__":
    unittest.main()
