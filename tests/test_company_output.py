"""Company output destinations; no network or Excel process is needed."""
import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from excel import build as B, finalize as FZ


class TestCompanyOutput(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.cfg = {
            "corp_code": "00126380", "corp_name": "삼성전자",
            "as_of": "2026-10-27", "start_year": 2024,
            "end_year": 2025, "include_half": False, "fs_divs": ["OFS"],
        }

    def build(self, **kwargs):
        model = mock.Mock(unit_text="원")
        collection = mock.Mock(logs=[])
        with mock.patch.object(B.CFG, "OUTPUT_DIR", str(self.root)), \
                mock.patch.object(B.CFG, "load", return_value={"python": "python"}), \
                mock.patch.object(B.M, "load"), \
                mock.patch.object(B.P, "collect", return_value=collection) as collect, \
                mock.patch.object(B.C, "compose", return_value=model), \
                mock.patch.object(B.W, "write") as write, \
                mock.patch.object(FZ, "finalize", return_value={
                    "errors": 0, "checks": {"TRUE": 0, "FALSE": 0},
                }) as finalize:
            result = B.build(self.cfg, lambda msg: None, client=mock.Mock(), **kwargs)
        return result, write, finalize, collect

    def test_default_creates_company_folder_and_preserves_existing_output(self):
        old = self.root / "삼성전자_2026-10-26.xlsm"
        old.write_bytes(b"existing-output")
        result, write, finalize, _ = self.build()
        folder = self.root / "삼성전자"
        self.assertTrue(folder.is_dir())
        self.assertEqual(result["path"], str(folder / "삼성전자_2026-10-27.xlsm"))
        self.assertEqual(write.call_args.args[1], str(folder / "삼성전자_2026-10-27.xlsx"))
        self.assertEqual(finalize.call_args.args[:2], (
            str(folder / "삼성전자_2026-10-27.xlsx"), result["path"],
        ))
        self.assertEqual(old.read_bytes(), b"existing-output")

    def test_different_companies_have_separate_folders(self):
        first, _, _, _ = self.build()
        self.cfg["corp_name"] = "휴젤"
        second, _, _, _ = self.build()
        self.assertEqual(Path(first["path"]).parent, self.root / "삼성전자")
        self.assertEqual(Path(second["path"]).parent, self.root / "휴젤")

    def test_explicit_output_directory_is_preserved(self):
        target = self.root / "custom"
        result, _, _, _ = self.build(out_dir=str(target))
        self.assertEqual(Path(result["path"]).parent, target)

    def test_company_folder_cannot_escape_output_root(self):
        for name in ("..", "../회사", "테스트/회사"):
            with self.subTest(name=name):
                self.cfg["corp_name"] = name
                result, _, _, _ = self.build()
                folder = Path(result["path"]).parent.resolve()
                self.assertEqual(os.path.commonpath([str(folder), str(self.root)]), str(self.root))
                self.assertNotEqual(folder, self.root)

    def test_company_folder_lock_refuses_before_collection(self):
        folder = self.root / "삼성전자"
        folder.mkdir()
        (folder / "~$삼성전자_2026-10-27.xlsm").write_bytes(b"lock")
        with mock.patch.object(B.CFG, "OUTPUT_DIR", str(self.root)), \
                mock.patch.object(B.CFG, "load", return_value={"python": "python"}), \
                mock.patch.object(B.P, "collect") as collect:
            with self.assertRaisesRegex(RuntimeError, "출력 파일이 Excel에 열려 있습니다"):
                B.build(self.cfg, lambda msg: None, client=mock.Mock())
            collect.assert_not_called()


if __name__ == "__main__":
    unittest.main()
