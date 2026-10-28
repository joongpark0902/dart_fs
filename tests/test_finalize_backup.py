import os, shutil, tempfile, unittest
from datetime import datetime


class TestBackupExisting(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def test_none_when_missing(self):
        from excel import finalize as FZ
        self.assertIsNone(FZ.backup_existing(os.path.join(self.d, "x.xlsm")))

    def test_moves_with_mtime_name_and_suffix(self):
        from excel import finalize as FZ
        p = os.path.join(self.d, "x.xlsm")
        open(p, "w").write("a")
        ts = datetime(2026, 1, 2, 3, 4, 5).timestamp()
        os.utime(p, (ts, ts))
        b = FZ.backup_existing(p)
        self.assertFalse(os.path.exists(p))
        self.assertEqual(os.path.basename(b), "backup_20260102-030405_rebuild.xlsm")
        self.assertEqual(open(b).read(), "a")
        open(p, "w").write("b")
        os.utime(p, (ts, ts))
        b2 = FZ.backup_existing(p)
        self.assertEqual(os.path.basename(b2), "backup_20260102-030405_rebuild_2.xlsm")
        self.assertEqual(open(b).read(), "a")

    def test_permission_error_becomes_korean_runtime_error(self):
        from unittest import mock
        from excel import finalize as FZ
        p = os.path.join(self.d, "x.xlsm")
        open(p, "w").write("a")
        with mock.patch.object(FZ.shutil, "move", side_effect=PermissionError(13, "denied")):
            with self.assertRaises(RuntimeError) as cm:
                FZ.backup_existing(p)
        self.assertIn("열려 있음", str(cm.exception))
        self.assertIsInstance(cm.exception.__cause__, PermissionError)
        self.assertTrue(os.path.exists(p))


class TestBuildLock(unittest.TestCase):
    def test_lock_file_refuses_before_collect(self):
        from unittest import mock
        from excel import build as B
        d = tempfile.mkdtemp()
        try:
            cfg = {"corp_code": "OB", "corp_name": "테스트 회사", "as_of": "2026-10-26", "start_year": 2024,
                   "end_year": 2025, "include_half": False, "fs_divs": ["OFS"]}
            open(os.path.join(d, "~$테스트_회사_2026-10-26.xlsm"), "w").write("x")
            client = mock.Mock()
            with mock.patch.object(B.P, "collect") as coll:
                with self.assertRaises(RuntimeError) as cm:
                    B.build(cfg, lambda m: None, client=client, out_dir=d)
            self.assertIn("출력 파일이 Excel에 열려 있습니다", str(cm.exception))
            coll.assert_not_called()
            client.list_disclosures.assert_not_called()
        finally:
            shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
