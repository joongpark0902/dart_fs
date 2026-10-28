import os, sys, tempfile, unittest
from unittest import mock
from core import config as CFG


class TestConfig(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = os.path.join(self.tmp.name, "config.txt")
        p = mock.patch.object(CFG, "CONFIG_PATH", self.path)
        p.start(); self.addCleanup(p.stop)

    def tearDown(self):
        self.tmp.cleanup()

    def _read(self):
        with open(self.path, encoding="utf-8") as f:
            return f.read()

    def test_save_key_only_has_no_python_line(self):
        CFG.save("ABC")
        txt = self._read()
        self.assertIn("api_key = ABC", txt); self.assertNotIn("python", txt)

    def test_save_with_python(self):
        CFG.save("ABC", python="C:/py/python.exe")
        txt = self._read()
        self.assertIn("api_key = ABC", txt); self.assertIn("python = C:/py/python.exe", txt)

    def test_existing_python_and_extra_lines_preserved(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("api_key = OLD\npython = C:/p/python.exe\nfoo = bar\n")
        CFG.save("NEW")
        txt = self._read()
        self.assertIn("api_key = NEW", txt); self.assertIn("python = C:/p/python.exe", txt); self.assertIn("foo = bar", txt)
        self.assertNotIn("OLD", txt)

    def test_quoted_value_stripped(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write('api_key = "QQQ"\npython = \'C:/x/python.exe\'\n')
        c = CFG.load()
        self.assertEqual(c["api_key"], "QQQ"); self.assertEqual(c["python"], "C:/x/python.exe")

    def test_python_with_control_char_falls_back(self):
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("api_key = K\npython = C:\\Users\x07user\\python.exe\n")
        c = CFG.load()
        self.assertEqual(c["python"], sys.executable); self.assertEqual(c["api_key"], "K")


if __name__ == "__main__":
    unittest.main()
