import json, os, threading, unittest, urllib.error, urllib.request
from unittest import mock
import server as SV

class _ServerCase(unittest.TestCase):
    """스레드 서버 한 대를 빈 포트에 띄우고 요청 헬퍼를 제공한다(테스트 메서드는 하위 클래스에)."""
    @classmethod
    def setUpClass(cls):
        cls.srv = SV.make_server(0)
        cls.port = cls.srv.server_address[1]
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown(); cls.srv.server_close()

    def _get(self, path):
        with urllib.request.urlopen(f"http://127.0.0.1:{self.port}{path}") as r:
            return r.status, r.read()

    def _post(self, path, body, headers=None):
        h = {"Content-Type": "application/json"}; h.update(headers or {})
        req = urllib.request.Request(f"http://127.0.0.1:{self.port}{path}", data=json.dumps(body).encode(), headers=h)
        with urllib.request.urlopen(req) as r:
            return r.status, json.loads(r.read())

    def _status(self, path, body=None, headers=None, raw_headers=None):
        h = dict(raw_headers) if raw_headers is not None else {"Content-Type": "application/json"}
        h.update(headers or {})
        req = urllib.request.Request(f"http://127.0.0.1:{self.port}{path}", data=None if body is None else json.dumps(body).encode(), headers=h)
        try:
            with urllib.request.urlopen(req) as r:
                return r.status
        except urllib.error.HTTPError as e:
            return e.code


class TestServer(_ServerCase):
    def test_index_and_config(self):
        st, body = self._get("/")
        self.assertEqual(st, 200); self.assertIn(b"dart_fs", body)
        st, body = self._get("/api/config")
        self.assertEqual(set(json.loads(body)), {"has_key", "python", "server_started", "code_version"})

    def test_config_has_version_info(self):
        st, body = self._get("/api/config")
        d = json.loads(body)
        for k in ("server_started", "code_version"):
            self.assertIsInstance(d[k], str); self.assertTrue(d[k])

    def test_search_uses_client(self):
        fake = mock.Mock(); fake.search.return_value = [{"corp_code": "1", "corp_name": "오비맥주", "stock_code": ""}]
        with mock.patch.object(SV, "get_client", return_value=fake):
            st, body = self._get("/api/search?q=%EC%98%A4%EB%B9%84")
        self.assertEqual(json.loads(body)[0]["corp_name"], "오비맥주")

    def test_plan_endpoint(self):
        fake = mock.Mock()
        fake.list_disclosures.return_value = [{"rcept_no": "A", "report_nm": "감사보고서 (2025.12)", "rcept_dt": "20260410"}]
        with mock.patch.object(SV, "get_client", return_value=fake):
            st, body = self._post("/api/plan", {"corp_code": "1", "corp_name": "x", "as_of": "2026-10-26", "start_year": 2025, "end_year": 2025, "include_half": False, "fs_divs": ["OFS"]})
        self.assertEqual(body["periods"], ["FY2025"]); self.assertEqual(body["reports"][0]["rcept_no"], "A"); self.assertEqual(body["base_report"], "A")

    def test_build_job_streams_events(self):
        with mock.patch.object(SV, "run_build", side_effect=lambda cfg, log: (log("하나"), {"path": "p.xlsm", "verify": {"errors": 0, "checks": {"TRUE": 1, "FALSE": 0, "해당 없음": 0}, "false_cells": []}})[1]):
            st, body = self._post("/api/build", {"corp_code": "1", "corp_name": "x", "as_of": "2026-10-26", "start_year": 2025, "end_year": 2025, "include_half": False, "fs_divs": ["OFS"], "note_include": {"recommended": True}})
            job = body["job"]
            with urllib.request.urlopen(f"http://127.0.0.1:{self.port}/api/events/{job}") as r:
                text = r.read().decode("utf-8")
        self.assertIn('"type": "log"', text); self.assertIn('"type": "done"', text); self.assertIn("p.xlsm", text)

    def test_build_passes_auto_unit(self):
        seen = {}

        def fake(cfg, log):
            seen["cfg"] = cfg
            return {"path": "p.xlsm", "verify": {"errors": 0, "checks": {"TRUE": 0, "FALSE": 0, "해당 없음": 0}, "false_cells": []}}
        with mock.patch.object(SV, "run_build", side_effect=fake):
            st, body = self._post("/api/build", {"corp_code": "1", "corp_name": "x", "as_of": "2026-10-26", "start_year": 2025,
                                                 "end_year": 2025, "include_half": False, "fs_divs": ["OFS"], "unit_display": 0})
            with urllib.request.urlopen(f"http://127.0.0.1:{self.port}/api/events/{body['job']}") as r:
                r.read()
        self.assertEqual(st, 200)
        self.assertEqual(seen["cfg"]["unit_display"], 0)
        self.assertEqual(SV.P.BuildConfig.from_dict(seen["cfg"]).unit_display, 0)

    def test_ui_unit_default_is_auto(self):
        st, body = self._get("/")
        html = body.decode("utf-8")
        self.assertIn('<input type="radio" name="unit" value="0" checked>자동(공시 단위)', html)
        self.assertNotIn('name="unit" value="1" checked', html)

    def test_presets_roundtrip(self):
        with mock.patch.object(SV, "PRESET_DIR", SV.tempfile.mkdtemp()):
            self._post("/api/presets", {"name": "t1", "note_include": {"유형자산": True}})
            st, body = self._get("/api/presets")
        self.assertEqual(json.loads(body)["t1"], {"유형자산": True})

    def test_guard_bad_origin_host_ctype(self):
        self.assertEqual(self._status("/api/presets", {"name": "x", "note_include": {}}, {"Origin": "http://evil.example"}), 403)
        self.assertEqual(self._status("/api/config", {"api_key": ""}, raw_headers={"Content-Type": "text/plain"}), 403)
        self.assertEqual(self._status("/api/config", headers={"Host": "evil.example"}), 403)
        self.assertEqual(self._status("/api/config", headers={"Origin": "http://evil.example"}), 403)
        self.assertEqual(self._status("/api/config", headers={"Origin": f"http://localhost:{self.port}", "Host": f"localhost:{self.port}"}), 200)

    def test_config_ignores_python(self):
        with mock.patch.object(SV.CFG, "save") as sv:
            self._post("/api/config", {"api_key": " k ", "python": "evil.exe"})
        sv.assert_called_once_with("k")

    def test_open_folder_scope(self):
        with mock.patch.object(SV.subprocess, "Popen") as po:
            self.assertEqual(self._status("/api/open-folder", {"path": "C:\\Windows"}), 403)
            po.assert_not_called()
            os.makedirs(SV.CFG.OUTPUT_DIR, exist_ok=True)
            st, _ = self._post("/api/open-folder", {"path": SV.CFG.OUTPUT_DIR})
            self.assertEqual(st, 200); po.assert_called_once()

    def test_open_folder_company_subfolder(self):
        root = SV.tempfile.mkdtemp()
        sub = os.path.join(root, "output", "샘플전자")
        os.makedirs(sub)
        evil = os.path.join(root, "output_evil")                      # 접두가 같은 형제 폴더는 거부
        os.makedirs(evil)
        with mock.patch.object(SV.CFG, "OUTPUT_DIR", os.path.join(root, "output")), \
                mock.patch.object(SV.subprocess, "Popen") as po:
            st, _ = self._post("/api/open-folder", {"path": os.path.join(sub, "샘플전자_2026-10-27.xlsm")})
            self.assertEqual(st, 200)
            self.assertEqual(po.call_args.args[0], ["explorer", sub])
            self.assertEqual(self._status("/api/open-folder", {"path": evil}), 403)
            self.assertEqual(self._status("/api/open-folder", {"path": os.path.join(sub, "..", "..")}), 403)
            self.assertEqual(po.call_count, 1)

    def test_open_folder_job_result(self):
        d = SV.tempfile.mkdtemp()
        SV._jobs["zz"] = {"q": None, "history": [], "done": True, "result": {"type": "done", "result": {"path": os.path.join(d, "a.xlsm")}}}
        try:
            with mock.patch.object(SV.subprocess, "Popen") as po:
                st, _ = self._post("/api/open-folder", {"path": os.path.join(d, "a.xlsm")})
                self.assertEqual(st, 200); po.assert_called_once()
        finally:
            SV._jobs.pop("zz", None)

    def test_preset_name_sanitized(self):
        d = SV.tempfile.mkdtemp()
        with mock.patch.object(SV, "PRESET_DIR", d):
            self._post("/api/presets", {"name": "a:b/../c", "note_include": {}})
            files = os.listdir(d)
        self.assertEqual(files, ["a_b____c.json"])

    def test_unknown_job_404(self):
        self.assertEqual(self._status("/api/job/nope"), 404)


class TestDartOnly(_ServerCase):
    """회사 제시 엑셀 입력(파일로 생성)은 보류 — 엔드포인트·화면 탭이 남아 있지 않아야 한다."""
    def test_file_endpoints_removed(self):
        for path in ("/api/file/scan", "/api/file/build", "/api/pick-files"):
            self.assertEqual(self._status(path, {}), 404, path)

    def test_ui_v2_has_no_file_tab(self):
        with open(os.path.join(SV.HERE, "web", "index_v2.html"), encoding="utf-8") as fh:
            html = fh.read()
        for s in ('data-tab="file"', 'id="tab-file"', "/api/file/", "/api/pick-files", "파일로 생성"):
            self.assertNotIn(s, html)
        self.assertEqual(html.count("<script>"), html.count("</script>"))


if __name__ == "__main__":
    unittest.main()
