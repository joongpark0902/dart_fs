"""dart_fs 브라우저 UI 서버. 표준 라이브러리만. python server.py [--port 8771] [--no-browser]"""
import argparse, datetime, json, os, queue, re, subprocess, sys, tempfile, threading, uuid, webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
from core import config as CFG, mapping as M, pipeline as P, selector as SEL     # noqa: E402
from core import file_collect as FC, normalize as N                               # noqa: E402
from core.fetch import DartClient                                                 # noqa: E402
from core.note_std import sheet_for                                               # noqa: E402
from excel.build import build as run_build_impl                                   # noqa: E402

FILE_EXTS = (".xlsx", ".xlsm", ".xls")

PRESET_DIR = CFG.PRESET_DIR
_client = None
def _code_version():
    try:
        r = subprocess.run(["git", "rev-parse", "--short", "HEAD"], cwd=HERE, capture_output=True, text=True, timeout=3)
        v = r.stdout.strip()
        if r.returncode == 0 and v:
            return v
    except Exception:
        pass
    try:
        return datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(HERE, "core", "pipeline.py"))).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return "unknown"


SERVER_STARTED = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
CODE_VERSION = _code_version()
_jobs = {}            # id -> {"q": Queue, "history": [], "done": bool, "result": ..}


def get_client():
    global _client
    conf = CFG.load()
    if not conf["api_key"]:
        raise RuntimeError("인증키가 없습니다. 화면 상단에서 입력하세요.")
    if _client is None or _client.api_key != conf["api_key"]:
        _client = DartClient(conf["api_key"], CFG.CACHE_DIR)
    return _client


def run_build(cfg, log):
    return run_build_impl(cfg, log, client=get_client())


def run_file_build(cfg, log):
    """파일 입력 생성: DART 클라이언트·인증키 없이 build.build(source=file)."""
    return run_build_impl(cfg, log)


# ---------------------------------------------------------------- 파일 입력(F4)

class BadRequest(ValueError):
    pass


def _check_input_path(p):
    """입력 엑셀 경로 검사 → 오류 문구(없으면 None). 절대 경로·존재·확장자만 본다(파일 내용은 읽지 않는다)."""
    p = str(p or "").strip()
    if not p:
        return "경로가 비어 있습니다"
    if not os.path.isabs(p):
        return "절대 경로만 받습니다"
    if not os.path.isfile(p):
        return "파일을 찾을 수 없습니다"
    ext = os.path.splitext(p)[1].lower()
    if ext not in FILE_EXTS:
        return "xlsx·xlsm 파일만 받습니다"
    if ext == ".xls":
        return "xls 는 아직 지원하지 않습니다 — Excel 에서 xlsx 로 저장한 뒤 넣어 주세요"
    return None


def _file_scan_json(body):
    """POST /api/file/scan {paths:[…]} 또는 {files:[{path, period_end?, kind?, fs_div?, unit?}], layout?}
    → {files:[preview…], errors:[{path, error}]}. 같은 파일명이 둘이면 key_base 를 'a.xlsx#2' 로(collect_files 와 같은 규칙)."""
    files = body.get("files")
    if files is None:
        paths = body.get("paths")
        if not isinstance(paths, list):
            raise BadRequest("paths 또는 files 목록이 필요합니다")
        files = [{"path": p} for p in paths]
    if not isinstance(files, list) or not all(isinstance(d, dict) for d in files):
        raise BadRequest("files 는 {path, …} 목록이어야 합니다")
    layout = body.get("layout") or {}
    if not isinstance(layout, dict):
        raise BadRequest("layout 은 사전이어야 합니다")
    mapping = M.load()
    out, errors, seen = [], [], {}
    for d in files:
        path = str(d.get("path") or "").strip()
        err = _check_input_path(path)
        if err:
            errors.append({"path": path, "error": err})
            continue
        base = os.path.basename(path)
        n = seen[base] = seen.get(base, 0) + 1
        key_base = base + (f"#{n}" if n > 1 else "")
        try:
            out.append(FC.preview(path, period_end=d.get("period_end") or None, kind=d.get("kind") or None,
                                  fs_div=d.get("fs_div") or "OFS", unit=d.get("unit"), key_base=key_base, layout=layout, mapping=mapping))
        except Exception as e:     # noqa: BLE001 — 파일 내용은 담지 않는다
            errors.append({"path": path, "error": f"읽지 못함: {type(e).__name__}: {str(e)[:200]}"})
    return {"files": out, "errors": errors}


def _check_file_cfg(cfg):
    """POST /api/file/build 의 cfg 검사(job 시작 전): source=file, 회사명, 파일 목록(절대 경로·존재·결산일)."""
    if not isinstance(cfg, dict):
        raise BadRequest("cfg 가 필요합니다")
    cfg = dict(cfg)
    cfg["source"] = "file"
    name = str(cfg.get("corp_name") or "").strip()
    if not name:
        raise BadRequest("회사명을 입력하세요")
    cfg["corp_name"] = name
    cfg["corp_code"] = "FILE:" + name          # 클라이언트 corp_code 는 받지 않는다(DART 코드로 오인 방지)
    files = cfg.get("files")
    if not isinstance(files, list) or not files:
        raise BadRequest("입력 파일이 없습니다")
    for d in files:
        if not isinstance(d, dict):
            raise BadRequest("files 항목은 {path, period_end, …} 이어야 합니다")
        err = _check_input_path(d.get("path"))
        if err:
            raise BadRequest(f"{os.path.basename(str(d.get('path') or ''))}: {err}")
        pe = str(d.get("period_end") or "")[:10]
        try:
            datetime.date.fromisoformat(pe)
        except ValueError:
            raise BadRequest(f"{os.path.basename(str(d['path']))}: 결산일을 입력하세요") from None
        if d.get("kind") not in (None, "", "annual", "half"):
            raise BadRequest(f"{os.path.basename(str(d['path']))}: 연간/반기 값이 잘못됐습니다")
    return cfg


_pick_lock = threading.Lock()
_pick_thread = None     # 마지막 파일 선택 창 스레드 — 시간 초과 뒤에도 창이 살아 있으면 새 Tk() 를 만들지 않는다
PICK_TIMEOUT = 600      # 초. 창이 안 닫혀도 요청·잠금이 영원히 묶이지 않게


def _ask_open_filenames():
    """서버 PC 에서 tkinter 파일 선택 창(xlsx/xlsm/xls). 별도 스레드에서 부른다. 실패하면 예외."""
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    try:
        root.withdraw()
        root.attributes("-topmost", True)
        root.update()
        return filedialog.askopenfilenames(parent=root, title="재무제표 엑셀 선택",
                                           filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")])
    finally:
        root.destroy()


def _pick_files():
    """POST /api/pick-files → {paths:[…]} 또는 {paths:[], error:"…"}. tkinter·화면이 없어도 서버는 죽지 않는다."""
    global _pick_thread
    if not _pick_lock.acquire(blocking=False):
        return {"paths": [], "error": "파일 선택 창이 이미 열려 있습니다"}
    out = {"paths": []}

    def run():
        try:
            paths = _ask_open_filenames() or ()
            if isinstance(paths, str):
                paths = [paths]
            out["paths"] = [os.path.abspath(p) for p in paths if p]
        except Exception as e:      # noqa: BLE001
            out["error"] = f"파일 선택 창을 열 수 없습니다: {type(e).__name__}: {str(e)[:200]} — 경로를 직접 붙여 넣으세요"
    try:
        if _pick_thread is not None and _pick_thread.is_alive():      # 시간 초과된 이전 창이 아직 열려 있음
            return {"paths": [], "error": "파일 선택 창이 이미 열려 있습니다 — 먼저 그 창을 닫거나 경로를 직접 붙여 넣으세요"}
        t = _pick_thread = threading.Thread(target=run, daemon=True)
        t.start()
        t.join(timeout=PICK_TIMEOUT)
        if t.is_alive():
            return {"paths": [], "error": "파일 선택 창 응답이 없습니다(시간 초과) — 경로를 직접 붙여 넣으세요"}
    finally:
        _pick_lock.release()
    return out


def _job_start(target):
    jid = uuid.uuid4().hex[:8]
    job = {"q": queue.Queue(), "history": [], "done": False, "result": None}
    _jobs[jid] = job
    def log(msg):
        ev = {"type": "log", "msg": str(msg)}
        job["history"].append(ev); job["q"].put(ev)
    def run():
        try:
            res = target(log)
            ev = {"type": "done", "result": res}
        except Exception as e:      # noqa: BLE001
            ev = {"type": "error", "msg": str(e)}
        job["history"].append(ev); job["q"].put(ev); job["done"] = True; job["result"] = ev
    threading.Thread(target=run, daemon=True).start()
    return jid


def _plan_json(cfg):
    client = get_client()
    disc = client.list_disclosures(cfg.corp_code, f"{cfg.start_year}0101", cfg.as_of.strftime("%Y%m%d"))
    plan = SEL.plan_reports(disc, cfg.start_year, cfg.end_year, cfg.as_of, cfg.include_half, set(cfg.fs_divs))
    return {"periods": [p.key for p in plan.periods],
            "reports": [{"rcept_no": r.rcept_no, "report_nm": r.report_nm, "rcept_dt": r.rcept_dt.isoformat(), "kind": r.kind,
                         "period_end": r.period_end.isoformat(), "corrected": r.corrected, "fs_divs": list(r.fs_divs)}
                        for r in sorted(plan.reports.values(), key=lambda r: r.period_end)],
            "base_report": plan.base_report.rcept_no if plan.base_report else None, "log": plan.log}


def _toc_json(cfg):
    mapping = M.load()
    base, toc = P.toc_preview(cfg, get_client(), mapping, lambda m: None)
    return {"base": {"rcept_no": base.rcept_no, "report_nm": base.report_nm},
            "entries": [{"no": e.no, "title": e.title, "title_norm": e.title_norm, "topic": e.topic, "recommended": e.recommended,
                         "table_count": e.table_count, "sheet": sheet_for(e.topic, mapping) if e.topic else "rN_기타"} for e in toc]}


def _presets():
    os.makedirs(PRESET_DIR, exist_ok=True)
    out = {}
    for f in os.listdir(PRESET_DIR):
        if f.endswith(".json"):
            with open(os.path.join(PRESET_DIR, f), encoding="utf-8") as fh:
                d = json.load(fh)
            out[d.get("name", f[:-5])] = d.get("note_include", d.get("include"))
    return out


def _norm(p):
    return os.path.normcase(os.path.realpath(p))


def _within(path, root):
    path, root = _norm(path), _norm(root)
    return path == root or path.startswith(root.rstrip("/\\") + os.sep)


def _open_allowed(d):
    if _within(d, CFG.OUTPUT_DIR):
        return True
    for job in list(_jobs.values()):
        ev = job.get("result")
        res = ev.get("result") if isinstance(ev, dict) and ev.get("type") == "done" else None
        if isinstance(res, dict) and res.get("path") and _norm(os.path.dirname(res["path"])) == _norm(d):
            return True
    return False


class Handler(BaseHTTPRequestHandler):
    def _guard(self, post):
        port = self.server.server_address[1]
        hosts = {f"127.0.0.1:{port}", f"localhost:{port}"}
        origin = self.headers.get("Origin")
        ok = self.headers.get("Host") in hosts and (origin is None or origin in {"http://" + h for h in hosts})
        if ok and post and not (self.headers.get("Content-Type") or "").lower().startswith("application/json"):
            ok = False
        if not ok:
            self._send(403, {"error": "허용되지 않은 출처"})
        return ok

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(code); self.send_header("Content-Type", ctype); self.send_header("Content-Length", str(len(data)))
        self.end_headers(); self.wfile.write(data)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query)
        if u.path.startswith("/api/") and not self._guard(False):
            return
        try:
            if u.path == "/":
                with open(os.path.join(HERE, "web", "index.html"), "rb") as f:
                    return self._send(200, f.read(), "text/html; charset=utf-8")
            if u.path == "/api/config":
                c = CFG.load(); return self._send(200, {"has_key": bool(c["api_key"]), "python": c["python"], "server_started": SERVER_STARTED, "code_version": CODE_VERSION})
            if u.path == "/api/search":
                return self._send(200, get_client().search(q.get("q", [""])[0]))
            if u.path == "/api/presets":
                return self._send(200, _presets())
            if u.path.startswith("/api/job/"):
                j = _jobs.get(u.path.rsplit("/", 1)[1])
                if not j:
                    return self._send(404, {"error": "없음"})
                return self._send(200, {"done": j["done"], "result": j["result"]})
            if u.path.startswith("/api/events/"):
                return self._sse(u.path.rsplit("/", 1)[1])
            return self._send(404, {"error": "not found"})
        except Exception as e:     # noqa: BLE001
            return self._send(500, {"error": str(e)})

    def _sse(self, jid):
        job = _jobs.get(jid)
        if not job:
            return self._send(404, {"error": "job 없음"})
        self.send_response(200); self.send_header("Content-Type", "text/event-stream; charset=utf-8"); self.send_header("Cache-Control", "no-cache"); self.end_headers()
        sent = 0
        try:
            while True:
                hist = job["history"]
                while sent < len(hist):
                    self.wfile.write(f"data: {json.dumps(hist[sent], ensure_ascii=False, default=str)}\n\n".encode("utf-8")); self.wfile.flush(); sent += 1
                if job["done"] and sent >= len(job["history"]):
                    return
                try:
                    job["q"].get(timeout=1.0)
                except queue.Empty:
                    pass
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            return          # 브라우저가 연결을 닫음

    def do_POST(self):
        u = urlparse(self.path)
        if u.path.startswith("/api/") and not self._guard(True):
            return
        try:
            body = self._body()
            if u.path == "/api/config":
                CFG.save(body.get("api_key", "").strip()); return self._send(200, {"ok": True})
            if u.path == "/api/plan":
                return self._send(200, _plan_json(P.BuildConfig.from_dict(body)))
            if u.path == "/api/toc":
                return self._send(200, _toc_json(P.BuildConfig.from_dict(body)))
            if u.path == "/api/build":
                cfg = dict(body, source="dart")       # 파일 입력은 /api/file/build 로만
                return self._send(200, {"job": _job_start(lambda log: run_build(cfg, log))})
            if u.path == "/api/file/scan":
                return self._send(200, _file_scan_json(body))
            if u.path == "/api/file/build":
                cfg = _check_file_cfg(body.get("cfg") if "cfg" in body else body)
                return self._send(200, {"job": _job_start(lambda log: run_file_build(cfg, log))})
            if u.path == "/api/pick-files":
                return self._send(200, _pick_files())
            if u.path == "/api/refresh":
                wbp = body["workbook"]
                def run(log):
                    conf = CFG.load()
                    env = dict(os.environ, PYTHONIOENCODING="utf-8")
                    p = subprocess.Popen([conf["python"], os.path.join(HERE, "excel", "refresh.py"), "--workbook", wbp], cwd=HERE, env=env,
                                         stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace")
                    tail = []
                    for line in p.stdout:          # 줄 단위로 바로 로그(SSE)로 흘린다
                        line = line.rstrip("\r\n"); log(line); tail = (tail + [line])[-10:]
                    if p.wait() != 0:
                        raise RuntimeError(("업데이트 실패: " + " / ".join(tail))[-800:])
                    return {"path": wbp}
                return self._send(200, {"job": _job_start(run)})
            if u.path == "/api/presets":
                os.makedirs(PRESET_DIR, exist_ok=True)
                name = re.sub(r"[^0-9A-Za-z가-힣_\- ]", "_", str(body["name"])).strip(". ")[:40].strip(". ") or "preset"
                with open(os.path.join(PRESET_DIR, f"{name}.json"), "w", encoding="utf-8") as f:
                    json.dump({"name": name, "note_include": body["note_include"]}, f, ensure_ascii=False, indent=1)
                return self._send(200, {"ok": True})
            if u.path == "/api/open-folder":
                p = body.get("path", "")
                d = p if os.path.isdir(p) else os.path.dirname(p)
                if not (os.path.isdir(d) and _open_allowed(d)):
                    return self._send(403, {"error": "허용되지 않은 폴더"})
                subprocess.Popen(["explorer", d])
                return self._send(200, {"ok": True})
            return self._send(404, {"error": "not found"})
        except BadRequest as e:
            return self._send(400, {"error": str(e)})
        except Exception as e:     # noqa: BLE001
            return self._send(500, {"error": str(e)})


def make_server(port):
    return ThreadingHTTPServer(("127.0.0.1", port), Handler)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--port", type=int, default=8771); ap.add_argument("--no-browser", action="store_true")
    a = ap.parse_args()
    srv = make_server(a.port)
    url = f"http://127.0.0.1:{srv.server_address[1]}/"
    print(f"dart_fs 서버: {url}")
    if not a.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
