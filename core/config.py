"""config.txt(인증키)·경로. 다른 앱의 config 는 절대 읽지 않는다."""
import os, re, sys
APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(APP_DIR, "config.txt")
CACHE_DIR = os.path.join(APP_DIR, "cache")
OUTPUT_DIR = os.path.join(APP_DIR, "output")
CORPCODE_PATH = os.path.join(CACHE_DIR, "CORPCODE.xml")
PRESET_DIR = os.path.join(APP_DIR, "mapping", "presets")


def _unquote(v):
    v = v.strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'":
        return v[1:-1]
    return v


def _read_raw():
    out = {}
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, _, v = line.partition("=")
                out[k.strip().lower()] = _unquote(v)
    except OSError:
        pass
    return out


_CTRL = re.compile(r"[\x00-\x1f\x7f]")


def load():
    cfg = {"api_key": "", "python": sys.executable}
    cfg.update(_read_raw())
    # 제어문자가 섞인 경로(예: 셸 echo 가 "\a" 를 BEL 로 바꾼 값)는 복구할 수 없고 엑셀 셀에도 못 쓴다 → 현재 파이썬.
    if _CTRL.search(cfg.get("python") or ""):
        cfg["python"] = sys.executable
    return cfg


def save(api_key, python=None):
    cur = _read_raw()
    cur["api_key"] = api_key
    if python:
        cur["python"] = python
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write("# DART OpenAPI 인증키 — 본인 키. 공유·커밋 금지.\n")
        f.write(f"api_key = {cur.pop('api_key')}\n")
        for k, v in cur.items():
            f.write(f"{k} = {v}\n")
    return CONFIG_PATH
