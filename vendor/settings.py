"""실행 경로와 인증키 저장을 다룬다.

exe로 묶으면 __file__ 은 임시 해제 폴더(_MEIxxxx)를 가리키므로
저장 폴더·CORPCODE 캐시는 실행파일이 놓인 폴더 기준으로 잡는다.
"""
import os
import sys

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

DEFAULT_DOWNLOADS = os.path.join(APP_DIR, "downloads")
CORPCODE_PATH = os.path.join(APP_DIR, "CORPCODE.xml")
CONFIG_PATH = os.path.join(APP_DIR, "config.txt")

CONFIG_HEADER = (
    "# DART OpenAPI 인증키 (https://opendart.fss.or.kr 에서 발급)\n"
    "# 이 파일에는 본인 인증키가 들어 있습니다. 공유하거나 git에 올리지 마세요.\n"
)


def load_api_key():
    """config.txt에서 인증키를 읽는다. 없으면 빈 문자열."""
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("#") or "=" not in line:
                    continue
                name, _, value = line.partition("=")
                if name.strip().lower() == "api_key":
                    return value.strip()
    except OSError:
        pass
    return ""


def save_api_key(key):
    """인증키를 실행파일 옆 config.txt에 저장한다."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write(f"{CONFIG_HEADER}api_key = {key}\n")
    return CONFIG_PATH
