"""dart_downloader 4개 모듈을 vendor/ 로 스냅샷 복사하고 SOURCES.md 에 기록한다. 원본은 읽기만 한다."""
import hashlib, os, shutil, datetime as dt
# dart_downloader 저장소를 이 폴더 옆에 둔다고 가정(환경변수 DART_DOWNLOADER 로 바꿀 수 있음)
SRC = os.environ.get("DART_DOWNLOADER", os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "dart_downloader"))
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DST = os.path.join(HERE, "vendor")
FILES = ["dart_client.py", "downloader.py", "xml_fix.py", "settings.py"]

def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        h.update(f.read())
    return h.hexdigest()

def main():
    os.makedirs(DST, exist_ok=True)
    lines = ["# vendor 스냅샷 출처", "", f"복사 일시: {dt.datetime.now():%Y-%m-%d %H:%M}", "", "| 파일 | 원본 | sha256 |", "|---|---|---|"]
    for f in FILES:
        src = os.path.join(SRC, f)
        shutil.copyfile(src, os.path.join(DST, f))
        lines.append(f"| {f} | {src} | {sha256(src)} |")
    lines += ["", "수정 금지. 다시 맞추려면 `python scripts/sync_vendor.py`."]
    with open(os.path.join(DST, "SOURCES.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print("vendor 동기화 완료")

if __name__ == "__main__":
    main()
