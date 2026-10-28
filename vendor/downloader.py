import os
import re
import zipfile

import requests

from xml_fix import DART_BR_ENTITY

# ── 내부 상수 ────────────────────────────────────────────────────────────────
_DOC_URL = "https://opendart.fss.or.kr/api/document.xml"


# ── 4. 문서 다운로드 ──────────────────────────────────────────────────────────
_ILLEGAL_FS_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]+')
_DOC_NAME_TAG     = re.compile(r"<DOCUMENT-NAME[^>]*>(.*?)</DOCUMENT-NAME>",
                               re.IGNORECASE | re.DOTALL)


# ── 저장 폴더·파일명 규칙 (tkinter 앱과 web/ 서버가 함께 쓴다) ──────────────
def report_year(report_nm):
    """
    report_nm 끝의 결산기 표기에서 사업연도를 뽑는다.
    '사업보고서 (2024.12)' → '2024',  '연결감사보고서 (2025.12)' → '2025'
    """
    if "(" in report_nm:
        tail = report_nm.split("(")[-1][:4]
        if tail.isdigit():
            return tail
    return "unknown"


def report_folder(report_nm, report_types):
    """
    저장 폴더에 쓸 보고서 유형명. 연결감사보고서는 별도 폴더로 분리해
    같은 해 별도·연결 감사보고서가 한 폴더에 섞이지 않게 한다.
    """
    if "연결감사보고서" in report_nm:
        return "연결감사보고서"
    return next((rt for rt in report_types if rt in report_nm), "기타")


_QUARTER_BY_MONTH = {"03": "1분기", "06": "2분기", "09": "3분기", "12": "4분기"}


def report_basename(report_nm, rtype, year):
    """
    저장 파일명의 앞부분. 기본은 '보고서유형_연도'.
    분기보고서만 한 해에 두 번(1·3분기) 나오므로 분기를 덧붙여 구분한다.
      '분기보고서 (2025.03)' → '분기보고서_2025_1분기'
    """
    base = f"{rtype}_{year}"
    if rtype != "분기보고서":
        return base
    m = re.search(r"\((\d{4})\.(\d{2})", report_nm)
    if not m:
        return base
    month = m.group(2)
    return f"{base}_{_QUARTER_BY_MONTH.get(month, month + '월')}"


def safe_filename(name, fallback="문서"):
    """윈도우 파일명으로 쓸 수 없는 문자를 걷어내고 길이를 자른다."""
    cleaned = _ILLEGAL_FS_CHARS.sub(" ", name or "")
    cleaned = " ".join(cleaned.split()).strip(" .")
    return cleaned[:80] or fallback


def _read_document_name(path):
    """
    공시 원문 첫머리의 <DOCUMENT-NAME>을 읽어 문서 제목을 돌려준다.
    (첨부문서 파일명에 쓴다. 전체를 파싱하지 않고 앞부분만 훑는다)
    """
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            head = f.read(65536)
    except OSError:
        return ""
    m = _DOC_NAME_TAG.search(head)
    if not m:
        return ""
    # 제목 안의 태그·엔티티 제거
    text = re.sub(r"<[^>]*>", "", m.group(1))
    text = text.replace(DART_BR_ENTITY, " ")
    return " ".join(text.split())


def _rename_extracted(save_dir, rcept_no, extracted, base_name, log):
    """
    접수번호 파일명({rcept_no}.xml)을 알아보기 쉬운 이름으로 바꾼다.
      본문     → {base_name}.xml            예: 사업보고서_2024.xml
      첨부문서 → {base_name}_{문서제목}.xml  예: 사업보고서_2024_감사보고서.xml
    반환값: 최종 파일명 리스트
    """
    # 본문 파일은 '{rcept_no}.확장자', 첨부는 '{rcept_no}_00760.xml' 형태다.
    # 접수번호로 딱 떨어지는 게 없으면 이름이 가장 짧은 것을 본문으로 본다.
    def is_main(f):
        return os.path.splitext(f)[0] == rcept_no

    main_first = sorted(extracted, key=lambda f: (not is_main(f), len(f), f))
    final, used = [], set()

    for idx, fname in enumerate(main_first):
        src = os.path.join(save_dir, fname)
        if not os.path.exists(src):
            continue
        ext = os.path.splitext(fname)[1] or ".xml"

        if idx == 0:
            stem = base_name
        else:
            doc_name = safe_filename(_read_document_name(src), fallback=f"첨부{idx}")
            stem = f"{base_name}_{doc_name}"

        candidate = f"{stem}{ext}"
        seq = 2
        while candidate.lower() in used or (
            os.path.exists(os.path.join(save_dir, candidate)) and candidate != fname
        ):
            candidate = f"{stem}_{seq}{ext}"
            seq += 1
        used.add(candidate.lower())

        if candidate != fname:
            try:
                os.replace(src, os.path.join(save_dir, candidate))
            except OSError as e:
                log(f"파일명 변경 실패 ({fname}): {e}")
                candidate = fname
        final.append(candidate)

    return final


def download_document(api_key, rcept_no, save_dir, log_fn=None, base_name=None):
    """
    rcept_no에 해당하는 공시 문서를 save_dir에 다운로드·압축해제한다.
    base_name을 주면 '사업보고서_2024.xml'처럼 알아보기 쉬운 이름으로 저장한다.

    이미 .done_{rcept_no} 마커가 있으면 건너뛴다. 마커를 접수번호별로 두므로
    같은 폴더에 여러 공시가 들어가도(1·3분기 보고서, 별도·연결 감사보고서 등)
    서로를 건너뛰지 않는다. 마커 파일에는 그 공시가 만든 파일 목록이 들어간다.
    반환값: {"status": "성공"|"건너뜀"|"실패", "files": [...]}
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    done_marker   = os.path.join(save_dir, f".done_{rcept_no}")
    legacy_marker = os.path.join(save_dir, ".done")   # 구버전(폴더당 1개) 마커

    def existing_files():
        """마커에 적힌 파일 중 실제로 남아 있는 것들."""
        try:
            with open(done_marker, encoding="utf-8") as f:
                listed = [ln.strip() for ln in f if ln.strip()]
        except OSError:
            listed = []
        found = [f for f in listed if os.path.exists(os.path.join(save_dir, f))]
        if found:
            return found
        # 마커가 비었거나(구버전) 이름이 바뀐 경우 접수번호로 되짚는다
        if not os.path.isdir(save_dir):
            return []
        return [f for f in os.listdir(save_dir)
                if f.startswith(rcept_no) and not f.startswith(".done")]

    def write_marker(files):
        with open(done_marker, "w", encoding="utf-8") as f:
            f.write("\n".join(files))

    if os.path.exists(done_marker):
        log(f"[건너뜀] {rcept_no} (이미 존재)")
        return {"status": "건너뜀", "files": existing_files()}

    # 구버전으로 받아둔 폴더: 이 접수번호 파일이 실제로 있으면 완료로 인정하고 마커만 갱신
    if os.path.exists(legacy_marker):
        old = [f for f in os.listdir(save_dir)
               if f.startswith(rcept_no) and not f.startswith(".done")]
        if old:
            write_marker(old)
            log(f"[건너뜀] {rcept_no} (이미 존재)")
            return {"status": "건너뜀", "files": old}

    os.makedirs(save_dir, exist_ok=True)

    resp = requests.get(_DOC_URL, params={"crtfc_key": api_key, "rcept_no": rcept_no})
    if not resp.content.startswith(b"PK"):
        msg = resp.text[:200]
        log(f"[실패] {rcept_no}: {msg}")
        return {"status": "실패", "files": [], "error": msg}

    zip_path = os.path.join(save_dir, f"{rcept_no}.zip")
    with open(zip_path, "wb") as f:
        f.write(resp.content)

    with zipfile.ZipFile(zip_path) as z:
        extracted = z.namelist()
        z.extractall(save_dir)

    os.remove(zip_path)

    if base_name:
        extracted = _rename_extracted(save_dir, rcept_no, extracted,
                                      safe_filename(base_name), log)

    write_marker(extracted)
    log(f"[성공] {rcept_no} → {len(extracted)}개 파일")
    return {"status": "성공", "files": extracted}
