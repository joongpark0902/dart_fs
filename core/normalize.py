"""텍스트·금액·단위·제목 정규화. 모든 파서가 이 모듈만 거쳐 문자열을 다룬다."""
import re

BR = "\u2028"                     # vendor xml_fix 가 &cr; 를 바꿔 둔 줄바꿈 표식
_WS = re.compile(r"[\s\u3000\u2028\u00a0]+")
_ROMAN = "IVXLⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ"
_NUM_PREFIX = re.compile(
    rf"^(?:[{_ROMAN}]+\s*\.|\(\s*\d{{1,2}}\s*\)|\d{{1,2}}(?:\.\d+)+\s*|\d{{1,2}}\s*\.|[가-하]\s*\.|[①-⑳])\s*")
# 주석 참조만 제거: (주1)·(주 2,3)·(주석 1)·(*)·(*1)·( *2 ). (주)한국, (주식결제형) 같은 일반 괄호는 보존한다.
# 각주 별표 꼴은 note_parser._FOOTMARK 와 같다 — '기타(*)'(FY2022)와 '기타 (*1)'(FY2023~)이 같은 행으로 묶여야 한다.
_NOTE_REF = re.compile(r"\(\s*주\s*\d+[^)]*\)|\(\s*주석\s*[\d,\s]+\)|\(\s*\*+\s*\d*\s*\)")
_UNIT_WORDS = {"원": 1, "천원": 10**3, "만원": 10**4, "십만원": 10**5, "백만원": 10**6, "천만원": 10**7,
               "억원": 10**8, "십억원": 10**9, "백억원": 10**10}
_UNIT_RE = re.compile(r"단위\s*[:：]?\s*([^)\]]*?)\s*[)\]]")
_HEADING = re.compile(r"^(\d{1,2})\s*\.\s*(?!\d)(\S.*)$")
_BODY_MARKERS = [r"\(\d+\)", r"회사는", r"회사가", r"연결실체는", r"연결회사는", r"당사는", r"당기말", r"당기 및", r"당기와",
                 r"당기 중", r"다음은", r"(?<![변운이활])동 (?=재무제표|회사|사항|금액|거래|계약|기준|주석)", r"보고기간", r"재무제표 작성", r"회사의 ", r"\s[가-힣]+(?:습니다|입니다)"]
_FISCAL = re.compile(r"제\s*(\d{1,3})\s*\(?\s*(당|전전|전)?\s*\)?\s*기")


def clean_text(s):
    """공백류(전각·U+2028·NBSP 포함)를 한 칸으로 줄이고 양끝을 자른다."""
    return _WS.sub(" ", s or "").strip()


def compact(s):
    """공백을 전부 뺀 문자열. 제목 비교용."""
    return _WS.sub("", s or "")


def norm_label(s):
    """계정명·행 라벨 비교 키. 번호 접두어·주석 참조·공백·구분 기호를 뺀다."""
    t = clean_text(s)
    for _ in range(2):
        t = _NUM_PREFIX.sub("", t)
    t = _NOTE_REF.sub("", t)
    t = t.replace("ㆍ", "·").replace("・", "·")
    t = re.sub(r"[\s·,\-/_:：]", "", t)
    return t


def parse_amount(s):
    """'(1,234)'·'△1,234'·'-1,234' 는 음수, '-'·빈칸·문자는 None. 소수점은 float."""
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return s
    t = clean_text(s).replace(",", "").replace(" ", "")
    if t in ("", "-", "－", "—", "–"):
        return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1]
    if t[:1] in "△▲Δ-－":
        neg, t = True, t[1:]
    if not re.fullmatch(r"\d+(?:\.\d+)?", t):
        return None
    v = float(t) if "." in t else int(t)
    return -v if neg else v


class Percent(float):
    """'83.83%' 를 숫자로 읽은 값(0.8383). writer 가 0.0% 서식을 입히는 표식(값은 그냥 float)."""
    __slots__ = ()


def parse_percent(s):
    """'83.83%'·'(1.5%)'·'-2 %' → Percent(0.8383 등). '%' 로 끝나는 숫자가 아니면 None."""
    if not isinstance(s, str):
        return None
    t = clean_text(s).replace(",", "").replace(" ", "")
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg, t = True, t[1:-1]
    if t[:1] in "△▲Δ-－":
        neg, t = True, t[1:]
    if not t.endswith("%") or not re.fullmatch(r"\d+(?:\.\d+)?", t[:-1]):
        return None
    v = round(float(t[:-1]) / 100, 12)
    return Percent(-v if neg else v)


def parse_unit(s):
    """'(단위 : 천원)' → 1000. 단위 표기가 없으면 None."""
    m = _UNIT_RE.search(clean_text(s))
    if not m:
        return None
    word = compact(m.group(1))
    return _UNIT_WORDS.get(word)


def unit_multiplier(word):
    """단위 낱말 '천원'·'백만 원' → 배수. 단위 낱말이 아니면 None."""
    return _UNIT_WORDS.get(compact(word))


_COL_UNIT = re.compile(r"\(\s*(?:단위\s*[:：]?\s*)?([^()]{1,8}?)\s*\)")


def split_col_unit(label):
    """열 라벨 속 단위 표기 '자본금(천원)'·'금액(단위: 백만원)' → ('자본금', 1000). 단위 낱말이 아니면('(주)'·'(*)') (label, None).
    표시용 라벨은 바꾸지 않고 단위 판정·라벨 비교에만 쓴다."""
    if not label:
        return label, None
    for m in _COL_UNIT.finditer(label):
        u = _UNIT_WORDS.get(compact(m.group(1)))
        if u is not None:
            return (clean_text(label[:m.start()] + " " + label[m.end():]) or label), u
    return label, None


def split_heading(s):
    """'9. 재고자산당기말 및 …' → ('9', '재고자산', '당기말 및 …'). 주석 제목이 아니면 None."""
    t = clean_text(s)
    m = _HEADING.match(t)
    if not m:
        return None
    no, rest = m.group(1), m.group(2)
    cut = len(rest)
    for pat in _BODY_MARKERS:
        mm = re.search(pat, rest)
        if mm and 0 < mm.start() < cut:
            cut = mm.start()
    title = rest[:cut].strip(" :,-·")
    if not title:
        return None
    if len(title) > 40:
        title = title[:40]
    return no, title, rest[cut:].strip()


def period_role(s):
    """헤더·캡션 텍스트에서 당기/전기/전전기 역할을 읽는다. 계정명은 None."""
    c = compact(s)
    if not c or len(c) > 24:
        return None
    m = _FISCAL.search(clean_text(s))
    if m and m.group(2):
        return {"당": "cur", "전": "prev", "전전": "prev2"}[m.group(2)]
    if m:  # '제 28 기 (당기)': 회차 접두어를 떼고 나머지로 판정
        c = re.sub(r"^제\d+기", "", c)
    mm = re.fullmatch(r"\(?(전전|당|전)\)?(?:반|분)?기\)?(?:말|초)?\)?", c)
    if mm:
        return {"당": "cur", "전": "prev", "전전": "prev2"}[mm.group(1)]
    return None


def fiscal_no(s):
    """'제 28(당) 기말' → 28."""
    m = _FISCAL.search(clean_text(s))
    return int(m.group(1)) if m else None


def sce_row_norm(label):
    """원문 자본변동표 행의 비교 키: '2025.1.1(당기초)'·'2024.12.31(전기말)' 은 연도와 무관하게 '기초'·'기말'."""
    return "기초" if "기초" in label else ("기말" if "기말" in label else norm_label(label))
