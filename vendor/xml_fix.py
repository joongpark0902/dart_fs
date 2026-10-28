import re

_VALID_ENTITY = re.compile(r"&(?:\#\d+|\#x[\da-fA-F]+|amp|lt|gt|quot|apos);")
_TAG_START    = re.compile(r"</?[A-Za-z_!][\w:.-]*")
_ANY_TAG      = re.compile(r"<(/?)([A-Za-z_][\w:.-]*)[^<>]*?(/?)>")

# DART가 줄바꿈용으로 쓰는 비표준 엔티티. 표준 XML 파서는 모르므로
# 보정 단계에서 줄바꿈 표식(U+2028)으로 바꿔두고, HTML 변환 때 <br>로 편다.
_DART_BR_ENTITY = "&cr;"
_BR_SENTINEL    = " "


# ── 10. XML 보정 ──────────────────────────────────────────────────────────────
def _scan_real_tags(xml_text):
    """
    문서 전체를 훑어 '진짜 태그로 쓰인 이름'을 모은다.

    DART 원문 본문에는 <HL036>, <WORLDWIDE> 처럼 꺾쇠로 감싼 제품명·약어가
    그대로 들어 있다. 생김새가 태그와 똑같아 파서가 열린 태그로 오인하고
    'mismatched tag'로 죽는다. 닫는 태그(</X>)나 자기완결 태그(<X/>)로
    한 번이라도 등장한 이름만 실제 태그로 인정하면 이 둘을 갈라낼 수 있다.
    반환값: (닫는 태그로 등장한 이름 set, 자기완결로만 등장한 이름 set)
    """
    closing, selfclosing = set(), set()
    for slash, name, sc in _ANY_TAG.findall(xml_text):
        if slash:
            closing.add(name.upper())
        elif sc:
            selfclosing.add(name.upper())
    return closing, selfclosing - closing


def fix_xml(xml_text):
    """
    DART XML에서 발생하는 오염을 보정한다:
      - 텍스트/속성값 내 날것 & → &amp;
      - 텍스트 내 비태그 < → &lt;
      - 본문에 쓰인 <HL036> 같은 꺾쇠 표기 → 텍스트로 이스케이프
      - 비표준 엔티티 &cr; → 줄바꿈 표식
    유효한 entity 및 정상 태그는 그대로 유지한다.
    """
    closing_tags, self_closing_tags = _scan_real_tags(xml_text)

    result = []
    i = 0
    n = len(xml_text)
    br_len = len(_DART_BR_ENTITY)

    while i < n:
        c = xml_text[i]

        if c == "<":
            if xml_text[i:i+9] == "<![CDATA[":
                end = xml_text.find("]]>", i + 9)
                if end != -1:
                    result.append(xml_text[i:end + 3])
                    i = end + 3
                    continue
            if xml_text[i:i+4] == "<!--":
                end = xml_text.find("-->", i + 4)
                if end != -1:
                    result.append(xml_text[i:end + 3])
                    i = end + 3
                    continue
            if xml_text[i:i+2] == "<?":
                end = xml_text.find("?>", i + 2)
                if end != -1:
                    result.append(xml_text[i:end + 2])
                    i = end + 2
                    continue
            m = _TAG_START.match(xml_text, i)
            if m:
                name = m.group().lstrip("</").upper()
                if name.startswith("!") or name in closing_tags:
                    i, tag_str = _parse_tag(xml_text, i, n)
                    result.append(tag_str)
                    continue
                if name in self_closing_tags:
                    # <BR> 처럼 다른 곳에선 <BR/>로 닫히는 태그 → 자기완결로 맞춰준다
                    i, tag_str = _parse_tag(xml_text, i, n)
                    if not tag_str.endswith("/>"):
                        tag_str = tag_str[:-1] + "/>"
                    result.append(tag_str)
                    continue
                # 닫히는 법이 없는 이름 = 본문에 쓰인 꺾쇠 표기. 텍스트로 넘긴다.
            result.append("&lt;")
            i += 1

        elif c == "&":
            if xml_text[i:i + br_len].lower() == _DART_BR_ENTITY:
                result.append(_BR_SENTINEL)
                i += br_len
                continue
            m = _VALID_ENTITY.match(xml_text, i)
            if m:
                result.append(m.group())
                i = m.end()
            else:
                result.append("&amp;")
                i += 1

        else:
            result.append(c)
            i += 1

    return "".join(result)


def _parse_tag(s, i, n):
    """태그 내부를 순회하며 속성값 안의 & 만 보정. (end_pos, fixed_str) 반환"""
    result = []
    in_quote = None
    while i < n:
        c = s[i]
        if in_quote:
            if c == in_quote:
                result.append(c)
                in_quote = None
                i += 1
            elif c == "&":
                m = _VALID_ENTITY.match(s, i)
                if m:
                    result.append(m.group())
                    i = m.end()
                else:
                    result.append("&amp;")
                    i += 1
            else:
                result.append(c)
                i += 1
        else:
            if c in ('"', "'"):
                in_quote = c
                result.append(c)
                i += 1
            elif c == ">":
                result.append(">")
                i += 1
                break
            else:
                result.append(c)
                i += 1
    return i, "".join(result)


# HTML 변환 쪽에서도 이 표식을 <br>로 펴야 한다
BR_SENTINEL = _BR_SENTINEL

# 다운로드 쪽에서도 <DOCUMENT-NAME>에서 이 엔티티를 걷어낸 뒤 XML을 손본다
DART_BR_ENTITY = _DART_BR_ENTITY
