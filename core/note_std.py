"""주석 제목 → 표준 주제. 표 형태 분류는 note_parser.table_records 가 돌려주는 kind 를 쓴다."""

def _hit(kw, title_norm):
    if not kw:
        return False
    if kw.startswith("="):
        return len(kw) > 1 and title_norm == kw[1:]
    return kw in title_norm


def detect_topic(title_norm, mapping):
    for t in mapping.topics:
        if any(_hit(kw, title_norm) for kw in t.keywords):
            return t.topic, t.recommended
    return None, False


def assign_topics(entries, mapping):
    for e in entries:
        e.topic, e.recommended = detect_topic(e.title_norm, mapping)
        e.include = True


def sheet_for(topic, mapping):
    return next((t.sheet for t in mapping.topics if t.topic == topic), None)
