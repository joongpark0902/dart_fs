"""재무제표 수치 → FsRecord. API JSON(상장·사업보고서 제출사)과 원문 XML(감사보고서) 두 경로."""
import re
from core import normalize as N
from core import xml_doc as X
from core.model import FsRecord

STATEMENTS = [("BS", r"재무상태표|대차대조표"), ("CIS", r"포괄손익계산서"), ("IS", r"손익계산서"),
              ("SCE", r"자본변동표"), ("CF", r"현금흐름표")]
_SJ_ORDER = {"BS": 0, "IS": 1, "CIS": 2, "CF": 3, "SCE": 4}
_API_FLOW = {"IS", "CIS", "CF"}
_SJ_NAME = {"BS": "재무상태표", "IS": "손익계산서", "CIS": "포괄손익계산서", "SCE": "자본변동표", "CF": "현금흐름표"}
_YEAR = re.compile(r"(20\d{2})\s*[.년]")
_LABEL_HDR = re.compile(r"과\s*목|구\s*분|계정과목|항\s*목")


def sce_detail_label(detail):
    d = (detail or "").strip()
    if d in ("", "-"):
        return ""
    # 총계 열은 '연결재무제표 [member]'·'별도재무제표 [member]' 한 마디. 2022년 이전 응답은 구성요소에도
    # '[member]' 를 붙이므로('자본 [member]|…|자본금 [member]') [member] 여부로 총계를 판정하면 안 된다.
    if "|" not in d and "재무제표" in d:
        return "총계"
    last = d.split("|")[-1]
    return N.clean_text(last.replace("[구성요소]", "").replace("[member]", ""))


def _amt(v):
    return N.parse_amount(v) if v not in (None, "") else None


def api_records(items, fs_div, periods, rcept_no):
    """fnlttSinglAcntAll 응답 → 레코드. 연간: cur/prev/prev2. 반기: 손익·현금흐름은 누적(add) 열, BS 전기는 prev_bs(연간 기말)."""
    out = []
    if "cur" not in periods:
        return out
    half = periods["cur"].kind == "half"
    for it in sorted(items, key=lambda x: (_SJ_ORDER.get(x["sj_div"], 9), int(x.get("ord") or 0))):
        sj = it["sj_div"]
        acc_id = it.get("account_id") or ""
        acc_id = None if (not acc_id or acc_id.startswith("-")) else acc_id
        detail = sce_detail_label(it.get("account_detail")) if sj == "SCE" else None
        label = N.clean_text(it["account_nm"])
        if not half:
            cols = [("cur", it.get("thstrm_amount")), ("prev", it.get("frmtrm_amount")), ("prev2", it.get("bfefrmtrm_amount"))]
        elif sj in ("IS", "CIS"):      # 손익: 누적(add) 열만. 비면 3개월 값을 반기 누적으로 오인하지 않도록 레코드 생략
            cols = [("cur", it.get("thstrm_add_amount")), ("prev", it.get("frmtrm_add_amount"))]
        elif sj == "CF":               # 현금흐름: 반기 thstrm_amount 가 이미 누적이라 add 가 비면 대체 허용
            cols = [("cur", it.get("thstrm_add_amount") or it.get("thstrm_amount")),
                    ("prev", it.get("frmtrm_add_amount") or it.get("frmtrm_amount"))]
        else:
            cols = [("cur", it.get("thstrm_amount")), ("prev_bs", it.get("frmtrm_amount"))]
        for role, raw in cols:
            p = periods.get(role)
            if p is None or raw in (None, ""):
                continue
            out.append(FsRecord(p.key, fs_div, sj, int(it.get("ord") or 0), label, N.norm_label(label),
                                _amt(raw), 1, detail, acc_id, rcept_no, "API", "prev" if role == "prev_bs" else role))
    return out


def _fs_section(nodes):
    return X.section(nodes,
                     lambda n: n.kind == "title" and "재무제표" in N.compact(n.title) and "주석" not in N.compact(n.title),
                     lambda n: n.kind == "title" and N.compact(n.title).startswith("주석"))


def _title_table(grid):
    """제목표면 (sj_div, unit, {role: fiscal_no}) 아니면 None."""
    if grid.header or len(grid.body) > 6:
        return None
    t = grid.text()
    head = N.compact(t)[:12]
    for sj, pat in STATEMENTS:
        if re.search(pat, head):
            fiscal = {}
            for m in re.finditer(r"제\s*(\d{1,3})\s*\(?\s*(당|전전|전)\s*\)?\s*기", t):
                fiscal[{"당": "cur", "전": "prev", "전전": "prev2"}[m.group(2)]] = int(m.group(1))
            if not fiscal:      # '제 26 기말 2023년 12월 31일 현재' 처럼 (당)(전) 표시가 없으면 첫 회차 = 당기
                m = re.search(r"제\s*(\d{1,3})\s*기", t)
                if m:
                    fiscal["cur"] = int(m.group(1))
            return sj, (N.parse_unit(t) or 1), fiscal
    return None


def _period_groups(header_row, periods, fiscal):
    """헤더 첫 줄에서 열 인덱스 → Period. 같은 텍스트가 이어진 colspan 그룹을 하나로.
    회차만 있는 헤더('제 25 기말')는 당기 회차에서 전기·전전기를 유도하고, '제 N 기초' 는 제 N-1 기말로 본다."""
    fiscal = dict(fiscal)
    if "cur" in fiscal:
        fiscal.setdefault("prev", fiscal["cur"] - 1)
        fiscal.setdefault("prev2", fiscal["cur"] - 2)
    groups, c = [], 0
    while c < len(header_row):
        t = header_row[c]
        e = c
        while e + 1 < len(header_row) and header_row[e + 1] == t:
            e += 1
        role = N.period_role(t)
        fn = N.fiscal_no(t)
        if fn is not None and "기초" in N.compact(t) and "기말" not in N.compact(t):
            fn -= 1
            role = None         # 기초 잔액은 회차-1 기말로만 판정
        if role is None and fn is not None:
            role = next((k for k, v in fiscal.items() if v == fn), None)
        if role and role in periods:
            groups.append((list(range(c, e + 1)), periods[role], role))
        c = e + 1
    return groups


def _is_label_col(header_row, body, c):
    h = header_row[c] if c < len(header_row) else ""
    if _LABEL_HDR.search(h):
        return True
    vals = [r[c] for r in body if c < len(r) and r[c]]
    return bool(vals) and sum(1 for v in vals if N.parse_amount(v) is None) > len(vals) * 0.7


def _parse_body(grid, sj, unit, periods, fiscal, fs_div, rcept_no, ord_start, warnings=None):
    out, ord_ = [], ord_start
    hdr = grid.header[0]
    groups = _period_groups(hdr, periods, fiscal)
    if not groups:
        if warnings is not None:
            msg = f"{_SJ_NAME.get(sj, sj)}: 본문 표에서 기간 열을 찾지 못해 건너뜀"
            if msg not in warnings:
                warnings.append(msg)
        return out, ord_
    note_cols = {c for c, h in enumerate(hdr) if "주석" in N.compact(h)}
    label_cols = [c for c in range(grid.ncols) if c not in note_cols and not any(c in g[0] for g in groups) and _is_label_col(hdr, grid.body, c)]
    if not label_cols:
        label_cols = [0]
    for row in grid.body:
        label = N.clean_text(" ".join(row[c] for c in label_cols if row[c]))
        if not label:
            continue
        ord_ += 1
        for cols, period, role in groups:
            vals = [N.parse_amount(row[c]) for c in cols if c < len(row)]
            vals = [v for v in vals if v is not None]
            out.append(FsRecord(period.key, fs_div, sj, ord_, label, N.norm_label(label),
                                vals[0] if vals else None, unit, None, None, rcept_no, "XML", role))
    return out, ord_


def _parse_sce(grid, unit, periods, fs_div, rcept_no, ord_start, warnings):
    """자본변동표: 행 라벨의 연도(2025.1.1 …)로 기간을 잡고, 열 = 자본 구성요소."""
    out, ord_ = [], ord_start
    hdr = grid.header[-1] if grid.header else []
    body = grid.body
    if not hdr and body and _LABEL_HDR.search(body[0][0]):   # 헤더가 TH 가 아니라 본문 첫 행인 표
        hdr, body = body[0], body[1:]
    col_labels = {}
    for c in range(1, grid.ncols):
        h = N.clean_text(hdr[c]) if c < len(hdr) else ""
        if h and "주석" not in N.compact(h):
            col_labels[c] = "총계" if N.compact(h) == "총계" else h
    by_year = {p.end.year: p for p in periods.values()}
    cur = None
    for row in body:
        label = N.clean_text(row[0])
        if not label:
            continue
        m = _YEAR.search(label)
        if m and int(m.group(1)) in by_year:
            cur = by_year[int(m.group(1))]
        if cur is None:
            continue
        ord_ += 1
        norm = N.sce_row_norm(label)
        role = "cur" if cur is periods.get("cur") else "prev"
        for c, cl in col_labels.items():
            if c >= len(row):
                continue
            out.append(FsRecord(cur.key, fs_div, "SCE", ord_, label, norm, N.parse_amount(row[c]), unit, cl, None, rcept_no, "XML", role))
    if col_labels and not out:
        warnings.append("자본변동표 행에서 연도를 찾지 못해 건너뜀")
    return out, ord_


def _is_body_table(g):
    """재무제표 본문표 모양: 테두리 있는 표(BORDER=1)이고 첫 헤더 칸이 '과 목'·'구 분' 류."""
    return str(g.border) == "1" and bool(g.header) and bool(g.header[0]) and bool(_LABEL_HDR.search(g.header[0][0] or ""))


def xml_statements(root, fs_div, periods, rcept_no):
    nodes = _fs_section(X.stream(root))
    recs, units, fiscal_all, warnings = [], {}, {}, []
    seen = []                # 제목표가 나온 재무제표 (순서 유지)
    cur = None               # (sj, unit, fiscal)
    produced = False         # 현재 재무제표가 본문표에서 이미 레코드를 냈나
    ord_ = 0
    for n in nodes:
        if n.kind != "table":
            continue
        g = X.table_grid(n.elem)
        tt = _title_table(g)
        if tt:
            cur = tt
            units[tt[0]] = tt[1]
            if tt[0] not in seen:
                seen.append(tt[0])
            fiscal_all.update(tt[2])
            ord_ = 0
            produced = False
            continue
        if _is_body_table(g) and (cur is None or produced):
            where = "첫 제목표 앞" if cur is None else f"{_SJ_NAME.get(cur[0], cur[0])} 뒤"
            first = next((N.clean_text(r[0]) for r in g.body if r and r[0]), "")
            warnings.append(f"제목표 없는 본문표 발견({where}, 첫 행 '{first}') — 이전 재무제표로 흡수될 수 있음")
        if cur is None or X.is_caption_table(g) or (not g.header and cur[0] != "SCE"):
            continue
        sj, unit, fiscal = cur
        if sj == "SCE":
            got, ord_ = _parse_sce(g, unit, periods, fs_div, rcept_no, ord_, warnings)
        else:
            got, ord_ = _parse_body(g, sj, unit, periods, fiscal, fs_div, rcept_no, ord_, warnings)
        recs.extend(got)
        produced = produced or bool(got)
    for sj in seen:
        if not any(r.sj_div == sj for r in recs):
            warnings.append(f"{_SJ_NAME.get(sj, sj)}: 본문 표를 해석하지 못함")
    if not recs:
        warnings.append("재무제표 섹션에서 표를 찾지 못함")
    return recs, {"units": units, "fiscal": fiscal_all, "warnings": warnings}
