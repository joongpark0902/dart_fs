"""레코드 → 연도 결합 격자(Block). 행 결합 기준은 정규화 라벨 정확일치만(설계 §3.5)."""
import re
from collections import OrderedDict
from core import normalize as N
from core.model import Row, Block


def _keyed(seq):
    """같은 norm 이 한 기간에 여러 번 나오면 (norm, n) 으로 구분. 항목은 (label, norm, value[, account_id])."""
    seen, out = {}, []
    for item in seq:
        label, norm, value = item[:3]
        k = seen.get(norm, 0)
        seen[norm] = k + 1
        out.append(((norm, k), label, value, item[3] if len(item) > 3 else None))
    return out


def join_rows(groups, period_order):
    """최신 기간 순서를 기준으로 행을 합친다. 과거에만 있는 행은 같은 기간의 직전 행 바로 뒤(없으면 맨 앞)에 넣는다.

    groups: {period_key: [(label, norm, value) 또는 (label, norm, value, account_id), ...]}.
    4번째 원소(account_id)가 있으면 같은 (norm, k번째) 키로 합쳐진 Row.account_id 에 최신 기간 우선으로 채운다.
    """
    present = [p for p in period_order if groups.get(p)]
    if not present:
        return []
    base = present[-1]
    order = OrderedDict()            # key -> Row
    for key, label, value, aid in _keyed(groups[base]):
        order[key] = Row(label, key[0], {base: value}, account_id=aid)
    for p in present[:-1]:
        prev_key = None
        for key, label, value, aid in _keyed(groups[p]):
            if key in order:
                order[key].values[p] = value
                if not order[key].account_id:
                    order[key].account_id = aid
            else:
                items = list(order.items())
                pos = 0 if prev_key is None else next(i for i, (k, _) in enumerate(items) if k == prev_key) + 1
                items.insert(pos, (key, Row(label, key[0], {p: value}, account_id=aid)))
                order = OrderedDict(items)
            prev_key = key
    return list(order.values())


def _norm_of(rec):
    return rec.label_norm if hasattr(rec, "label_norm") else rec.row_norm


def _block(records, periods, title, key_label, key_value):
    groups, units = {}, {}
    for r in records:
        groups.setdefault(r.period_key, []).append((key_label(r), _norm_of(r), key_value(r), getattr(r, "account_id", None)))
        units[r.period_key] = r.unit
    rows = join_rows(groups, periods)
    latest = next((p for p in reversed(periods) if p in units), None)
    return Block(title, units.get(latest, 1), list(periods), rows, meta={"units": units})


def sce_detail_key(r):
    """자본변동표 열(자본 구성요소) 비교 키: 연도별 띄어쓰기 차이('주식발행 초과금'/'주식발행초과금')를 같은 열로."""
    return N.norm_label(r.detail) if r.sj_div == "SCE" and r.detail else None


def fs_blocks(records, periods):
    out = []
    order = {"BS": 0, "IS": 1, "CIS": 2, "CF": 3, "SCE": 4}
    titles = {"BS": "재무상태표", "IS": "손익계산서", "CIS": "포괄손익계산서", "CF": "현금흐름표", "SCE": "자본변동표"}
    keys = sorted({(r.fs_div, r.sj_div, sce_detail_key(r)) for r in records},
                  key=lambda k: (k[0], order.get(k[1], 9), k[2] or ""))
    for fs_div, sj, dkey in keys:
        recs = sorted([r for r in records if r.fs_div == fs_div and r.sj_div == sj and sce_detail_key(r) == dkey],
                      key=lambda r: r.ord)
        b = _block(recs, periods, titles.get(sj, sj), lambda r: r.label, lambda r: r.amount)
        # 열 제목은 최신 기간의 원문 라벨
        latest = max(recs, key=lambda r: periods.index(r.period_key) if r.period_key in periods else -1)
        detail = latest.detail if sj == "SCE" else None
        b.col_label, b.fs_div, b.sj_div = detail, fs_div, sj
        unify_units(b)
        out.append(b)
    return out


_PERIOD_TOK = re.compile(r"\(\s*(?:당|전|전전)\s*(?:반)?기\s*(?:말|초)?\s*\)|전전기|당반기|전반기|당기말|전기말|당기|전기")
_UNIT_CAP = re.compile(r"\(\s*단위[^)]*\)")
_FOOT = re.compile(r"\(\s*주\s*\d+[^)]*\)|\(\*+\)")


_SUBNO = re.compile(r"^\s*(?:\(\s*(\d{1,2})\s*\)|(\d{1,2})\s*\))")


_FISCAL_TOK = re.compile(r"제\s*\d{1,3}\s*\(?\s*(?:당|전전|전)?\s*\)?\s*(?:반)?기\s*(?:말|초)?")
# 기간 표기 바로 앞 원문자(① 제 25(당) 기 / ② 전기 중 …): 당기표·전기표를 가르는 번호라 서명에서 뺀다
_CIRCLED_PERIOD = re.compile(r"[①-⑳]\s*(?=제\s*\d{1,3}\s*\(?\s*(?:당|전전|전)?\s*\)?\s*(?:반)?기|전전기|당반기|전반기|당기|전기)")


def caption_sig(caption):
    """표 캡션에서 기간 표기(당기·전기·제 N(당) 기)·단위·주석 표식을 뺀 서명. 당기표와 전기표, 연도별 표를 같은 표로 묶는 키."""
    t = _SUBNO.sub(lambda m: f"항{m.group(1) or m.group(2)} ", caption or "")
    t = _CIRCLED_PERIOD.sub("", t)
    t = _FISCAL_TOK.sub("", t)
    t = _UNIT_CAP.sub("", t)
    t = _PERIOD_TOK.sub("", t)
    t = _FOOT.sub("", t)
    return N.norm_label(t)


def note_blocks(records, periods):
    out = []
    # 같은 (fs_div, 주석, 기간) 안에서 서명이 같은 서로 다른 표(table_idx)는 등장 순서대로 #2, #3 을 붙인다.
    tables = {}
    for r in records:
        tables.setdefault((r.fs_div, r.note_title and N.norm_label(r.note_title), r.period_key, caption_sig(r.caption)), [])
        lst = tables[(r.fs_div, r.note_title and N.norm_label(r.note_title), r.period_key, caption_sig(r.caption))]
        if r.table_idx not in lst:
            lst.append(r.table_idx)
    groups, base_of = OrderedDict(), {}
    for r in records:
        nt = r.note_title and N.norm_label(r.note_title)
        sig = caption_sig(r.caption)
        n = tables[(r.fs_div, nt, r.period_key, sig)].index(r.table_idx) + 1
        k = (r.fs_div, nt, sig if n == 1 else f"{sig}#{n}", N.norm_label(r.col_label) if r.col_label else None)
        groups.setdefault(k, []).append(r)
        base_of[k] = (r.fs_div, nt, sig)
    counts = {}                      # (fs_div, 주석, 서명) -> {기간: 표 수}
    for (fs_div, nt, p, sig), lst in tables.items():
        counts.setdefault((fs_div, nt, sig), {})[p] = len(lst)
    for k, recs in groups.items():
        fs_div = k[0]
        latest = max(recs, key=lambda r: periods.index(r.period_key) if r.period_key in periods else -1)
        b = _block(recs, periods, latest.caption, lambda r: r.row_label, lambda r: r.value)
        b.col_label, b.fs_div = latest.col_label, fs_div
        b.note_no, b.note_title, b.topic = latest.note_no, latest.note_title, latest.topic
        if len(set(counts[base_of[k]].values())) > 1:
            b.meta.setdefault("warnings", []).append(
                f"주석 {b.note_no} {b.note_title}: 같은 캡션 표 수가 기간마다 다름 — 블록 정렬 확인 필요")
        b.meta["sigs"] = [k[2]]
        out.append((recs, b))
    blocks = [b for _, b in _merge_overlapping(out, periods)]
    for b in blocks:
        unify_units(b)
    return blocks


def _period_set(recs):
    return {r.period_key for r in recs}


def _row_keys(b):
    """겹침 비교용 정규화 행 라벨: 빈 라벨·숫자 없는 행(소제목) 제외."""
    return {r.norm for r in b.rows if r.norm and any(isinstance(v, (int, float)) and not isinstance(v, bool)
                                                     for v in r.values.values())}


def _overlap_ratio(a, b):
    """병합 가능하면 행 겹침 비율(0.5 이상), 아니면 None."""
    (ra, ba), (rb, bb) = a, b
    if ba.topic is None or bb.topic is None or (ba.fs_div, ba.topic) != (bb.fs_div, bb.topic):
        return None
    if (N.norm_label(ba.col_label) if ba.col_label else None) != (N.norm_label(bb.col_label) if bb.col_label else None):
        return None
    if _period_set(ra) & _period_set(rb):
        return None
    ka, kb = _row_keys(ba), _row_keys(bb)
    if not (ka and kb):
        return None
    ratio = len(ka & kb) / min(len(ka), len(kb))
    return ratio if ratio >= 0.5 else None


def _merge_overlapping(items, periods):
    """캡션이 연도별로 달라 갈라진 같은 표(같은 fs_div·topic·열)를 행 겹침으로 합친다. 후보가 여럿이면
    겹침 비율이 가장 높은 앞 블록에 붙인다(동률이면 먼저 나온 블록)."""
    out = []
    for it in items:
        j, best = None, None
        for i, o in enumerate(out):
            ratio = _overlap_ratio(o, it)
            if ratio is not None and (best is None or ratio > best):
                j, best = i, ratio
        if j is None:
            out.append(it)
            continue
        (ra, ba), (rb, bb) = out[j], it
        recs = ra + rb
        latest = max(recs, key=lambda r: periods.index(r.period_key) if r.period_key in periods else -1)
        b = _block(recs, periods, latest.caption, lambda r: r.row_label, lambda r: r.value)
        b.col_label, b.fs_div = latest.col_label, latest.fs_div
        b.note_no, b.note_title, b.topic = latest.note_no, latest.note_title, latest.topic
        fy = lambda rs: "·".join(p for p in periods if p in _period_set(rs))
        b.meta["warnings"] = (ba.meta.get("warnings", []) + bb.meta.get("warnings", [])
                              + [f"주석 {b.note_no} {b.note_title}: 캡션이 달라도 행이 겹쳐 병합 — 기간 {fy(ra)} + {fy(rb)}"])
        b.meta["sigs"] = ba.meta.get("sigs", []) + bb.meta.get("sigs", [])
        out[j] = (recs, b)
    return out


def unify_units(block):
    """기간별 단위가 다르면 최신 기간 단위로 환산한다. 경고 목록 반환."""
    units = block.meta.get("units") or {}
    warns = []
    target = block.unit or 1
    for p, u in list(units.items()):
        if u and u != target:
            for row in block.rows:
                v = row.values.get(p)
                if isinstance(v, (int, float)) and not isinstance(v, N.Percent):   # 비율은 단위 환산 안 함
                    nv = v * u / target
                    row.values[p] = int(nv) if float(nv).is_integer() else nv
            units[p] = target
            warns.append(f"{block.title}: {p} 단위 {u} → {target} 환산")
    if warns:
        block.meta.setdefault("warnings", []).extend(warns)
    return warns


def apply_existing_order(block, existing_labels):
    """기존 시트의 행 라벨 순서를 유지하고 새 행은 끝에 붙인다. 사라진 행은 빈 값 + 플래그."""
    by_norm = {}
    for r in block.rows:
        by_norm.setdefault(r.norm, []).append(r)
    rows, used = [], set()
    for lab in existing_labels:
        n = N.norm_label(lab)
        # 같은 라벨 → 같은 norm → (자본변동표) '2025.1.1(당기초)' 류는 '기초'·'기말' 키. 원문 SCE 행은 norm 이
        # 라벨 정규화와 달라서(fs_parser) 라벨 정규화만으로는 기존 행을 못 찾고 '사라짐'+중복 행이 생겼다.
        cand = ([r for r in block.rows if r.label == lab and id(r) not in used]
                or [r for r in by_norm.get(n, []) if id(r) not in used])
        if not cand and block.sj_div == "SCE":
            cand = [r for r in by_norm.get(N.sce_row_norm(lab), []) if id(r) not in used]
        if cand:
            rows.append(cand[0]); used.add(id(cand[0]))
        else:
            rows.append(Row(lab, n, {}, ["공시에서 사라짐"]))
    for r in block.rows:
        if id(r) not in used:
            rows.append(r)
    return Block(block.title, block.unit, block.periods, rows, block.col_label, block.note_no, block.note_title,
                 block.topic, block.fs_div, block.sj_div, dict(block.meta))
