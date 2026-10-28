"""`검증` 시트 — 재무제표 내부 정합(A)·주석 ↔ 재무제표(B)·주석 합계 검산(C).

비교 규칙: 양변을 원으로 환산해(mFS 셀 ×cfg_unit, raw 셀 ×블록 공시 단위) 천원으로 반올림한 뒤 같은지 본다
`=ROUND(L/1000,0)=ROUND(R/1000,0)`. 차이 열은 원 단위 `=ROUND(L-R,0)`. 허용폭(±1) 없음.
원문에 없는 합계·값은 만들지 않는다: 참조할 셀이 하나라도 없으면 그 칸은 `해당 없음`(기울임)이고 비고에 무엇이 없는지 적는다.

행·열 안정성(§ 표식·_layout)은 쓰지 않는다: 이 시트는 사용자 입력이 없고 [업데이트] 때 통째로 다시 쓰며,
모델 시트는 mFS·raw 를 참조하지 이 시트를 참조할 일이 없기 때문이다. 열 순서는 mFS 와 같은 기간 순서를 따른다
(검증 셀 E열부터 = mFS 기간 열과 같은 순번이라 mFS 요약 COUNTIF 가 열 단위로 맞는다).
"""
import re
from collections import OrderedDict, namedtuple
from core import layout as L, normalize as N
from excel import compose as C
from excel.grid import Sheet

SHEET = "검증"
NA = "해당 없음"
_TOTAL_NORMS = ("합계", "총계", "계")
_END_COLS = (None, "합계", "총계", "계", "장부금액", "장부가액", "순장부금액")   # 행 변동표(기초…기말 행)의 열 라벨 허용값
_END_COL_RANKS = (lambda c: _is_end(c), lambda c: c in ("장부금액", "장부가액", "순장부금액"),
                  lambda c: c in _TOTAL_NORMS)                                # 열 변동표(기초…기말 열)의 기말 열 우선순위
_GROSS_WORDS = ("취득원가", "감가상각누계액", "상각누계액", "손상차손누계액", "정부보조금")
_SCE_TOTAL = ("총계", "합계", "자본총계", "계")
_NET_INCOME = ("당기순이익", "당기순손실", "당기순손익")
_REV_EXCLUDE = ("계약자산", "계약부채", "이월", "부문", "지역", "비유동자산", "주요수익")   # B7 후보에서 빼는 캡션 낱말
# '주요수익': '연결회사의 주요 수익은 … 판매입니다' 요약표(중단영업 수익 포함 합계). 낱말 '주요' 만으로는
# '주요 제품에서 … 수익의 구분'·'주요 지리적 시장과 주요 제품에 따라 구분' 같은 본 매출 구분표까지 빠진다.
_TAX_ROW = re.compile(r"^법인세(비용|수익)")                                # B13 행: 법인세비용·법인세수익(A+B)…
_TAX_EXCLUDE = ("당기법인세", "이연법인세", "총", "차감전")              # '법인세비용차감전순이익' 은 법인세 행이 아니다
_ADJUST_WORDS = ("회계이익", "유효세율")                                   # 법인세비용과 회계이익의 관계 조정표(C-합계 제외)
_BORROW_STDS = ("단기차입금", "유동성장기부채", "장기차입금")
_LEASE_TOTAL_EXCLUDE = ("만기", "할인", "현금흐름")                           # B9 '합계' 후보에서 빼는 표(리스부채 만기분석)

CRes = namedtuple("CRes", "lhs rhs reason")          # C 그룹 결과: lhs/rhs 항 목록 또는 reason(해당 없음 사유)


def _is_begin(norm):
    """변동표 기초 행·열: 정규화 라벨이 기초… 또는 당기초… 로 시작."""
    return (norm or "").startswith(("기초", "당기초"))


def _is_end(norm):
    """변동표 기말 행·열: 정규화 라벨이 기말… 또는 당기말… 로 시작."""
    return (norm or "").startswith(("기말", "당기말"))


class Pick:
    """주석 선택기 결과: rows = 좌변에 쓸 Placed 들(여럿이면 더한다), note = 비고에 덧붙일 글(선택 근거),
    reason = 못 고른 사유(해당 없음 비고). rows 가 비면 거짓."""
    __slots__ = ("rows", "note", "reason")

    def __init__(self, rows=(), note=None, reason=None):
        self.rows, self.note, self.reason = list(rows), note, reason

    def __bool__(self):
        return bool(self.rows)

    def __repr__(self):
        return f"Pick({len(self.rows)} rows, note={self.note!r}, reason={self.reason!r})"


def _as_pick(x):
    """선택기 반환값(Placed | None | Pick) → Pick."""
    if isinstance(x, Pick):
        return x
    return Pick([x]) if x is not None else Pick()


def _first(*sels):
    """선택기를 차례로 시도해 처음 고른 것. 모두 못 고르면 사유(reason)가 있는 첫 결과(없으면 None)."""
    def f(pk):
        fallback = None
        for sel in sels:
            x = sel(pk)
            if x:
                return x
            if fallback is None and isinstance(x, Pick) and x.reason:
                fallback = x
        return fallback
    return f


# ---------------------------------------------------------------- 수식

def _won(terms):
    """[(부호, 참조, 단위)] → (원 단위 식, atomic). 단위는 정수 또는 cfg_unit 이름. 같은 단위끼리 묶어 `(a+b)*1000`."""
    groups = OrderedDict()
    for sg, ref, u in terms:
        groups.setdefault(u, []).append((sg, ref))
    parts = []
    for u, items in groups.items():
        inner = "".join((sg if i or sg == "-" else "") + ref for i, (sg, ref) in enumerate(items))
        if u == 1:
            parts.append((inner, len(items) == 1))
        elif len(items) == 1:
            parts.append((f"{inner}*{u}", True))
        else:
            parts.append((f"({inner})*{u}", True))
    if len(parts) == 1:
        return parts[0]
    expr = "".join(("" if i == 0 or p.startswith("-") else "+") + p for i, (p, _) in enumerate(parts))
    return expr, False


def _paren(expr, atomic):
    return expr if atomic else f"({expr})"


def check_formula(lhs, rhs):
    """(검증식, 차이식). 둘 다 `=ROUND(` 로 시작해 finalize.verify 가 검증 셀로 알아본다(차이식은 값이 bool 이 아니라 제외)."""
    l_, la = _won(lhs)
    r_, ra = _won(rhs)
    f = f"=ROUND({_paren(l_, la)}/1000,0)=ROUND({_paren(r_, ra)}/1000,0)"
    d = f"=ROUND({l_}-{_paren(r_, ra)},0)"
    return f, d


def or_formula(lhs1, lhs2, rhs):
    """`=OR(ROUND(L1/1000,0)=ROUND(R/1000,0),ROUND(L2/1000,0)=ROUND(R/1000,0))` — 둘 중 하나만 맞으면 TRUE.
    차이식은 첫 좌변 기준 `=ROUND(L1-R,0)`(머리글 '좌변 − 우변' 과 같은 방향)."""
    f1, d1 = check_formula(lhs1, rhs)
    f2, _ = check_formula(lhs2, rhs)
    return f"=OR({f1[1:]},{f2[1:]})", d1


def _plain(ref):
    return ref.replace("'", "")


def ref_text(terms, ranges=False):
    """좌변·우변 참조 열 문구: `mFS!E6+mFS!E7`. ranges=True 면 같은 열 연속 행(모두 +)은 `rN_x!E10:E12`."""
    if ranges and len(terms) >= 2 and all(sg == "+" for sg, _, _ in terms):
        cells = [_plain(ref).split("!") for _, ref, _ in terms]
        sheets = {s for s, _ in cells}
        cols = {"".join(ch for ch in a if ch.isalpha()) for _, a in cells}
        rows = [int("".join(ch for ch in a if ch.isdigit())) for _, a in cells]
        if len(sheets) == 1 and len(cols) == 1 and rows == list(range(rows[0], rows[0] + len(rows))):
            return f"{cells[0][0]}!{cells[0][1]}:{cells[-1][1]}"
    return "".join((sg if i or sg == "-" else "") + _plain(ref) for i, (sg, ref, _) in enumerate(terms))


# ---------------------------------------------------------------- 주석 블록 선택기

def _cl(b):
    return N.norm_label(b.col_label) if b.col_label else None


def has(b, *words):
    """블록 캡션(제목) 또는 열 라벨에 낱말이 있나(정규화 비교)."""
    text = N.norm_label(b.title or "") + "|" + (N.norm_label(b.col_label) if b.col_label else "")
    return any(w in text for w in words)


def is_cost(b):
    return has(b, "금융원가", "금융비용")


def is_rou(b):
    return has(b, "사용권자산")


def main_ppe(b):
    """유형자산 주석 안의 사용권자산 변동표(기초…기말 행)는 유형자산 전체 변동표가 아니다(B10 몫)."""
    return not is_rou(b)


def main_intangible(b):
    """영업권만 다루는 표(현금창출단위별 영업권 변동 등)는 무형자산 변동표가 아니다. '영업권 및 무형자산' 표는 포함."""
    return not (has(b, "영업권") and not has(b, "무형자산"))


def _blocks(note_placed, fs_div, topic):
    blocks = OrderedDict()
    for p in note_placed:
        if p.block.fs_div == fs_div and p.block.topic == topic:
            blocks.setdefault(id(p.block), []).append(p)
    return list(blocks.values())


def _val(p, pk):
    if pk not in p.block.periods:
        return None
    return p.row.values.get(pk)


def _isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, (bool, N.Percent))


def _one(ps, pred):
    m = [p for p in ps if pred(p.row.norm or "")]
    return m[0] if len(m) == 1 else None


def _sig(b):
    return (b.note_no, L.caption_sig(b.title or ""))


def _norms(ps):
    return {p.row.norm for p in ps if p.row.norm}


def _overlap(a, b):
    return bool(a and b) and len(a & b) / min(len(a), len(b)) >= 0.5


def _end_column_form(blocks, pk, ok):
    """열마다 블록인 변동표(구분 × 기초·취득·…·기말): 같은 표(주석·캡션 서명 같고 행이 겹침)에 `기초` 열 블록이 있을 때만,
    기말 열(기말* > 장부금액류 > 합계류, 합계류는 그 표의 마지막 열일 때만) 블록의 합계 행.
    내역표(취득원가·누계액·장부가액)는 기초 열이 없어 고르지 않는다."""
    begins = [ps for ps in blocks if _is_begin(_cl(ps[0].block))]
    if not begins:
        return None
    for k, rank in enumerate(_END_COL_RANKS):
        for ps in blocks:
            b = ps[0].block
            c = _cl(b)
            if c is None or not rank(c) or not ok(b):
                continue
            if not any(_sig(bp[0].block) == _sig(b) and _overlap(_norms(bp), _norms(ps)) for bp in begins):
                continue
            if k == 2 and ps[0].rownum != max(q[0].rownum for q in blocks if _sig(q[0].block) == _sig(b)):
                continue                      # 합계류 열은 그 표의 마지막 열일 때만(중간 소계 열 제외)
            tot = _one(ps, lambda n: n in _TOTAL_NORMS)
            if tot is not None and _val(tot, pk) is not None:
                return tot
    return None


def _note_row(note_placed, fs_div, topic, pk, *, how, labels=(), where=None, pattern=None, exclude=()):
    """검증에 쓸 주석 행(Placed). 못 찾으면 None. 기간마다 따로 고른다(당기표·전기표 블록이 갈린 경우).
    how="total": 열 라벨 없는 블록 중 정규화 라벨이 정확히 `합계`(없으면 `…합계`) 인 행이 하나뿐이고 그 기간 값이 있는 첫 블록.
    how="end": 변동표만 — (행 변동표) `기초…`(당기초…) 행과 `기말…`(당기말…) 행이 모두 있고 열 라벨이 없거나
               합계·장부금액류인 블록에서, 그 기간 값 있는 기말 행이 정확히 하나면 그 행(여럿이면 '기말 행 여러 개');
               (열 변동표) `_end_column_form`. 합계 행 폴백 없음(가상자산 장부가액 표 오선택 방지).
    how="row": labels 를 우선순위대로, 그 라벨과 정확히 일치하는 행이 하나뿐인 첫 블록(열 라벨 없는 블록 먼저).
    how="sub": 소제목 행(정규화 라벨 = labels, 없으면 그 글자로 시작하는 값 없는 행) 아래 첫 합계 행 → Pick(note '… 소제목 아래 합계').
               소제목 행 자체에 그 기간 값이 있으면 그 행.
    how="total_all": 열 라벨 없는 첫 블록의 `합계` 행 전부(그 기간 값 있는 것) → Pick(여럿이면 note '합계 행 n개 합산').
    how="last": 정규화 라벨이 pattern 에 맞고 exclude 낱말이 없는, 그 기간 값 있는 행 중 마지막 행 → Pick.
    where(block) 로 블록을 더 거른다. 반환은 Placed | None, 새 방식(sub·total_all·last)과 end 의 '해당 없음' 사유는 Pick."""
    blocks = _blocks(note_placed, fs_div, topic)

    def ok(b):
        return where is None or where(b)

    def col_none_first():
        return sorted(blocks, key=lambda ps: 0 if _cl(ps[0].block) is None else 1)

    if how == "total":
        for ps in blocks:
            b = ps[0].block
            if _cl(b) is not None or not ok(b):
                continue
            tot = _one(ps, lambda n: n == "합계") or _one(ps, lambda n: n.endswith("합계"))
            if tot is not None and _val(tot, pk) is not None:
                return tot
        return None
    if how == "row":
        order = sorted(blocks, key=lambda ps: 0 if _cl(ps[0].block) is None else 1)
        for lab in labels:
            for ps in order:
                if not ok(ps[0].block):
                    continue
                p = _one(ps, lambda n, lab=lab: n == lab)
                if p is not None and _val(p, pk) is not None:
                    return p
        return None
    if how == "end":
        # 행 변동표 후보를 등장 순서대로 모은다. 열 라벨 없는(bare) 후보가 하나 이하면 예전처럼 첫 후보.
        # bare 후보가 둘 이상이면 합계·장부금액류 열 라벨 후보 → 열 변동표 → 없으면 '해당 없음'(어느 표가 전체인지 알 수 없음).
        cands, multi = [], []
        for ps in blocks:
            b = ps[0].block
            if _cl(b) not in _END_COLS or not ok(b):
                continue
            if not any(_is_begin(p.row.norm) for p in ps):
                continue
            ends = [p for p in ps if _is_end(p.row.norm) and _val(p, pk) is not None]
            if len(ends) == 1:
                cands.append(ends[0])
            elif ends:
                multi.append(b)                 # 그 기간 값 있는 기말 행이 여럿 — 어느 것인지 알 수 없음
        bare = [p for p in cands if _cl(p.block) is None]
        if len(bare) <= 1 and cands:
            return cands[0]
        labeled = [p for p in cands if _cl(p.block) is not None]
        if labeled:
            return labeled[0]
        col = _end_column_form(blocks, pk, ok)
        if col is not None:
            return col
        if len(bare) >= 2:
            return Pick(reason="변동표 여러 개: " + ", ".join(_short(_caption_summary(p.block)) for p in bare) + " — 합계 표 없음")
        if multi:
            return Pick(reason="기말 행 여러 개: " + ", ".join(_short(_caption_summary(b)) for b in multi))
        return None
    if how == "sub":
        for lab in labels:
            for ps in col_none_first():
                if not ok(ps[0].block):
                    continue
                head = next((i for i, p in enumerate(ps) if p.row.norm == lab), None)
                if head is None:        # 'OO(*)' 는 정규화로 같아지지만, 그 밖의 '…로 시작' 은 값 없는 소제목 행만
                    head = next((i for i, p in enumerate(ps) if (p.row.norm or "").startswith(lab)
                                 and not any(_isnum(v) for v in p.row.values.values())), None)
                if head is None:
                    continue
                if _val(ps[head], pk) is not None:
                    return Pick([ps[head]])
                tot = next((p for p in ps[head + 1:] if p.row.norm in _TOTAL_NORMS), None)
                if tot is not None and _val(tot, pk) is not None:
                    return Pick([tot], f"{lab} 소제목 아래 합계")
        return None
    if how == "total_all":
        for ps in blocks:
            b = ps[0].block
            if _cl(b) is not None or not ok(b):
                continue
            idx = ([i for i, p in enumerate(ps) if p.row.norm == "합계"]
                   or [i for i, p in enumerate(ps) if (p.row.norm or "").endswith("합계")])
            if not idx:
                continue
            # 구조 판단: 마지막 합계 행과 그 앞 합계 행 사이에 숫자 행이 없으면 마지막 행이 총합계(그룹 합계 + 총합계 → 이중 합산 방지)
            if len(idx) >= 2 and not any(any(_isnum(v) for v in p.row.values.values()) for p in ps[idx[-2] + 1:idx[-1]]):
                idx = idx[-1:]
            tots = [ps[i] for i in idx if _val(ps[i], pk) is not None]
            if tots:
                return Pick(tots, f"합계 행 {len(tots)}개 합산" if len(tots) > 1 else None)
        return None
    if how == "last":
        for ps in col_none_first():
            if not ok(ps[0].block):
                continue
            m = [p for p in ps if pattern.match(p.row.norm or "") and not any(x in p.row.norm for x in exclude)
                 and _val(p, pk) is not None]
            if m:
                return Pick([m[-1]])
        return None
    raise ValueError(how)


def _revenue_pick(note_placed, fs_div, pk):
    """B7: topic 매출구분 중 캡션에 계약자산·계약부채·이월·부문·지역·비유동자산 이 없는 블록(열 라벨 없음)을 배치 순서로 보며,
    기간마다 첫 `합계` 행(없으면 첫 `소계` 행, note '소계')에 그 기간 값이 있는 첫 블록. 합계 행이 여러 개여도 첫 행.
    첫 블록이 일부 연도만 다루면(연도별 양식 차이) 다음 블록으로 — 첫 후보가 아닌 블록을 골랐으면 note 에 블록 캡션."""
    k = 0
    for ps in _blocks(note_placed, fs_div, "매출구분"):
        b = ps[0].block
        if _cl(b) is not None or has(b, *_REV_EXCLUDE):
            continue
        k += 1
        for norm, note in (("합계", None), ("소계", "소계")):
            row = (next((p for p in ps if p.row.norm == norm), None)
                   or next((p for p in ps if (p.row.norm or "").endswith(norm)), None))
            if row is not None and _val(row, pk) is not None:
                if k > 1:
                    note = f"블록 '{_short(_caption_summary(b), 40)}'" + (f" {note}" if note else "")
                return Pick([row], note)
    return None


def _capital_pick(note_placed, fs_div, pk):
    """B14: 자본금 행(자본금·보통주자본금) → 열 라벨(단위 표기 뗀 비교)에 자본금이 든 블록의 합계 행(하나·값 있음).
    합계 행이 없으면 해당 없음(비고 '합계 행 없음') — 연도별 표가 합쳐져 보통주 행이 중복될 수 있어 행 합산은 하지 않는다."""
    p = _note_row(note_placed, fs_div, "자본", pk, how="row", labels=("자본금", "보통주자본금"))
    if p is not None:
        return Pick([p])
    reason = None
    for ps in _blocks(note_placed, fs_div, "자본"):
        b = ps[0].block
        if not (b.col_label and "자본금" in N.norm_label(N.split_col_unit(b.col_label)[0])):
            continue
        tots = [p for p in ps if p.row.norm in _TOTAL_NORMS]
        if len(tots) == 1 and _val(tots[0], pk) is not None:
            return Pick(tots)
        if not tots and reason is None:
            reason = "합계 행 없음"
    return Pick(reason=reason) if reason else None


# ---------------------------------------------------------------- C 그룹(블록 자체 검산)

def _term(p, pk):
    return ("+", C._ref(p.sheet, p.rownum, C.RAW_FIRST_VALUE_COL + p.block.periods.index(pk)), p.block.unit or 1)


def c_total(ps, pk):
    """합계 검산: 정규화 라벨이 합계·총계·계 인 행이 정확히 하나면 `SUM(위 행들) = 합계`. 위 행에 소계 행이 있으면
    소계 행들 + 마지막 소계 뒤의 비소계 행들(영업부문 표의 '소계 + 연결조정 = 합계'), 숫자 행(비율 제외)만 개별 참조로 더한다.
    합계 행 없음/여럿 → None(검산 없음), 값 없음 → CRes(reason)."""
    tot = [i for i, p in enumerate(ps) if p.row.norm in _TOTAL_NORMS]
    if len(tot) != 1:
        return None
    if any(w in N.compact(ps[0].block.title or "") for w in _ADJUST_WORDS):   # 법인세 조정표: 첫 행 회계이익은 합산 대상이 아님
        return CRes(None, None, "조정표")
    t = tot[0]
    if _val(ps[t], pk) is None:
        return CRes(None, None, "합계 값 없음")
    above = ps[:t]
    subs = [i for i, p in enumerate(above) if (p.row.norm or "").endswith("소계")]
    chosen = ([above[i] for i in subs] + above[subs[-1] + 1:]) if subs else above
    lhs = [_term(p, pk) for p in chosen if _isnum(_val(p, pk))]
    if not lhs:
        return CRes(None, None, "더할 숫자 행 없음")
    return CRes(lhs, [_term(ps[t], pk)], None)


def c_roll(ps, pk):
    """변동 검산: `기초…`(당기초…) 행과 그 뒤 `기말…`(당기말…) 행이 있으면(구조), 기간마다 그 기간 값 있는 기초·기말 행이
    하나씩일 때 `기초 + SUM(사이 숫자 행) = 기말`. 연도 병합 블록(FY2024 '기말' + FY2025 '당기말')도 기간마다 고른다.
    기초·기말 사이에 취득원가·누계액·정부보조금 행이 있으면 '총액·누계액 구분 표', 값 있는 기초/기말 행이 여럿이면
    '기초 행 여러 개'/'기말 행 여러 개'(검산 항목에서 조용히 빼지 않는다). 사이에 소계 행이 있으면 c_total 과 같이
    소계 행들 + 마지막 소계 뒤 행들만 더한다(퇴직급여 변동표: 당기손익 소계·재측정요소 소계·기여금·지급액)."""
    bi = [i for i, p in enumerate(ps) if _is_begin(p.row.norm)]
    ei = [i for i, p in enumerate(ps) if _is_end(p.row.norm)]
    if not bi or not ei or max(ei) <= min(bi):
        return None
    if any(w in (p.row.norm or "") for p in ps[min(bi) + 1:max(ei)] for w in _GROSS_WORDS):
        return CRes(None, None, "총액·누계액 구분 표")
    bv = [i for i in bi if _val(ps[i], pk) is not None]
    ev = [i for i in ei if _val(ps[i], pk) is not None]
    if len(bv) > 1:
        return CRes(None, None, "기초 행 여러 개")
    if len(ev) > 1:
        return CRes(None, None, "기말 행 여러 개")
    if not bv or not ev:
        return CRes(None, None, "기초·기말 값 없음")
    b, e = bv[0], ev[0]
    if e <= b:
        return CRes(None, None, "기말 행이 기초 행보다 앞")
    between = ps[b + 1:e]
    subs = [i for i, p in enumerate(between) if (p.row.norm or "").endswith("소계")]
    if subs:                                   # c_total 과 같은 소계 규칙: 소계 행들 + 마지막 소계 뒤 행들(항목·소계 이중 합산 방지)
        between = [between[i] for i in subs] + between[subs[-1] + 1:]
    lhs = [_term(ps[b], pk)] + [_term(p, pk) for p in between if _isnum(_val(p, pk))]
    return CRes(lhs, [_term(ps[e], pk)], None)


# ---------------------------------------------------------------- 항목 만들기

class _Item:
    """검증 행 하나. per[pk] = ("ok", lhs, rhs) | ("na", 사유). note = 고정 비고."""
    def __init__(self, group, no, label, note=None, ranges=False):
        self.group, self.no, self.label, self.note, self.ranges = group, no, label, note, ranges
        self.per = OrderedDict()

    def ok(self, pk, lhs, rhs):
        self.per[pk] = ("ok", lhs, rhs)

    def na(self, pk, reason):
        self.per[pk] = ("na", reason)

    def ok_or(self, pk, lhs1, lhs2, rhs):
        """두 식 중 하나만 맞으면 TRUE(OR). 차이는 첫 식 기준."""
        self.per[pk] = ("or", lhs1, lhs2, rhs)


def _mref(ctx, fs_div, sj, std, i):
    rr = ctx["rows_of"].get((fs_div, sj, std)) if (fs_div, sj, std, i) in ctx["filled"] else None
    return C._ref("mFS", rr, C.RAW_FIRST_VALUE_COL + i) if rr else None


def _mterms(ctx, fs_div, i, stds):
    """[(부호, sj, std)] → ([항], [없는 std]). mFS 셀은 ×cfg_unit 로 원."""
    terms, missing = [], []
    for sg, sj, std in stds:
        ref = _mref(ctx, fs_div, sj, std, i)
        if ref:
            terms.append((sg, ref, C.UNIT_NAME))
        else:
            missing.append(std)
    return terms, missing


def _set_pair(it, pk, lhs, lmiss, rhs, rmiss):
    miss = lmiss + rmiss
    if miss:
        it.na(pk, "없음: " + ", ".join(miss))
    else:
        it.ok(pk, lhs, rhs)


def _prev_annual(periods, pk):
    """FY2025 → FY2024 가 열에 있으면 그 순번, 반기 열이거나 직전 연도 열이 없으면 None."""
    m = C._PK.match(pk or "")
    if not m or not m.group(1):
        return None
    prev = f"FY{int(m.group(1)) - 1}"
    return periods.index(prev) if prev in periods else None


def _sce_block(fs_placed, fs_div):
    """rFS 자본변동표 중 열(detail) 정규화가 총계·합계·자본총계·계 인 블록 하나([Placed]). 없거나 여럿이면 None."""
    blocks = OrderedDict()
    for p in fs_placed.get(fs_div) or []:
        b = p.block
        if b.sj_div == "SCE" and b.col_label and N.norm_label(b.col_label) in _SCE_TOTAL:
            blocks.setdefault(id(b), []).append(p)
    return list(blocks.values())[0] if len(blocks) == 1 else None


_A_SPEC = [
    # (번호, 라벨, 좌변 [(부호, sj, std)], 우변 [(부호, sj, std)])
    ("A1", "자산총계 = 부채총계 + 자본총계", [("+", "BS", "자산총계")], [("+", "BS", "부채총계"), ("+", "BS", "자본총계")]),
    ("A2", "유동자산 + 비유동자산 = 자산총계", [("+", "BS", "유동자산"), ("+", "BS", "비유동자산")], [("+", "BS", "자산총계")]),
    ("A3", "유동부채 + 비유동부채 = 부채총계", [("+", "BS", "유동부채"), ("+", "BS", "비유동부채")], [("+", "BS", "부채총계")]),
    ("A4", "지배기업소유주지분 + 비지배지분 = 자본총계", [("+", "BS", "지배기업소유주지분"), ("+", "BS", "비지배지분")], [("+", "BS", "자본총계")]),
    ("A5", "매출액 − 매출원가 = 매출총이익", [("+", "IS", "매출액"), ("-", "IS", "매출원가")], [("+", "IS", "매출총이익")]),
    ("A6", "매출총이익 − 판매비와관리비 = 영업이익", [("+", "IS", "매출총이익"), ("-", "IS", "판매비와관리비")], [("+", "IS", "영업이익")]),
    ("A7", "법인세비용차감전순이익 − 법인세비용 = 당기순이익", [("+", "IS", "법인세비용차감전순이익"), ("-", "IS", "법인세비용")], [("+", "IS", "당기순이익")]),
    ("A8", "당기순이익 + 기타포괄손익 = 총포괄이익", [("+", "IS", "당기순이익"), ("+", "IS", "기타포괄손익")], [("+", "IS", "총포괄이익")]),
    ("A9", None, None, None),           # 아래 _cf_items: 영업+투자+재무+환율 = 기말 − 기초 (표시 방식 중립)
    ("A10", None, None, None),          # 아래 _cf_items: 기초 + 증가(환율 포함/제외 어느 쪽이든) = 기말
    ("A11", "CF 기말현금및현금성자산 = BS 현금및현금성자산", [("+", "CF", "기말현금및현금성자산")], [("+", "BS", "현금및현금성자산")]),
]


_FX_NONE_NOTE = "환율변동효과 행 없음"
_FX_EITHER_NOTE = "순증가에 환율효과 포함·제외 어느 쪽이든 인정; 차이는 환율 제외 식 기준"


def _cf_items(ctx, fs_div, no):
    """현금흐름표 A9·A10 — 원문 표시 방식(순증가에 환율변동효과 포함 여부)에 중립. 데이터로 식을 고르지 않는다.
    A9: 영업 + 투자 + 재무 (+ 환율변동효과, 있으면) = 기말현금 − 기초현금.
    A10: 환율변동효과가 있으면 OR(기초 + 증가 = 기말, 기초 + 증가 + 환율 = 기말), 없으면 기초 + 증가 = 기말."""
    periods = ctx["periods"]
    fx = ("+", "CF", "환율변동효과")
    if no == "A9":
        it = _Item("A", no, "영업+투자+재무+환율변동효과 = 기말현금 − 기초현금")
        for i, pk in enumerate(periods):
            l_, lm = _mterms(ctx, fs_div, i, [("+", "CF", "영업활동현금흐름"), ("+", "CF", "투자활동현금흐름"), ("+", "CF", "재무활동현금흐름")])
            f_, fm = _mterms(ctx, fs_div, i, [fx])
            if fm and not lm:
                it.note = _FX_NONE_NOTE
            r_, rm = _mterms(ctx, fs_div, i, [("+", "CF", "기말현금및현금성자산"), ("-", "CF", "기초현금및현금성자산")])
            _set_pair(it, pk, l_ + f_, lm, r_, rm)
        return it
    it = _Item("A", no, "기초현금 + 현금증가(환율 포함/제외) = 기말현금")
    for i, pk in enumerate(periods):
        l_, lm = _mterms(ctx, fs_div, i, [("+", "CF", "기초현금및현금성자산"), ("+", "CF", "현금및현금성자산의증가")])
        r_, rm = _mterms(ctx, fs_div, i, [("+", "CF", "기말현금및현금성자산")])
        f_, fm = _mterms(ctx, fs_div, i, [fx])
        if lm or rm:
            _set_pair(it, pk, l_, lm, r_, rm)
        elif fm:
            it.ok(pk, l_, r_)
        else:
            it.ok_or(pk, l_, l_ + f_, r_)
            it.note = _FX_EITHER_NOTE
    return it


def _a_items(ctx, fs_div, fs_placed):
    periods = ctx["periods"]
    out = []
    for no, label, lhs, rhs in _A_SPEC:
        if no == "A4" and fs_div != "CFS":
            continue
        if no in ("A9", "A10"):
            out.append(_cf_items(ctx, fs_div, no))
            continue
        it = _Item("A", no, label)
        for i, pk in enumerate(periods):
            l_, lm = _mterms(ctx, fs_div, i, lhs)
            r_, rm = _mterms(ctx, fs_div, i, rhs)
            _set_pair(it, pk, l_, lm, r_, rm)
        out.append(it)
    it = _Item("A", "A12", "CF 기초현금및현금성자산 = 직전 연도 BS 현금및현금성자산")
    for i, pk in enumerate(periods):
        j = _prev_annual(periods, pk)
        if j is None:
            it.na(pk, "직전 연도 없음")
            continue
        l_, lm = _mterms(ctx, fs_div, i, [("+", "CF", "기초현금및현금성자산")])
        r_, rm = _mterms(ctx, fs_div, j, [("+", "BS", "현금및현금성자산")])
        _set_pair(it, pk, l_, lm, r_, [f"{x}({periods[j]})" for x in rm])
    out.append(it)
    sce = _sce_block(fs_placed, fs_div)
    it13 = _Item("A", "A13", "자본변동표 기말 총계 = BS 자본총계")
    it14 = _Item("A", "A14", "자본변동표 총계 열 당기순이익 = IS 당기순이익")
    total_pks = []                                   # A13 에 '자본총계' 행을 쓴 기간(비고)
    for i, pk in enumerate(periods):
        if sce is None:
            it13.na(pk, "없음: 자본변동표 총계 블록(하나로 정해지지 않음)")
            it14.na(pk, "없음: 자본변동표 총계 블록(하나로 정해지지 않음)")
            continue
        # 기간마다: 기말 행(…기말…) 또는 '자본총계' 행 중 그 기간 값 있는 행이 하나면 그 행. API 원천 상장사는 연차 기말이
        # '자본총계' 행에, 반기 기말이 '기말자본' 행에 갈려 있다(삼성전자·휴젤 rFS_연결).
        ends = [p for p in sce if N.sce_row_norm(p.row.label) == "기말" or N.norm_label(p.row.label) == "자본총계"]
        vals = [p for p in ends if _val(p, pk) is not None]
        r_, rm = _mterms(ctx, fs_div, i, [("+", "BS", "자본총계")])
        if len(vals) > 1:
            it13.na(pk, "자본변동표 기말 행 여러 개" + (f"; 없음: {', '.join(rm)}" if rm else ""))
        elif not vals:
            _set_pair(it13, pk, [], ["자본변동표 기말 행"], r_, rm)
        else:
            _set_pair(it13, pk, [_term(vals[0], pk)], [], r_, rm)
            if not rm and N.norm_label(vals[0].row.label) == "자본총계":
                total_pks.append(pk)
        m = C._PK.match(pk or "")
        if not (m and m.group(1)):
            it14.na(pk, "반기 자본변동표 제외")        # 반기 열: 반기 자본변동표 행이 연차와 달라 비교하지 않는다(연차만)
            continue
        nis = [p for p in sce if N.norm_label(p.row.label).startswith(_NET_INCOME) and _val(p, pk) is not None]
        ni = nis[0] if len(nis) == 1 else None       # 기간마다: 당기순손실(FY2021~)·당기순이익(FY2025) 행이 따로 있어도 값 있는 행이 하나면
        r_, rm = _mterms(ctx, fs_div, i, [("+", "IS", "당기순이익")])
        if ni is None or _val(ni, pk) is None:
            _set_pair(it14, pk, [], ["자본변동표 당기순이익 행"], r_, rm)
        else:
            _set_pair(it14, pk, [_term(ni, pk)], [], r_, rm)
    if total_pks:
        ok_pks = [pk for pk in periods if (it13.per.get(pk) or ("na",))[0] == "ok"]
        it13.note = "자본총계 행" if total_pks == ok_pks else f"{'·'.join(total_pks)} 는 자본총계 행"
    out += [it13, it14]
    return out


def _b_items(ctx, fs_div, note_placed):
    periods = ctx["periods"]
    out = []

    def add_note(it, text):
        if text and text not in (it.note or ""):
            it.note = f"{it.note}; {text}" if it.note else text

    def pair(no, label, need, select, rhs_of, note=None):
        """select(pk) → 주석 Placed|None|Pick, rhs_of(i, pk, pick) → ([항], [없는 것], 추가 비고|None).
        Pick.note(선택 근거)는 고른 기간 전부에 해당하면 그대로, 일부면 'FY2021·FY2022 는 소계' 꼴로 비고에."""
        it = _Item("B", no, label, note)
        notes, picked = OrderedDict(), []
        for i, pk in enumerate(periods):
            pick = _as_pick(select(pk))
            r_, rm, extra = rhs_of(i, pk, pick)
            add_note(it, extra)
            if pick:
                picked.append(pk)
                if pick.note:
                    notes.setdefault(pick.note, []).append(pk)
            if not pick and pick.reason:
                it.na(pk, pick.reason + (f"; 없음: {', '.join(rm)}" if rm else ""))
            else:
                _set_pair(it, pk, [_term(p, pk) for p in pick.rows], [] if pick else [f"주석 {need}"], r_, rm)
        for text, pks in notes.items():
            add_note(it, text if len(pks) == len(picked) else f"{'·'.join(pks)} 는 {text}")
        out.append(it)

    def mfs(*stds):
        return lambda i, pk, pick: _mterms(ctx, fs_div, i, [("+", sj, std) for sj, std in stds]) + (None,)

    def sel(topic, **kw):
        return lambda pk: _note_row(note_placed, fs_div, topic, pk, **kw)

    first = _first

    pair("B1", "판관비 합계 = IS 판매비와관리비", "판관비 합계 행", sel("판관비", how="total"), mfs(("IS", "판매비와관리비")))
    pair("B2", "성격별비용 합계 = IS 매출원가 + 판매비와관리비", "성격별비용 합계 행", sel("성격별비용", how="total"),
         mfs(("IS", "매출원가"), ("IS", "판매비와관리비")))
    pair("B3", "유형자산 변동표 기말 = BS 유형자산", "유형자산 변동표 기말", sel("유형자산", how="end", where=main_ppe),
         mfs(("BS", "유형자산")))

    def rhs_intangible(i, pk, pick):
        gw = _mref(ctx, fs_div, "BS", "영업권", i)
        stds = [("+", "BS", "무형자산")] + ([("+", "BS", "영업권")] if gw else [])
        t, m = _mterms(ctx, fs_div, i, stds)
        return t, m, ("BS 영업권 행 있음: 우변 = 무형자산 + 영업권" if gw else None)
    pair("B4", "무형자산 변동표 기말 = BS 무형자산", "무형자산 변동표 기말", sel("무형자산", how="end", where=main_intangible),
         rhs_intangible)
    pair("B5", "금융수익 주석 합계 = IS 금융수익", "금융수익 합계 행",
         sel("금융손익", how="total", where=lambda b: not is_cost(b)), mfs(("IS", "금융수익")))
    pair("B6", "금융비용 주석 합계 = IS 금융비용", "금융비용 합계 행",
         sel("금융손익", how="total", where=is_cost), mfs(("IS", "금융비용")))
    pair("B7", "매출 구분 합계 = IS 매출액", "매출 구분 합계 행",
         lambda pk: _revenue_pick(note_placed, fs_div, pk), mfs(("IS", "매출액")))

    def rhs_borrowings(i, pk, pick):
        """우변 = BS 단기차입금 + 유동성장기부채 + 장기차입금 중 있는 것(+ 사채는 고른 블록 캡션에 '사채' 가 있을 때만)."""
        bond = bool(pick) and has(pick.rows[0].block, "사채")
        stds = list(_BORROW_STDS) + (["사채"] if bond else [])
        t, m = _mterms(ctx, fs_div, i, [("+", "BS", s) for s in stds])
        if not t:
            return [], stds, None
        return t, [], "우변 = " + " + ".join(s for s in stds if s not in m) + ("" if bond else " (사채 제외)")
    pair("B8", "차입금 주석 합계 = BS 단기차입금 + 유동성장기부채 + 장기차입금", "차입금 합계 행",
         sel("차입금", how="total_all", where=lambda b: has(b, "차입금")), rhs_borrowings)
    pair("B9", "리스부채 주석 = BS 유동리스부채 + 비유동리스부채", "리스부채 행",
         first(sel("리스", how="row", labels=("리스부채",), where=lambda b: has(b, "리스부채")),
               sel("리스", how="row", labels=("합계",),                                         # 만기분석(할인 전 현금흐름) 표 합계 제외
                   where=lambda b: has(b, "리스부채") and not has(b, *_LEASE_TOTAL_EXCLUDE)),
               sel("리스", how="sub", labels=("리스부채",))),                                     # 소제목 '리스부채' 아래 합계
         mfs(("BS", "유동리스부채"), ("BS", "비유동리스부채")))
    pair("B10", "사용권자산 변동표 기말 = BS 사용권자산", "사용권자산 변동표 기말",
         first(sel("리스", how="end", where=is_rou), sel("유형자산", how="end", where=is_rou),   # 유형자산 주석 안의 사용권자산 표도
               sel("리스", how="sub", labels=("사용권자산",))),                                  # 변동표가 없으면 소제목 아래 합계
         mfs(("BS", "사용권자산")))
    pair("B11", "순확정급여부채 주석 = BS 순확정급여부채", "순확정급여부채 행",
         sel("퇴직급여", how="row", labels=("순확정급여부채", "순확정급여부채(자산)")), mfs(("BS", "순확정급여부채")))

    def rhs_provision(i, pk, pick):
        if _mref(ctx, fs_div, "BS", "충당부채", i):
            return _mterms(ctx, fs_div, i, [("+", "BS", "충당부채")]) + (None,)
        t, m = _mterms(ctx, fs_div, i, [("+", "BS", "유동충당부채"), ("+", "BS", "비유동충당부채")])
        if not t:
            return [], ["충당부채"], None
        return t, [], ("우변 = " + " + ".join(s for s in ("유동충당부채", "비유동충당부채") if s not in m))
    pair("B12", "충당부채 변동표 기말 = BS 충당부채(또는 유동 + 비유동)", "충당부채 변동표 기말", sel("충당부채", how="end"), rhs_provision)
    pair("B13", "법인세비용 주석 = IS 법인세비용", "법인세비용 행",
         sel("법인세", how="last", pattern=_TAX_ROW, exclude=_TAX_EXCLUDE), mfs(("IS", "법인세비용")))
    pair("B14", "자본금 주석 = BS 자본금", "자본금 행", lambda pk: _capital_pick(note_placed, fs_div, pk), mfs(("BS", "자본금")))
    return out


def _short(text, n=60):
    return text if len(text) <= n else text[:n - 1] + "…"


def _caption_summary(b):
    """C 항목의 블록 요약: 캡션에 기간·단위 표식 말고 내용이 있으면(caption_sig 비어 있지 않음) 캡션, 아니면 주석 제목
    (`주석 N.` 접두 없이). 열 블록이면 ` — 열 라벨`."""
    base = b.title if L.caption_sig(b.title or "") else (b.note_title or b.title or "")
    return f"{base} — {b.col_label}" if b.col_label else base


def _c_items(ctx, fs_div, note_placed):
    periods = ctx["periods"]
    blocks = OrderedDict()
    for p in note_placed:
        if p.sheet != C.ETC_SHEET and p.block.fs_div == fs_div:
            blocks.setdefault((p.sheet, id(p.block)), []).append(p)
    out, n, used = [], 0, {}
    for (sheet, _), ps in blocks.items():
        for kind, fn in (("합계", c_total), ("변동", c_roll)):
            if not periods or fn(ps, periods[0]) is None:      # 검산 대상 행이 없는 블록(기간과 무관)
                continue
            n += 1
            label = f"{sheet} · {_short(_caption_summary(ps[0].block))} · {kind}"
            k = used.get(label, 0) + 1
            used[label] = k
            it = _Item("C", f"C{n}", label if k == 1 else f"{label} #{k}", ranges=True)
            for pk in periods:
                res = fn(ps, pk)
                if res is None or res.lhs is None:
                    it.na(pk, res.reason if res else "값 없음")
                else:
                    it.ok(pk, res.lhs, res.rhs)
            out.append(it)
    return out


# ---------------------------------------------------------------- 시트

def _reasons(it, periods):
    by = OrderedDict()
    for pk in periods:
        st = it.per.get(pk)
        if st and st[0] == "na":
            by.setdefault(st[1], []).append(pk)
    parts = [r if len(pks) == len(periods) else f"{r}({'·'.join(pks)})" for r, pks in by.items()]
    if it.note:
        parts.append(it.note)
    return "; ".join(parts)


def build(col, fs_placed, note_placed, ctx):
    """검증 시트. ctx = compose._mfs_sheet 가 돌려준 {"periods", "rows_of", "filled"}.
    반환 (Sheet, {fs_div: (첫 검증 행, 마지막 검증 행)}) — mFS 요약 COUNTIF 범위."""
    s = Sheet(SHEET)
    periods = list(ctx["periods"])
    n = len(periods)
    D, E = C.RAW_LABEL_COL, C.RAW_FIRST_VALUE_COL
    diff0 = E + n + 1
    lcol, rcol, ncol = diff0 + n, diff0 + n + 1, diff0 + n + 2
    s.set(1, D, "검증 (TRUE 여야 함) — 양변을 원으로 환산해 천원 반올림 후 비교, 차이는 원 단위", "title")
    for c, h in ((1, "구분"), (2, "번호"), (D, "항목"), (lcol, "좌변 참조"), (rcol, "우변 참조"), (ncol, "비고")):
        s.set(2, c, h, "hdr")
        s.set(3, c, "", "hdr")
    for i, pk in enumerate(periods):
        s.set(2, E + i, pk, "hdr"); s.set(3, E + i, "TRUE/FALSE", "hdr")
        s.set(2, diff0 + i, f"차이 {pk}", "hdr"); s.set(3, diff0 + i, "원(좌변 − 우변)", "hdr")
    for r, side in ((2, "top"), (3, "bottom")):
        for c in range(1, ncol + 1):
            cell = s.cells.get((r, c))
            if cell is None:
                s.set(r, c, "", "hdr")
            cell = s.cells[(r, c)]
            cell.border = side
    r = 5
    ranges = {}
    for fs_div in C._fs_divs(col):
        if not fs_placed.get(fs_div):
            continue
        name = C._fs_name(fs_div)
        first = last = None
        for group, title, items in (("A", "재무제표 내부 정합", _a_items(ctx, fs_div, fs_placed)),
                                    ("B", "주석 ↔ 재무제표", _b_items(ctx, fs_div, note_placed)),
                                    ("C", "주석 합계 검산", _c_items(ctx, fs_div, note_placed))):
            s.set(r, D, f"{group}. {title} — {name}", "section")
            r += 1
            for it in items:
                s.set(r, 1, it.group, "text"); s.set(r, 2, it.no, "text"); s.set(r, D, it.label, "text")
                texts = None
                for i, pk in enumerate(periods):
                    st = it.per.get(pk) or ("na", "값 없음")
                    if st[0] in ("ok", "or"):
                        f, d = check_formula(st[1], st[2]) if st[0] == "ok" else or_formula(st[1], st[2], st[3])
                        s.set(r, E + i, f, "check", formula=True)
                        s.set(r, diff0 + i, d, "num", formula=True)
                        if texts is None:
                            left = ref_text(st[1], it.ranges)
                            if st[0] == "or":                     # 두 갈래: `mFS!E10+mFS!E11 | +mFS!E12`
                                left += " | " + "".join(sg + _plain(ref) for sg, ref, _ in st[2][len(st[1]):])
                            texts = (left, ref_text(st[-1], it.ranges))
                    else:
                        s.set(r, E + i, NA, "italic")
                if texts:
                    s.set(r, lcol, texts[0], "text"); s.set(r, rcol, texts[1], "text")
                note = _reasons(it, periods)
                if note:
                    s.set(r, ncol, note, "italic")
                first = first or r
                last = r
                r += 1
            r += 1
        ranges[fs_div] = (first, last)
    s.col_widths = {1: 5, 2: 6, 3: 2, D: 52, lcol: 34, rcol: 34, ncol: 44}
    for i in range(n):
        s.col_widths[E + i] = C.VALUE_COL_WIDTH
        s.col_widths[diff0 + i] = C.VALUE_COL_WIDTH
    s.col_widths[E + n] = 2
    s.freeze = f"{C._col_letter(E)}4"
    return s, ranges


def row_ranges(s):
    """만들어진 검증 시트에서 fs_div 별 (첫 검증 행, 마지막 검증 행)을 다시 읽는다(테스트·확인용)."""
    out, cur = {}, None
    for r in range(1, s.max_row() + 1):
        sec = s.cells.get((r, C.RAW_LABEL_COL))
        if sec is not None and sec.fmt == "section":
            cur = "CFS" if str(sec.v).endswith("— 연결") else "OFS"
            continue
        no = s.cells.get((r, 2))
        if cur and no is not None and no.v not in (None, "", "번호"):
            a, b = out.get(cur, (r, r))
            out[cur] = (min(a, r), max(b, r))
    return out
