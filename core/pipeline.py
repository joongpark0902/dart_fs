"""수집 오케스트레이션: 보고서 선택 → API/원문 → 레코드 → 재작성 기준 선택 → 주석."""
import re
from dataclasses import dataclass, field, asdict
from datetime import date
from core import fs_parser as F, note_parser as NP, note_std as S, selector as SEL, xml_doc as X
from core import normalize as N
from core.model import NoteRecord

_FISCAL = re.compile(r"제\s*(\d{1,3})\s*기")
_AUDIT_TYPES = ("감사보고서", "연결감사보고서", "반기검토보고서", "반기연결검토보고서")


@dataclass
class BuildConfig:
    corp_code: str
    corp_name: str
    as_of: date
    start_year: int
    end_year: int
    include_half: bool
    fs_divs: list
    unit_display: int = 0            # 0 = 자동(공시 단위 판정) | 1·1000·1000000 = 직접 선택
    note_include: object = None      # None=전부 | {"recommended": True} | {title_norm: bool}
    preset: str = "default"

    def to_dict(self):
        d = asdict(self)
        d["as_of"] = self.as_of.isoformat()
        return d

    @staticmethod
    def from_dict(d):
        as_of = d.get("as_of")
        if not as_of:
            raise ValueError("as_of 필요")
        as_of = date.fromisoformat(as_of) if isinstance(as_of, str) else as_of
        return BuildConfig(d["corp_code"], d.get("corp_name", ""), as_of, int(d["start_year"]), int(d["end_year"]),
                           bool(d.get("include_half", False)), list(d.get("fs_divs") or ["OFS"]),
                           _unit_of(d.get("unit_display")), d.get("note_include"), d.get("preset") or "default")


def _unit_of(v):
    """설정의 표시 단위: 없거나 빈 값·0 = 자동(0)."""
    return 0 if v in (None, "") else int(v)


@dataclass
class Collected:
    cfg: object
    plan: object
    periods: list
    fs_records: list = field(default_factory=list)
    note_records: list = field(default_factory=list)
    toc: list = field(default_factory=list)
    restate_diffs: list = field(default_factory=list)
    auditor_rows: list = field(default_factory=list)
    logs: list = field(default_factory=list)
    mapping: object = None
    # (fs_div, sj_div) -> 최신 출처의 단위 / ("fiscal", period_key) -> 제 N 기
    units: dict = field(default_factory=dict)


def fiscal_from_api(items):
    for it in items or []:
        m = _FISCAL.search(it.get("thstrm_nm") or "")
        if m:
            return int(m.group(1))
    return None


def _log(col, log, msg):
    """메시지를 col.logs 와 콜백 양쪽에 남긴다."""
    if col is not None:
        col.logs.append(msg)
    if log is not None:
        log(msg)


def _plan(cfg, client):
    disclosures = client.list_disclosures(cfg.corp_code, f"{cfg.start_year}0101", cfg.as_of.strftime("%Y%m%d"))
    return SEL.plan_reports(disclosures, cfg.start_year, cfg.end_year, cfg.as_of, cfg.include_half, set(cfg.fs_divs))


def _audit_xml_path(client, cfg, report, fs_div):
    docs = client.documents(cfg.corp_code, report.rcept_no)
    key = "audit_cfs" if fs_div == "CFS" else "audit_ofs"
    return docs.get(key) or (docs.get("main") if report.rtype in _AUDIT_TYPES else None)


def toc_preview(cfg, client, mapping, log, plan=None):
    plan = plan or _plan(cfg, client)
    if plan.base_report is None:
        raise RuntimeError("기준 보고서(연간)가 없습니다. 연도 범위·기준일을 확인하세요.")
    fs_div = "CFS" if "CFS" in cfg.fs_divs else "OFS"
    path = _audit_xml_path(client, cfg, plan.base_report, fs_div)
    if not path:
        raise RuntimeError(f"{plan.base_report.report_nm} 에서 {fs_div} 감사보고서 원문을 찾지 못했습니다.")
    toc = NP.note_toc(X.stream(X.load_xml(path)))
    S.assign_topics(toc, mapping)
    log(f"기준 보고서 {plan.base_report.report_nm} ({plan.base_report.rcept_no}) 주석 {len(toc)}개")
    return plan.base_report, toc


def _wanted(entry, cfg):
    inc = cfg.note_include
    if inc is None:
        return True
    if isinstance(inc, dict) and inc.get("recommended") is True and len(inc) == 1:
        return entry.recommended
    return bool(inc.get(entry.title_norm, inc.get(entry.title, False)))


def _fs_from_report(cfg, client, report, fs_div, col, log):
    """한 보고서·한 fs_div 의 FsRecord 전부(cur/prev/prev2 포함), 단위, 기간 사전."""
    reprt = "11011" if report.kind == "annual" else "11012"
    items = None
    if report.rtype in ("사업보고서", "반기보고서"):
        items = client.fnltt_all(cfg.corp_code, report.period_end.year, reprt, fs_div, expect_rcept_no=report.rcept_no,
                                 filed_on=report.rcept_dt)
        w = getattr(client, "last_warning", None)
        if w:
            _log(col, log, f"{report.report_nm} {fs_div}: {w}")
    if items:
        periods = SEL.fiscal_periods_for(report, fiscal_from_api(items))
        _log(col, log, f"{report.report_nm} {fs_div}: API {len(items)}행")
        recs = F.api_records(items, fs_div, periods, report.rcept_no)
        return recs, {r.sj_div: 1 for r in recs}, periods       # OpenDART API 금액은 항상 원
    path = _audit_xml_path(client, cfg, report, fs_div)
    if not path:
        _log(col, log, f"{report.report_nm} {fs_div}: 원문 없음 → 건너뜀")
        return [], {}, {}
    root = X.load_xml(path)
    periods = SEL.fiscal_periods_for(report, None)
    recs, info = F.xml_statements(root, fs_div, periods, report.rcept_no)
    fc = info["fiscal"].get("cur")
    if fc:      # 기간 라벨에 제 N 기를 붙이기 위해 한 번 더
        periods = SEL.fiscal_periods_for(report, fc)
        recs, info = F.xml_statements(root, fs_div, periods, report.rcept_no)
    for w in info["warnings"]:
        _log(col, log, f"{report.report_nm} {fs_div}: {w}")
    _log(col, log, f"{report.report_nm} {fs_div}: 원문 {len(recs)}행")
    return recs, info["units"], periods


def _fs_cmp(r):
    """재무제표 재작성 비교 키. 같은 계정명이 두 번 나오는 경우(당기순이익·총포괄이익의 '지배기업소유주지분')
    API 는 행 순서가 연도마다 달라 account_id 로 구분한다. 원문(XML)은 account_id 가 없어 계정명·순번으로."""
    return (r.label_norm, r.detail or "", getattr(r, "account_id", None) or "")


def _note_cmp(r):
    """주석 재작성 비교 키. 표 번호는 연도마다 밀리므로(전기 변동표·각주표가 끼어듦) 쓰지 않고
    (행, 열) 과 그 키의 등장 순번으로 맞춘다."""
    return (r.row_norm, r.col_label or "")


def _value(r):
    return r.amount if hasattr(r, "amount") else r.value


def _pick_restated(periods, sources_of, by_report, group_fn, cmp_fn, name_fn):
    """(기간, group_fn) 키마다 sources 순서대로 첫 번째로 레코드가 있는 출처를 채택하고,
    그다음 출처(레코드가 있는)와 값이 다른 항목을 재작성 차이로 기록."""
    chosen, diffs = [], []
    for pk in periods:
        per_key = {}         # group key -> [(rcept_no, records)]
        for rc, role in sources_of.get(pk, []):
            bucket = {}
            for r in by_report.get(rc, []):
                if r.period_key == pk and r.role == role and getattr(r, "raw_cell", None) is None:
                    bucket.setdefault(group_fn(r), []).append(r)
            for k, recs in bucket.items():
                per_key.setdefault(k, []).append((rc, recs))
        for k, got in per_key.items():
            chosen.extend(got[0][1])
            if len(got) < 2:
                continue
            first, cnt = {}, {}
            for r in got[0][1]:
                ck = cmp_fn(r)
                n = cnt[ck] = cnt.get(ck, -1) + 1
                first[(ck, n)] = _value(r)
            cnt = {}
            for r in got[1][1]:
                ck = cmp_fn(r)
                n = cnt[ck] = cnt.get(ck, -1) + 1
                a, v = first.get((ck, n)), _value(r)
                if a is not None and v is not None and a != v:
                    diffs.append({"period": pk, "key": " / ".join(str(x) for x in (*k, *name_fn(r))),
                                  "채택": a, "원보고서": v, "채택출처": got[0][0], "원보고서출처": got[1][0]})
    return chosen, diffs


def collect(cfg, client, mapping, log):
    plan = _plan(cfg, client)
    periods = [p.key for p in plan.periods]
    col = Collected(cfg, plan, periods, mapping=mapping)
    col.logs.extend(plan.log)
    if not periods:
        raise RuntimeError("대상 기간이 없습니다.")
    # 1) 재무제표: 필요한 보고서 × fs_div
    by_report = {}
    needed = {rc for lst in plan.sources.values() for rc, _ in lst}
    for rc in sorted(needed, key=lambda r: (plan.reports[r].period_end, r)):
        report = plan.reports[rc]
        for fs_div in cfg.fs_divs:
            if fs_div not in report.fs_divs:
                continue
            try:
                recs, units, pmap = _fs_from_report(cfg, client, report, fs_div, col, log)
            except Exception as e:
                _log(col, log, f"{report.report_nm} {fs_div}: 처리 실패 — {e}")
                continue
            by_report.setdefault(rc, []).extend(recs)
            for sj, u in units.items():
                col.units[(fs_div, sj)] = u
            for p in pmap.values():
                if p.fiscal_no:
                    col.units[("fiscal", p.key)] = p.fiscal_no
    col.fs_records, col.restate_diffs = _pick_restated(
        periods, plan.sources, by_report,
        lambda r: (r.fs_div, r.sj_div), _fs_cmp,
        lambda r: (r.label_norm, r.detail or ""))
    # 2) 주석: 연간 보고서만, 기준 보고서 목차로 선택
    if plan.base_report is not None:
        base, toc = toc_preview(cfg, client, mapping, log, plan)
        _log(col, None, f"기준 보고서 {base.report_nm} ({base.rcept_no})")
        col.toc = toc
        for e in toc:
            e.include = _wanted(e, cfg)
        wanted = {e.title_norm for e in toc if e.include}
        note_by_report = {}
        for rc in sorted({rc for lst in plan.note_sources.values() for rc, _ in lst}):
            report = plan.reports[rc]
            for fs_div in cfg.fs_divs:
                if fs_div in report.fs_divs:
                    try:
                        recs = _notes_for_report(cfg, client, report, fs_div, wanted, mapping, col, log)
                    except Exception as e:
                        _log(col, log, f"{report.report_nm} {fs_div}: 처리 실패 — {e}")
                        continue
                    note_by_report.setdefault(rc, []).extend(recs)
        annual = [p for p in periods if p.startswith("FY")]
        col.note_records, note_diffs = _pick_restated(
            annual, plan.note_sources, note_by_report,
            lambda r: (r.fs_div, N.norm_label(r.note_title)),
            _note_cmp, lambda r: (r.row_label, r.col_label or ""))
        col.restate_diffs.extend(note_diffs)
        seen = set()      # 기간 표시 없는 원표: 연도별로 공시된 그대로 모두 보존
        for rc in sorted({rc for lst in plan.note_sources.values() for rc, _ in lst}):
            for r in note_by_report.get(rc, []):
                k = (r.rcept_no, r.fs_div, r.note_no, r.table_idx, r.raw_cell, r.row_label, r.col_label)
                if r.raw_cell is not None and r.period_key in annual and k not in seen:
                    seen.add(k)
                    col.note_records.append(r)
        for e in toc:
            if e.include and not any(N.norm_label(r.note_title) == e.title_norm for r in col.note_records):
                _log(col, log, f"주석 '{e.title}': 어느 연도에서도 표를 만들지 못함")
    return col


def _notes_for_report(cfg, client, report, fs_div, wanted, mapping, col, log=None):
    """한 보고서·fs_div 의 선택된 주석 레코드. 감사인 행은 col.auditor_rows 에 추가."""
    rc = report.rcept_no
    path = _audit_xml_path(client, cfg, report, fs_div)
    if not path:
        _log(col, log, f"{report.report_nm} {fs_div}: 주석 원문 없음")
        return []
    root = X.load_xml(path)
    nodes = X.stream(root)
    entries = NP.note_toc(nodes)
    S.assign_topics(entries, mapping)
    _, info = F.xml_statements(root, fs_div, SEL.fiscal_periods_for(report, None), rc)
    per = SEL.fiscal_periods_for(report, info["fiscal"].get("cur"))
    out = []
    for e in entries:
        if e.title_norm not in wanted:
            continue
        for nt in NP.note_tables(nodes, e):
            recs, kind = NP.table_records(nt, e, per, fs_div, rc)
            if kind == "other":
                _log(col, log, f"{report.report_nm} 주석 {e.no} {e.title} 표{nt.idx+1}: 기간 표시 없음 → rN_기타")
                nc = max(nt.grid.ncols, 1)
                recs = [_raw_note_record(e, nt, per["cur"], fs_div, rc, rl, cl, v, (i // nc, i % nc))
                        for i, (rl, cl, v) in enumerate(_raw_cells(nt))]
            out.extend(recs)
    col.auditor_rows.append(_auditor_row(report, fs_div, nodes))
    return out


def _raw_cells(nt):
    """other 로 분류된 표를 셀 단위로(행 라벨, 열 라벨, 값). rN_기타 에 원표 그대로 쓰기 위한 재료."""
    g = nt.grid
    hdr = g.header[-1] if g.header else [f"열{i+1}" for i in range(g.ncols)]
    out = []
    for ri, row in enumerate(g.body):
        sub = NP.is_subheading(row, [0])           # 소제목 행: 라벨 열만 남기고 값 칸은 비운다
        for c in range(g.ncols):
            first = N.clean_text(row[0]) if row and row[0] else ""
            v = row[c] if c < len(row) else ""
            if c and sub:
                v = ""
            elif isinstance(v, str) and v.strip().endswith(("%", "%)")):
                v = N.parse_percent(v) if N.parse_percent(v) is not None else v
            out.append((first or f"행{ri+1}",
                        N.clean_text(hdr[c]) if c < len(hdr) else f"열{c+1}", v))
    return out


def _raw_note_record(e, nt, period, fs_div, rc, row_label, col_label, text, cell=None):
    return NoteRecord(period.key, fs_div, e.no, e.title, e.topic, nt.idx, nt.caption, row_label, N.norm_label(row_label),
                      col_label, text, N.parse_unit(nt.caption) or 1, rc, "cur", cell,
                      N.parse_unit(nt.caption) is not None)


_OPINION_KEYS = ("공정하게 표시", "적정", "한정", "부적정", "의견거절", "표명하지")


def _auditor_row(report, fs_div, nodes):
    """감사의견 문장(독립된 감사인의 감사보고서 절의 첫 의견 문단)과 감사인(회계법인·감사반 문단)을 한 줄로."""
    opinion, firm, firm_any = "", "", ""
    for n in nodes:
        if n.kind != "p":
            continue
        t = X.text(n.elem)
        in_sec = "독립된 감사인의 감사보고서" in (n.title or "")
        if not opinion and in_sec and "의견" in t and any(k in t for k in _OPINION_KEYS):
            opinion = t
        if re.search(r"(회계법인|감사반)", t) and len(t) < 80:
            if in_sec and not firm:
                firm = t
            elif not firm_any:
                firm_any = t
    firm = firm or firm_any
    return [report.period_end.isoformat(), report.report_nm, fs_div, report.rcept_no, firm, opinion]
