"""공통 자료구조. 파서 → layout → excel 사이에 오가는 것은 전부 여기 정의된 것만."""
from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass(frozen=True)
class Period:
    key: str                 # "FY2025" | "2026H1"
    end: date                # 2025-12-31 | 2026-06-30
    kind: str                # "annual" | "half"
    fiscal_no: Optional[int] = None   # 제 N 기
    label: str = ""          # "제 28 기" / "제 26 기 반기"

    @staticmethod
    def annual(year, fiscal_no=None):
        return Period(f"FY{year}", date(year, 12, 31), "annual", fiscal_no,
                      f"제 {fiscal_no} 기" if fiscal_no else "")

    @staticmethod
    def half(year, fiscal_no=None):
        return Period(f"{year}H1", date(year, 6, 30), "half", fiscal_no,
                      f"제 {fiscal_no} 기 반기" if fiscal_no else "")


@dataclass
class ReportRef:
    rcept_no: str
    report_nm: str
    rcept_dt: date
    kind: str                # "annual" | "half"
    rtype: str               # "사업보고서" | "반기보고서" | "감사보고서" | "연결감사보고서" | "반기검토보고서" | "반기연결검토보고서"
    period_end: date
    fs_divs: tuple           # ("CFS","OFS") 사업·반기보고서 / ("OFS",) 감사보고서 / ("CFS",) 연결감사보고서
    corrected: bool = False
    superseded: tuple = ()   # 같은 보고서의 이전 접수번호들(정정 전)


@dataclass
class FsRecord:
    period_key: str
    fs_div: str              # "CFS" | "OFS"
    sj_div: str              # "BS" | "IS" | "CIS" | "CF" | "SCE"
    ord: int
    label: str               # 공시 그대로
    label_norm: str
    amount: object           # int | float | None
    unit: int = 1
    detail: Optional[str] = None     # SCE 자본 구성요소 열 라벨
    account_id: Optional[str] = None
    rcept_no: str = ""
    source: str = "API"      # "API" | "XML"
    role: str = "cur"        # 이 값이 보고서의 cur/prev/prev2 중 어느 열에서 왔는지


@dataclass
class NoteRecord:
    period_key: str
    fs_div: str
    note_no: str
    note_title: str
    topic: Optional[str]
    table_idx: int           # 주석 안 표 순번(0부터)
    caption: str
    row_label: str
    row_norm: str
    col_label: Optional[str]
    value: object            # int | float | str | None
    unit: int
    rcept_no: str
    role: str = "cur"
    raw_cell: Optional[tuple] = None  # 기간 표시 없는 원표의 (행, 열) 위치 — 있으면 rN_기타 에 원표 그대로
    unit_known: bool = True           # 원표 캡션에 '(단위' 표기가 있었나(없으면 unit=1 은 가정값 → '원문 미표기')


@dataclass
class NoteEntry:
    no: str
    title: str
    title_norm: str
    topic: Optional[str]
    recommended: bool
    start: int               # stream 인덱스
    end: int
    table_count: int = 0
    include: bool = True


@dataclass
class Row:
    label: str
    norm: str
    values: dict = field(default_factory=dict)   # period_key -> value
    flags: list = field(default_factory=list)    # "미매핑" 등
    account_id: Optional[str] = None             # 재무제표 행의 account_id (같은 라벨이 여러 번이어도 행별로)


@dataclass
class Block:
    title: str               # 블록 제목(재무제표명 / 주석 캡션)
    unit: int
    periods: list            # period_key 순서(열)
    rows: list               # list[Row]
    col_label: Optional[str] = None
    note_no: str = ""
    note_title: str = ""
    topic: Optional[str] = None
    fs_div: str = ""
    sj_div: str = ""
    meta: dict = field(default_factory=dict)
