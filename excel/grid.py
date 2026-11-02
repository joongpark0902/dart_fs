"""엑셀 쓰기 전 중간 모델. writer(openpyxl/COM)가 이것만 본다.

Cell.fmt ∈ {"title","section","hdr","unit","num","text","bold","italic","flag","date","check","input"} 또는 None.
"input" = 사용자 입력 칸(연노랑 FFFFCC 채움 + 사방 점선 테두리, 글꼴·숫자 서식은 보통 셀과 같음).
Cell.border ∈ {None,"top","bottom","both"}: 가는 실선. "hdr" 셀은 border 가 None 이면 위·아래 둘 다(한 줄 머리글).
값이 normalize.Percent 면 writer 가 0.0% 서식을 입힌다.
"""
from dataclasses import dataclass, field


@dataclass
class Cell:
    v: object
    fmt: object = None
    formula: bool = False
    border: object = None


def merge_border(a, b):
    """두 테두리 지정을 합친다(top+bottom → both)."""
    if not a or a == b:
        return b or a
    if not b:
        return a
    return "both"


@dataclass
class Sheet:
    name: str
    cells: dict = field(default_factory=dict)        # (row, col) -> Cell, 1부터
    col_widths: dict = field(default_factory=dict)   # col -> 폭
    freeze: object = None                            # "A2" 등
    names: dict = field(default_factory=dict)        # 정의된 이름 -> "C12"
    is_data: bool = True                             # [업데이트]가 다시 쓰는 시트인가
    validations: list = field(default_factory=list)  # [(셀 주소 "C10", 목록 formula1 '"1,1000"')] 데이터 유효성(목록)

    def set(self, r, c, v, fmt=None, formula=False, border=None):
        self.cells[(r, c)] = Cell(v, fmt, formula, border)

    def rule(self, r, c1, c2, side):
        """r행 c1..c2 열에 가는 실선(side: top/bottom/both). 빈 칸이면 값 없는 셀을 만든다. 값을 쓴 뒤 부른다."""
        for c in range(c1, c2 + 1):
            cell = self.cells.get((r, c))
            if cell is None:
                self.cells[(r, c)] = Cell(None, None, False, side)
            else:
                cell.border = merge_border(cell.border, side)

    def max_row(self):
        return max((r for r, _ in self.cells), default=0)

    def max_col(self):
        return max((c for _, c in self.cells), default=0)

    def values_matrix(self):
        R, Cc = self.max_row(), self.max_col()
        m = [[None] * Cc for _ in range(R)]
        for (r, c), cell in self.cells.items():
            m[r - 1][c - 1] = cell.v
        return m


@dataclass
class WorkbookModel:
    sheets: list = field(default_factory=list)
    shift_warnings: list = field(default_factory=list)   # 업데이트 때 블록 밀림 경고(설정 상태 문구용)
    unit_value: int = 1                                  # mFS 표시 단위(cfg_unit 에 쓸 나눗수)
    unit_auto: bool = False                              # 자동 판정이었나
    unit_text: str = ""                                  # 설정 '단위 판정' 문구

    def get(self, name):
        return next(s for s in self.sheets if s.name == name)

    def add(self, sheet):
        self.sheets.append(sheet)
        return sheet
