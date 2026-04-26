"""
Dán TSV bảng quyết định (matrix 4+ cột mã) → .xlsx định dạng dự án:
Tahoma 8, nền xanh đậm cột mục + hàng mã, merge A/B/C theo cấu trúc, dropdown O ở vùng tick.
"""
from __future__ import annotations

import io
import re
from typing import List, Optional, Tuple

from excel_testplan_to_table import parse_tsv

# Nền xanh (dark blue 26 / #002060) — aRGB
FILL_DARK = "002060"
FONT_MAIN = "Tahoma"
FC_WHITE = "FFFFFFFF"
FC_BLACK = "FF000000"

_CASE_HEADER = re.compile(r"^(?:AUTH|ADM|STU|TCH)-\d{3}\s*$", re.IGNORECASE)
_MAIN_A = re.compile(r"^(?:Condition|Confirm|Result)$", re.IGNORECASE)
_EXACT_B = re.compile(
    r"^(?:Precondition|Return|Exception|Log message)$",
    re.IGNORECASE,
)

_DATE_PAT = re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b")


def _last_date_in_cell(s: str) -> str:
    """Executed Date: chỉ lấy ngày cuối (sau mũi tên hoặc bản ghi date cuối cùng)."""
    t = (s or "").replace("\n", " ").replace("\r", " ").strip()
    if not t:
        return ""
    for sep in ("→", "->", "=>"):
        if sep in t:
            t = t.split(sep)[-1].strip()
    dates = _DATE_PAT.findall(t)
    if dates:
        return dates[-1]
    # sau dấu ; lấy đoạn có date, hoặc cả đoạn cuối
    parts = re.split(r"[;]", t)
    for p in reversed(parts):
        p = p.strip()
        m = _DATE_PAT.search(p)
        if m:
            return m.group(0)
    return t


def _slash_right_label(s: str) -> str:
    """Failed/Passed -> Passed (lấy phần sau dấu / nếu đúng dạng A/B)."""
    t = (s or "").replace("\n", " ").replace("\r", " ").strip()
    if not t or "/" not in t:
        return t
    if t.startswith("http"):
        return t
    if (t.count("/") == 1) and "http" not in t:
        a, b = t.split("/", 1)
        if 0 < len(a) < 40 and 0 < len(b) < 40:
            return b.strip() or t
    return t


def _is_executed_date_row_b(b: str) -> bool:
    b = (b or "").strip()
    if not b:
        return False
    if b == "Executed Date" or b.startswith("Executed Date"):
        return True
    return b.startswith("Executed") and "Date" in b


def _is_passed_failed_row_b(b: str) -> bool:
    b = (b or "").strip().lower().replace(" ", "")
    if not b:
        return False
    return b == "passed/failed" or b.startswith("passed/failed")


def _is_type_nab_block_label(s: str) -> bool:
    t = (s or "").replace("\n", " ").replace("\r", " ").strip()
    if not t:
        return False
    if t.startswith("Type (") and "Normal" in t and "Abnormal" in t:
        return True
    return "N:" in t and "Normal" in t and "Abnormal" in t and "Boundary" in t


def _clear_bcd_type_from_result_row_if_same_line(grid: List[List[str]], r_res: int) -> None:
    """Bỏ chữ dài 'Type (N: Normal, …' ở B–D trên cùng hàng A=Result (không ngang hàng với 'Result')."""
    if r_res < 0 or r_res >= len(grid):
        return
    if _row_i_col0_strips(grid, r_res) != "Result":
        return
    row = grid[r_res]
    for ci in (1, 2, 3):
        if len(row) > ci and _is_type_nab_block_label(str(row[ci] or "")):
            row[ci] = ""


def _find_r_type_row(grid: List[List[str]], nrows: int) -> int:
    r1 = next(
        (
            i
            for i in range(nrows)
            if "Type" in _row_i_col1_strips(grid, i) and "Boundary" in (grid[i][1] or "")
        ),
        -1,
    )
    if r1 >= 0:
        return r1
    return next(
        (i for i in range(nrows) if _row_i_col1_strips(grid, i).startswith("Type (")),
        -1,
    )


def _join_b_c_d_into_b(grid: List[List[str]], nrows: int, h_row: Optional[int]) -> None:
    """Trước khi merge B:D: gom nội dung B,C,D vào ô B (tránh mất chữ khi chỉ C/D có mô tả). Bỏ qua hàng mã TCH (sẽ gộp A1:D1)."""
    for ri in range(nrows):
        if h_row is not None and ri == h_row:
            continue
        if ri >= len(grid):
            continue
        r = grid[ri]
        while len(r) < 4:
            r.append("")
        b = (r[1] or "").strip()
        c = (r[2] or "").strip()
        d = (r[3] or "").strip()
        if not b and not c and not d:
            continue
        parts = [p for p in (b, c, d) if p]
        if len(parts) == 1 and parts[0] == b and b:
            continue
        r[1] = " ".join(parts)
        r[2] = r[3] = ""


def _preprocess_data_columns(
    grid: List[List[str]],
    nrows: int,
    c0: int,
    c1: int,
    r_res: int,
) -> None:
    """Cột mã: ngày cuối, Passed/Failed lấy phần sau /, xóa 'Result' thừa ở B–D hàng bắt đầu khối Result."""
    use_cases = 0 <= c0 <= c1
    for ri in range(nrows):
        b = _row_i_col1_strips(grid, ri)
        if not b:
            continue
        row = grid[ri] if ri < len(grid) else []
        w = len(row)
        if use_cases and _is_executed_date_row_b(b):
            for ci in range(c0, c1 + 1):
                if ci < w and (row[ci] or "").strip():
                    row[ci] = _last_date_in_cell(str(row[ci] or ""))
        if use_cases and _is_passed_failed_row_b(b):
            for ci in range(c0, c1 + 1):
                if ci < w and (row[ci] or "").strip():
                    row[ci] = _slash_right_label(str(row[ci] or ""))
    if 0 <= r_res < nrows and r_res < len(grid):
        row = grid[r_res]
        for ci in (1, 2, 3):
            if len(row) > ci and (row[ci] or "").strip() == "Result":
                row[ci] = ""
    _clear_bcd_type_from_result_row_if_same_line(grid, r_res)


def _pad_grid(grid: List[List[str]]) -> Tuple[List[List[str]], int, int]:
    if not grid:
        return [[]], 0, 0
    w = max(len(r) for r in grid)
    out = []
    for r in grid:
        row = [r[i] if i < len(r) else "" for i in range(w)]
        out.append(row)
    return out, len(out), w


def _find_code_header_row(grid: List[List[str]]) -> Tuple[Optional[int], int, int]:
    for ri, r in enumerate(grid):
        for ci, c in enumerate(r):
            if c and _CASE_HEADER.match(c.strip()):
                c0 = ci
                c1 = ci
                for cj in range(ci + 1, len(r)):
                    if r[cj] and _CASE_HEADER.match(r[cj].strip()):
                        c1 = cj
                    else:
                        break
                return ri, c0, c1
    return None, 0, 0


def _row_i_col0_strips(grid: List[List[str]], i: int) -> str:
    r = grid[i]
    return (r[0] or "").strip() if r else ""


def _row_i_col1_strips(grid: List[List[str]], i: int) -> str:
    r = grid[i]
    if not r or len(r) < 2:
        return ""
    return (r[1] or "").strip()


def _row_all_empty(r: List[str]) -> bool:
    return not r or not any((c or "").strip() for c in r)


def _collapse_blank_line_between_result_and_type(grid: List[List[str]], r_res: int, r_type: int) -> Tuple[int, int]:
    """Bỏ hàng hoàn toàn rỗng ngay dưới hàng A=Result, trước hàng Type (tránh dòng B/C/D trống thừa). Trả (r_type mới, nrows mới) sau 1 bước."""
    nrows = len(grid)
    if r_res < 0 or r_type < 0 or r_type != r_res + 2:
        return r_type, nrows
    if r_res + 1 >= nrows:
        return r_type, nrows
    if not _row_all_empty(grid[r_res + 1]):
        return r_type, nrows
    del grid[r_res + 1]
    nrows2 = nrows - 1
    return r_type - 1, nrows2


def grid_to_workbook_bytes(grid: List[List[str]]) -> bytes:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("Cần cài: pip install openpyxl") from e

    grid, nrows, ncols = _pad_grid(grid)
    if nrows < 1 or ncols < 1:
        raise ValueError("Bảng trống.")

    h_row, c0, c1 = _find_code_header_row(grid)
    r_cond = next((i for i in range(nrows) if _row_i_col0_strips(grid, i) == "Condition"), -1)
    r_conf = next((i for i in range(nrows) if _row_i_col0_strips(grid, i) == "Confirm"), -1)
    r_res = next((i for i in range(nrows) if _row_i_col0_strips(grid, i) == "Result"), -1)
    if r_res < 0:
        r_res = next((i for i in range(nrows) if _row_i_col1_strips(grid, i) == "Result"), -1)
    r_type = _find_r_type_row(grid, nrows)
    r_type, nrows = _collapse_blank_line_between_result_and_type(grid, r_res, r_type)
    grid, nrows, ncols = _pad_grid(grid)
    if h_row is not None and 0 <= c0 <= c1:
        _preprocess_data_columns(grid, nrows, c0, c1, r_res)
    else:
        _preprocess_data_columns(grid, nrows, -1, -1, r_res)
    # Sau khi xóa "Type" trùng ở cùng hàng Result, tìm lại dòng Type cho merge B–D
    r_type = _find_r_type_row(grid, nrows)
    _join_b_c_d_into_b(grid, nrows, h_row)

    wb = Workbook()
    ws = wb.active
    if ws is None:  # pragma: no cover
        raise RuntimeError("no active sheet")
    ws.title = "Matrix"

    for ri, r in enumerate(grid):
        for ci, v in enumerate(r):
            c = ws.cell(row=ri + 1, column=ci + 1, value=v if v else None)
            c.font = Font(name=FONT_MAIN, size=8, color=FC_BLACK, bold=False)
            c.alignment = Alignment(vertical="center", wrap_text=True)

    thin = Side(style="thin", color="000000")
    b_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            ws.cell(r, c).border = b_all

    def merge_1a(r1: int, c1: int, r2: int, c2: int) -> None:
        if r1 > r2 or c1 > c2:
            return
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    # Cột A: Condition
    if r_cond >= 0 and r_conf > r_cond:
        merge_1a(r_cond + 1, 1, r_conf, 1)
    # Cột A: Confirm → hết dòng trước hàng bắt đầu khối Result (A = Result)
    if r_conf >= 0 and r_res > r_conf:
        merge_1a(r_conf + 1, 1, r_res, 1)
    # Cột A: Result (A=Result) + hàng 1: B..D (trước Type)
    if r_res >= 0:
        a_val = _row_i_col0_strips(grid, r_res) or "Result"
        ws.cell(row=r_res + 1, column=1, value=a_val)
        if r_type >= 0 and r_type > r_res and r_type + 3 < nrows:
            end_1b = min(r_type + 4, nrows)
            merge_1a(r_res + 1, 1, end_1b, 1)
        else:
            merge_1a(r_res + 1, 1, min(r_res + 5, nrows), 1)
    # B:D mọi hàng dữ liệu (nội dung đã gom vào B; trừ hàng mã TCH ở bước gộp A1:D1)
    if ncols >= 4:
        for ri in range(nrows):
            if h_row is not None and ri == h_row:
                continue
            merge_1a(ri + 1, 2, ri + 1, 4)
    # A1:D1: góc + nền dark blue 26 (cùng hàng mã TCH-xxx)
    if h_row is not None and nrows > h_row and ncols >= 4:
        merge_1a(h_row + 1, 1, h_row + 1, 4)

    if h_row is not None and c0 <= c1 and c1 >= 0:
        c_start = c0 + 1
        c_end = c1 + 1
        if r_type > h_row + 1:
            last_0 = r_type - 1
        elif r_type >= 0:
            last_0 = max(h_row, r_type - 1)
        else:
            last_0 = nrows - 5 if nrows > h_row + 5 else nrows - 1
        if last_0 > h_row:
            a1 = f"{get_column_letter(c_start)}{h_row + 2}"
            a2 = f"{get_column_letter(c_end)}{last_0 + 1}"
            dv = DataValidation(
                type="list",
                formula1='"O"',
                allow_blank=True,
                showErrorMessage=True,
            )
            ws.add_data_validation(dv)
            dv.add(f"{a1}:{a2}")

    fill = PatternFill(start_color=FILL_DARK, end_color=FILL_DARK, fill_type="solid")
    font_w = Font(name=FONT_MAIN, size=8, bold=True, color=FC_WHITE)
    a_align_main = Alignment(horizontal="center", vertical="top", wrap_text=True)
    for ri in range(nrows):
        a0 = _row_i_col0_strips(grid, ri)
        if a0 in ("Condition", "Confirm", "Result") or (a0 and _MAIN_A.match(a0)):
            cell = ws.cell(row=ri + 1, column=1)
            if not cell.value and a0:
                cell.value = a0
            cell.font = font_w
            cell.fill = fill
            cell.alignment = a_align_main
    for ri in range(nrows):
        b = _row_i_col1_strips(grid, ri)
        if b in ("Precondition", "Return", "Exception", "Log message") or (b and _EXACT_B.match(b)):
            c = ws.cell(row=ri + 1, column=2)
            c.font = Font(name=FONT_MAIN, size=8, bold=True, color=FC_BLACK)
    if h_row is not None and 0 <= c0 <= c1:
        corner = ws.cell(row=h_row + 1, column=1)
        corner.fill = fill
        corner.font = font_w
        corner.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=h_row + 1, column=c + 1)
            cell.font = Font(name=FONT_MAIN, size=8, bold=True, color=FC_WHITE)
            cell.fill = PatternFill(start_color=FILL_DARK, end_color=FILL_DARK, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if r_res >= 0:
        cr = ws.cell(row=r_res + 1, column=1)
        if not (cr.value and str(cr.value).strip()):
            cr.value = "Result"
        cr.font = font_w
        cr.fill = fill
        cr.alignment = a_align_main

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def tsv_to_xlsx_bytes(text: str) -> bytes:
    t = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not t.strip():
        raise ValueError("Dữ liệu trống.")
    g = parse_tsv(t)
    if not g:
        raise ValueError("Không parse được dòng nào.")
    return grid_to_workbook_bytes(g)
