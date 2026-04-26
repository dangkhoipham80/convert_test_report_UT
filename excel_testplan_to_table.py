#!/usr/bin/env python3
"""
Bước 1: Dán từ Excel (cột A–O) → bảng text (Markdown) dùng được cho Notion, GitHub, v.v.

- Ưu tiên: dữ liệu dạng TSV (Copy từ Excel, mỗi cột 1 tab).
- Hỗ trợ: ô nhiều dòng nếu Excel bọc trong dấu ngoặc kép.
- Tùy chọn: chuỗi "dính" (không tab) — tách theo mã (AUTH|ADM|STU|TCH)-###, xuất bảng 2 cột.
- `--format flat-tsv`: một bảng TSV, mỗi ô bỏ hết xuống dòng nội bộ (giống bản bạn dán vào Notion / text).
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
from typing import List, Sequence

# Nhận dạng dòng test case theo 4 loại mã
ID_RE = re.compile(r"^(?:AUTH|ADM|STU|TCH)-\d{3}\s*$", re.IGNORECASE)
CASE_ID_IN_TEXT = re.compile(r"\b((?:AUTH|ADM|STU|TCH)-\d{3})\b", re.IGNORECASE)

# 15 cột A–O
COLUMNS_15: List[str] = [
    "ID",
    "Tên case",
    "Steps",
    "Kết quả mong đợi",
    "Pre-conditions",
    "R1 Trạng thái",
    "R1 Ngày",
    "R1 Tester",
    "R2 Trạng thái",
    "R2 Ngày",
    "R2 Tester",
    "R3 Trạng thái",
    "R3 Ngày",
    "R3 Tester",
    "Ghi chú",
]


def _normalize_text(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\r", "\n").strip()


def _row_to_md(
    cells: Sequence[str],
    mode: str,
) -> str:
    """Xuất một dòng bảng Markdown; mode br dùng <br> cho xuống dòng (Notion)."""
    out: list[str] = []
    for c in cells:
        s = c or ""
        s = s.replace("\r\n", "\n").replace("\r", "\n")
        if mode == "br":
            s = s.strip().replace("\n", "<br>")
        elif mode == "keep":
            s = s.rstrip()
        elif mode == "slash":
            s = re.sub(r"\n+", " / ", s.strip())
        else:  # space
            s = re.sub(r"\n+", " ", s.strip())
        s = s.replace("\\", "\\\\").replace("|", "\\|")
        out.append(s)
    return "| " + " | ".join(out) + " |"


def parse_tsv(text: str) -> List[List[str]]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    f = io.StringIO(text)
    reader = csv.reader(f, dialect="excel-tab")
    return [row for row in reader]


def flatten_cell(s: str | None) -> str:
    """
    Bỏ mọi ký tự xuống dòng trong một ô; nối liền các phần (không thêm khoảng trắng),
    giống output sau khi copy từ Excel dạng một dòng/ô.
    """
    if not s:
        return ""
    r = s.replace("\r\n", "\n").replace("\r", "\n")
    return "".join(r.splitlines())


def _pad_row(row: List[str], width: int) -> List[str]:
    if len(row) < width:
        return row + [""] * (width - len(row))
    return row[:width]


def _is_likely_tsv(text: str) -> bool:
    if "\t" in text:
        return True
    return False


def _split_jammed(text: str) -> List[tuple[str, str]]:
    t = _normalize_text(text)
    m_iter = list(CASE_ID_IN_TEXT.finditer(t))
    if not m_iter:
        return []
    out: list[tuple[str, str]] = []
    for i, m in enumerate(m_iter):
        cid = m.group(1)
        start = m.end()
        end = m_iter[i + 1].start() if i + 1 < len(m_iter) else len(t)
        rest = t[start:end].strip()
        out.append((cid, rest))
    return out


def rows_to_markdown(
    rows: List[List[str]],
    *,
    mode: str,
    ncols: int,
    include_header: bool,
    skip_non_id: bool,
) -> str:
    if ncols == 15:
        col_names = COLUMNS_15
    else:
        col_names = COLUMNS_15[:ncols] if ncols < len(COLUMNS_15) else COLUMNS_15 + [
            f"Extra{j+1}" for j in range(len(COLUMNS_15), ncols)
        ]

    md_rows: list[str] = []
    if include_header:
        md_rows.append(
            _row_to_md([str(x) for x in col_names[:ncols]], mode="space"),
        )
        md_rows.append("|" + "|".join([" --- "] * ncols) + "|")

    for raw in rows:
        row = _pad_row([c or "" for c in raw], ncols)
        first = (row[0] or "").strip()
        if skip_non_id and first and not ID_RE.match(first):
            continue
        if not any((x or "").strip() for x in row):
            continue
        md_rows.append(_row_to_md(row, mode=mode))

    return "\n".join(md_rows) + "\n"


def _infer_ncols(rows: List[List[str]], default: int) -> int:
    m = default
    for r in rows:
        m = max(m, len(r))
    return max(1, m)


def rows_to_flat_tsv(
    rows: List[List[str]],
    *,
    ncols: int,
) -> str:
    """
    Tất cả dòng (header, section, AUTH-*) → một bảng TSV; mỗi ô là một dòng text (không \\n bên trong).
    """
    buf = io.StringIO()
    w = csv.writer(
        buf,
        dialect=csv.excel_tab,
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
    )
    for raw in rows:
        flat = [flatten_cell(c) for c in (raw or [])]
        row = _pad_row(flat, ncols)
        if not any(c.strip() for c in row):
            continue
        w.writerow(row)
    return buf.getvalue()


def jammed_to_md(pairs: List[tuple[str, str]], mode: str) -> str:
    lines = [
        _row_to_md(["ID", "Nội dung (cần copy lại TSV từ Excel nếu cần 15 cột)"], mode="space"),
        "|" + "|".join([" --- ", " --- "]) + "|",
    ]
    for cid, rest in pairs:
        lines.append(_row_to_md([cid, rest], mode=mode))
    return "\n".join(lines) + "\n"


def read_input(infile: str | None) -> str:
    if infile:
        with open(infile, encoding="utf-8", newline="") as f:
            return f.read()
    return sys.stdin.read()


def main() -> int:
    p = argparse.ArgumentParser(
        description="Excel A–O (TSV) → Markdown hoặc TSV mỗi ô một dòng (bước 1).",
    )
    p.add_argument(
        "--format",
        choices=("markdown", "flat-tsv"),
        default="markdown",
        help="markdown=bảng MD; flat-tsv=1 bảng TSV, bỏ xuống dòng trong từng ô (dán Notion).",
    )
    p.add_argument(
        "infile",
        nargs="?",
        help="File UTF-8 (có thể bỏ qua, dùng pipe hoặc dán).",
    )
    p.add_argument(
        "-o",
        "--out",
        help="Ghi file (mặc định: stdout).",
    )
    p.add_argument(
        "--ncols",
        type=int,
        default=0,
        help="Cố định số cột; 0 = tự ước lượng (tối thiểu 15 khi dùng TSV đủ 15 cột).",
    )
    p.add_argument(
        "--no-header",
        action="store_true",
        help="Không in dòng tên cột + separator.",
    )
    p.add_argument(
        "--multiline",
        choices=("br", "keep", "slash", "space"),
        default="br",
        help="Xuống dòng trong ô: br=<br> (Notion), keep, slash, space.",
    )
    p.add_argument(
        "--jammed",
        action="store_true",
        help="Bắt buộc chế độ chuỗi dính (không tab).",
    )
    p.add_argument(
        "--skip-section-rows",
        action="store_true",
        help="Bỏ dòng cột 1 không phải AUTH|ADM|STU|TCH-### (vd. dòng 'Login').",
    )
    p.add_argument(
        "--jammed-out",
        action="store_true",
        help="Khi gặp chuỗi dính: xuất bảng 2 cột thay vì cảnh báo.",
    )
    p.add_argument(
        "--bom",
        action="store_true",
        help="Ghi file output kèm BOM UTF-8 (tiện mở lại bằng Excel).",
    )
    args = p.parse_args()

    raw = read_input(args.infile)
    if not raw.strip():
        print("Không có dữ liệu. Dán TSV từ Excel vào stdin hoặc truyền file.", file=sys.stderr)
        return 1

    tsv = _is_likely_tsv(raw) and not args.jammed
    if not tsv or args.jammed:
        if args.jammed_out or (not tsv and len(_split_jammed(raw)) > 0):
            pairs = _split_jammed(raw)
            if not pairs and not tsv:
                print("Không tìm thấy mã AUTH|ADM|STU|TCH-###. Hãy copy từ Excel dạng nhiều cột (tab).", file=sys.stderr)
                return 1
            if pairs:
                out = jammed_to_md(pairs, args.multiline)
                if args.out:
                    with open(args.out, "w", encoding="utf-8") as f:
                        f.write(out)
                else:
                    sys.stdout.write(out)
                if not tsv and not args.jammed_out:
                    print(
                        "\n(Lưu ý: bản dính 1 dòng — chỉ tách 2 cột. "
                        "Nên copy từ Excel cả vùng A:O để đủ 15 cột TSV.)",
                        file=sys.stderr,
                    )
                return 0
        if not tsv:
            print("Input không có ký tự tab — có phải bạn cần --jammed-out ?", file=sys.stderr)
            return 1

    rows = parse_tsv(raw)
    if args.format == "flat-tsv":
        # Đúng số cột theo dữ liệu (ví dụ 15 cột từ Excel), không thêm cột rỗng tới 15 mặc định
        ncols = args.ncols if args.ncols > 0 else _infer_ncols(rows, 1)
    else:
        ncols = args.ncols if args.ncols > 0 else max(15, _infer_ncols(rows, 15))

    if args.format == "flat-tsv":
        out = rows_to_flat_tsv(rows, ncols=ncols)
    else:
        out = rows_to_markdown(
            rows,
            mode=args.multiline,
            ncols=ncols,
            include_header=not args.no_header,
            skip_non_id=args.skip_section_rows,
        )
    enc = "utf-8-sig" if args.bom and args.out else "utf-8"
    if args.out:
        with open(args.out, "w", encoding=enc, newline="") as f:
            f.write(out)
    else:
        sys.stdout.write(out)
    return 0


if __name__ == "__main__":
    if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except (AttributeError, OSError, ValueError):
            pass
    raise SystemExit(main())
