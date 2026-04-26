"""
Từ TSV bảng test: nhóm theo từng mục (Login, User Management, …) và
sinh Markdown: prompt cho AI xuất **bảng quyết định (decision / matrix)** — TSV
4 cột A–D + cột mã (E…), dán Excel.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional

from excel_testplan_to_table import parse_tsv

# Cho phép AUTH|ADM|STU|TCH
CASE_ID = re.compile(r"^(?:AUTH|ADM|STU|TCH)-\d{3}\s*$", re.IGNORECASE)


@dataclass
class TestRow:
    id: str
    title: str
    steps: str
    expected: str
    pre: str
    r1s: str
    r1d: str
    r1t: str
    r2s: str
    r2d: str
    r2t: str
    r3s: str
    r3d: str
    r3t: str
    note: str


@dataclass
class Section:
    name: str
    tests: List[TestRow] = field(default_factory=list)


def _is_header_row(row: List[str]) -> bool:
    if not row:
        return False
    a = (row[0] or "").strip().lower()
    return a.startswith("test case") or a == "id"


def _row_to_test(row: List[str], n: int) -> Optional[TestRow]:
    r = row + [""] * (15 - len(row)) if len(row) < 15 else row[:15]
    tid = (r[0] or "").strip()
    if not CASE_ID.match(tid):
        return None
    return TestRow(
        id=tid,
        title=r[1] or "",
        steps=r[2] or "",
        expected=r[3] or "",
        pre=r[4] or "",
        r1s=r[5] or "",
        r1d=r[6] or "",
        r1t=r[7] or "",
        r2s=r[8] or "",
        r2d=r[9] or "",
        r2t=r[10] or "",
        r3s=r[11] or "",
        r3d=r[12] or "",
        r3t=r[13] or "",
        note=r[14] or "",
    )


def _is_skip_section_row(first: str) -> bool:
    t = (first or "").strip()
    if not t:
        return True
    if re.match(r"(?i)^function\s+.+$", t):
        return True
    if t.lower() in ("test cases", "test case"):
        return True
    return False


def group_sections(rows: List[List[str]]) -> List[Section]:
    order: list[str] = []
    buckets: dict[str, List[TestRow]] = {}
    cur = "Mục chưa gán"

    for raw in rows:
        if not raw or not any((c or "").strip() for c in raw):
            continue
        if _is_header_row(raw):
            continue
        first = (raw[0] or "").strip()
        t = _row_to_test(raw, 15)
        if t:
            if cur not in buckets:
                buckets[cur] = []
                if cur not in order:
                    order.append(cur)
            buckets[cur].append(t)
        else:
            if _is_skip_section_row(first):
                continue
            cur = first
            if cur not in buckets:
                order.append(cur)
                buckets[cur] = []

    return [Section(n, buckets[n]) for n in order if len(buckets.get(n, [])) > 0]


def _tsv_table_line_test_data(t: TestRow) -> str:
    return (
        f"- **{t.id}** | *{t.title}* | "
        f"Pre-conditions: {t.pre or '(trống)'} | Kết quả mong đợi: {t.expected or '(trống)'} | "
        f"R1: {t.r1s} {t.r1d} {t.r1t} | R2: {t.r2s} {t.r2d} {t.r2t} | R3: {t.r3s} {t.r3d} {t.r3t} | Ghi chú: {t.note or '(trống)'}"
    )


def _row_excel_4col_matrix(a: str, b: str, c: str, d: str, matrix: list[str]) -> str:
    """Một dòng dán Excel: 4 cột A–D (nhãn) + cột mã E… (len(matrix) cột, thường bằng số mã thử)."""
    return "\t".join([a, b, c, d, *matrix])


def _skeleton_tsv_for_ids(ids: List[str], _section: str) -> str:
    """Khung bảng TSV: 4 cột A–D trước cột mã; AI điền O ở E…"""
    if not ids:
        return "\t( không có mã thử )"
    n = len(ids)
    z = [""] * n
    # Dòng 1: 4 Ô trống, rồi tiêu đề AUTH-… (E·…)
    line0 = _row_excel_4col_matrix("", "", "", "", list(ids))
    n_ids = ", ".join(ids)
    return "\n".join(
        [
            line0,
            _row_excel_4col_matrix("Condition", "Precondition", "", "", z),
            _row_excel_4col_matrix(
                "", "", "", f"(điền từng dòng pre-condition, sau đó {n} cột {n_ids} — dấu O, lấy từ cột pre-conditions ở dữ liệu đã nêu)",
                z,
            ),
            _row_excel_4col_matrix("Confirm", "Return", "", "", z),
            _row_excel_4col_matrix(
                "",
                "",
                "",
                f"(điền từng dòng kết quả/return, {n} cột O — từ cột «Kết quả mong đợi / Steps» ở từng mã; AI có thể gộp nếu cùng câu chữ).",
                z,
            ),
            _row_excel_4col_matrix("", "Exception", "", "", z),
            _row_excel_4col_matrix(
                "", "", "", f"(AI: thêm mới các dòng exception phù hợp, O đúng cột {n_ids}.)", z
            ),
            _row_excel_4col_matrix("", "Log message", "", "", z),
            _row_excel_4col_matrix(
                "", "", "", "(AI: thêm mới; có thể lấy ý từ cột ghi chú ở TSV; O đúng cột).", z
            ),
            _row_excel_4col_matrix("", "Result", "", "", z),
            _row_excel_4col_matrix(
                "",
                "Type (N: Normal, A: Abnormal, B: Boundary)",
                "",
                "",
                z,
            ),
            _row_excel_4col_matrix(
                "",
                "Passed/Failed",
                "",
                "",
                z,
            ),
            _row_excel_4col_matrix(
                "",
                "Executed Date",
                "",
                "",
                ["(điền)"] * n,
            ),
            _row_excel_4col_matrix("", "Defect ID (để trống)", "", "", z),
        ]
    )


def _login_sixcol_example(ids: List[str]) -> str:
    """
    Mẫu đúng cấu trúc Excel: 4 cột A–D (nhãn / mô tả dòng) rồi 6 cột mã từ E.
    Cột E = cột mã 1, …, J = cột mã 6. Chỉ dùng nếu len(ids)==6 và section Login.
    """
    if len(ids) != 6:
        return ""
    z = [""] * 6
    a, b, c, d, e, f = ids[0], ids[1], ids[2], ids[3], ids[4], ids[5]
    R = _row_excel_4col_matrix
    lines: list[str] = [
        R("", "", "", "", [a, b, c, d, e, f]),
        R("Condition", "Precondition", "", "", z),
        R("", "", "", "User exists in system", ["O", "", "O", "O", "O", ""]),
        R("", "", "", "Email does not exist in system", ["", "O", "", "", "", ""]),
        R("", "", "", "Password is incorrect", ["", "", "O", "", "", ""]),
        R("", "", "", "Account is activated", ["O", "", "", "O", "", "O"]),
        R("", "", "", "Account is not activated", ["", "", "", "O", "", ""]),
        R("", "", "", "First login (no previous login history)", ["", "", "", "", "O", ""]),
        R("", "", "", "User exists with multiple roles available", ["", "", "", "", "", "O"]),
        R("Confirm", "Return", "", "", z),
        R("", "", "", "Login successfully", ["O", "", "", "", "O", "O"]),
        R("", "", "", "Login unsuccessful", ["", "O", "O", "O", "", ""]),
        R("", "", "", "Access token is returned", ["O", "", "", "", "", ""]),
        R("", "", "", "Refresh token cookie is set", ["O", "", "", "", "", ""]),
        R("", "", "", "User information is returned", ["O", "", "", "", "", ""]),
        R("", "", "", "Activation email is sent", ["", "", "", "O", "", ""]),
        R("", "", "", "Welcome email is sent to user", ["", "", "", "", "O", ""]),
        R("", "", "", "Access token with selected role is returned", ["", "", "", "", "", "O"]),
        R("", "", "", "User redirected to role-specific dashboard", ["", "", "", "", "", "O"]),
        R("", "Exception", "", "", z),
        R("", "", "", "UnauthorizedException is thrown", ["", "O", "O", "", "", ""]),
        R("", "", "", "AccountNotActivatedException is thrown", ["", "", "", "O", "", ""]),
        R("", "Log message", "", "", z),
        R("", "", "", "Token expiry configuration issue in Round 1", ["O", "", "", "", "", ""]),
        R("", "", "", "Role selection dropdown not rendering in Round 1", ["", "", "", "", "", "O"]),
        R("", "Result", "", "", z),
        R("", "Type (N: Normal, A: Abnormal, B: Boundary)", "", "", ["N", "A", "A", "A", "N", "A"]),
        R("", "Passed/Failed", "", "", ["P", "P", "P", "P", "P", "P"]),
        R(
            "",
            "Executed Date",
            "",
            "",
            ["(điền)"] * 6,
        ),
        R("", "Defect ID (để trống)", "", "", z),
    ]
    return "\n".join(lines)


def _section_to_markdown(sec: Section) -> str:
    """Một mục → prompt: bảng quyết định (matrix) TSV 4 cột A–D + cột mã, dán Excel."""
    if not sec.tests:
        return ""
    ids = [t.id for t in sec.tests]
    n = len(ids)
    id_line = " ".join(f"`{x}`" for x in ids)
    bullet = "\n".join(_tsv_table_line_test_data(t) for t in sec.tests)
    is_login_six = sec.name.strip().lower() == "login" and n == 6
    if is_login_six:
        sk_main = _login_sixcol_example(ids)
        sk_alt = _skeleton_tsv_for_ids(ids, sec.name)
        skeleton = sk_main + "\n\n*— Bản rút gọn tùy biến số cột (nếu cần):*\n\n" + sk_alt
    else:
        skeleton = _skeleton_tsv_for_ids(ids, sec.name)
    code_fence = "```"
    return f"""### {sec.name}

**Đầu ra bắt buộc:** Một **khối TSV** (cột cách bằng **tab**) — **Copy → dán Excel** («Text tách cột bằng tab») — dạng **bảng quyết định (test design matrix)**: **A** = nhãn cột dọc lớn (`Condition`, `Confirm`, `Result`); **B** = tiêu đề từng khối: `Precondition` / `Return` / `Exception` / `Log message` (các từ này **chỉ ở cột B**, giống mẫu, không thay bằng cột A). Cột D = mô tả dòng, cột E trở đi = **{n} cột mã** ({id_line}) và dấu **O**. **Tối thiểu** phải có **bảng TSV hoàn chỉnh** này. (Tùy chọn) kèm VBA/Office Script hoặc ghi **ô bắt đầu** nếu khác `A1`. **Không** trả bảng 15 cột «một dòng — một case» (kiểu Admin) — **output prompt này là matrix**.

Bạn tạo **bảng quyết định (decision table)**: từ cột E, **mỗi cột = một mã thử** ({id_line}).

**Nguồn đáng tin (không tự sửa nội dung cột nghiệp vụ, chỉ sắp lên lưới + đánh dấu O nơi trùng):**
- Từ dữ liệu sau: cột *Pre-conditions* và *Kết quả / Steps* ở từng mã.

**Bạn tự tạo thêm (AI):**
- Các dòng *Exception* và *Log message* phù hợp, đánh dấu **O** ở cột mã tương ứng.
- Có thể bổ sung dòng mô tả nếu cần, miễn hợp nghiệp vụ từng mã.

**Lấy từ dữ liệu đã cho (sẵn, không tự tạo dữ liệu cột sản phẩm nếu TSV đã có):**
- Các cột mã: **R1 / R2 / R3** (Trạng thái, Ngày, Tester) — dùng cho hàng *Result* / *Passed* / *Executed date* theo từng cột mã.
- Cột *Defect ID* — **để trống** ở mọi ô cột mã (hàng này nếu có).

Dữ liệu gốc (theo từng mã):

{bullet}

**Quy tắc cột (bắt buộc — dán mới lệch nếu sai):**
- **A / B / D:** Cột A nhãn dọc lớn: `Condition`, `Confirm` (cùng hàng mở đầu với `Precondition` / `Return` ở B). Mô tả từng dòng ở **D**; **E…** = {n} cột mã. Trong vùng Confirm: `Return` / `Exception` / `Log message` ở **B**; A trống (merge A theo mục sau khi dán).
- **Tiêu đề mục Result (1 dòng):** chữ `Result` ghi ở **B**, **C trống** (TSV) → dán xong **merge B:C** cho ô này (một ô, như Ảnh 2). **A** trống ở dòng đó, hoặc gộp A thành cột mục lớn (xem dưới).
- **Bốn dòng** *Type* / *Passed/Failed* / *Executed Date* / *Defect ID:* nhãn ghi ở **B**, **C–D trống** (TSV) → **merge B:D** cho từng dòng. Dữ liệu từng cột mã ở **E…**; không nhét mô tả dài ở cột D cùng hàng.
- Dòng **tiêu đề mã thử** (E…): A–D trước hàng mã thử trống, rồi từ E là từng mã. **Dùng tab thật.**

**Định dạng chữ & màu (Excel — bắt buộc nếu xuất file mẫu / VBA / Office Script):**
- **Font:** mặc định **Tahoma**; cỡ chữ cơ sở **8 pt** (mọi ô).
- **Màu nền (dark blue, chuẩn dự án 26 / xanh đậm):** dùng cho **cột mục dọc** (ô chứa `Condition`, `Confirm`, `Result`) **và** hàng **tiêu đề cột mã** (`AUTH-…`, `TCH-…`, `ADM-…`, `STU-…`); nền = **xanh đậm** (thực tế: `RGB(0, 32, 96)` tương đương #002060 / *Dark Blue*; đồng bộ theo bảng màu tổ chức nếu mã 26 tương ứng màu xanh công ty), chữ **trắng**:
  - Nhãn cột mục **A:** `Condition`, `Confirm`, `Result` — **Tahoma Bold 8, trắng** trên nền xanh đậm.
  - **Hàng mã thử (E1…):** cùng nền xanh đậm, **Tahoma Bold 8, trắng** (mỗi cột = một mã, vd. `AUTH-001` …).
- **Tiêu đề cấp 2 ở cột B** (các ô: `Precondition`, `Return`, `Exception`, `Log message`, cùng hàng tương ứng): **Tahoma Bold 8, đen**, nền thường (trắng / nền lưới, không tô cùng nền xanh cột mục).
- **Mọi chữ còn lại** (mô tả ở D, bảng O/P/ngày, v.v.): **Tahoma 8, đen**, không in đậm (trừ khi bạn đã liệt kê ở trên).
- **Kẻ tất cả viền** cho vùng bảng. Khối **Result:** merge A dọc + B:C + B:D như quy tắc cột ở trên.
- **Ô tick (giao tuyến mã, chữ O):** không gõ O tay; dùng **Data validation (dropdown)** với nguồn danh sách **chỉ gồm chữ O** (và cho phép để trống nếu cần), giống mẫu. **Excel:** *Data* → *Data Validation* → *List* → nguồn `O` hoặc công thức `= "O"`. **Google Sheets:** *Data* → *Data validation* → *List of items* → `O`. Sau đó từng ô hiển thị mũi tên, chọn **O** khi cần đánh dấu.

{code_fence}vb
' Tùy chọn. Chỉnh h0, lc theo số dòng/cột. Nền xanh: RGB(0, 32, 96) = dark blue / chuẩn 26 / #002060
Sub ApplyMatrixMergesAndStyle()
    Dim DarkBlue26 As Long, h0 As Long, i As Long, lc As Long
    DarkBlue26 = RGB(0, 32, 96)
    h0 = 26
    On Error Resume Next
    With ActiveSheet.UsedRange
        .Font.Name = "Tahoma"
        .Font.Size = 8
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    Range("B" & h0 & ":C" & h0).Merge
    For i = 1 To 4: Range("B" & (h0 + i) & ":D" & (h0 + i)).Merge: Next i
    With Range("A" & h0 & ":A" & (h0 + 4))
        .Merge
        .Value = "Result"
        .Font.Name = "Tahoma"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Interior.Color = DarkBlue26
    End With
    lc = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    If lc >= 5 Then
        With ActiveSheet.Range(ActiveSheet.Cells(1, 5), ActiveSheet.Cells(1, lc))
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = DarkBlue26
        End With
    End If
    ' Thủ công: format ô Condition, Confirm, các khối cột A; Precondition, Return, Exception, Log message: Bold 8 đen
    ActiveSheet.UsedRange.Borders.LineStyle = xlContinuous
End Sub

' Dropdown ô tick O — CHỈNH r1,r2,c1,c2 = vùng matrix (bỏ hàng tiêu đề, bỏ khối Result/P-F-Date)
' Ví dụ Login 6 mã: hàng 2–25, cột E–J nếu Result từ hàng 26
Sub ApplyOTickDropdown()
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    r1 = 2: r2 = 25: c1 = 5: c2 = 10
    On Error Resume Next
    With ActiveSheet.Range(ActiveSheet.Cells(r1, c1), ActiveSheet.Cells(r2, c2))
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="O"
    End With
End Sub
{code_fence}
*(Nếu `Formula1:="O"` báo lỗi locale, dùng một ô tham chiếu, vd. $Z$1=O, hoặc `=CHAR(79)` tùy phiên bản. **Google Sheets:** vùng → Data validation → List → `O`.)*  
*(ColorIndex 26: kiểm tra theme. Office Script: tương tự Tahoma, màu, + validation `O` ở ô tick.)*

**Khung bảng TSV mẫu** (dán «Text tách cột bằng tab») — {n} cột mã ({id_line}):

{code_fence}text
{skeleton}
{code_fence}

**Gợi ý thứ tự trả lời:** (1) TSV (4 + {n} cột) để dán, (2) merge/màu/VBA, (3) **dropdown O** ở vùng tick (`ApplyOTickDropdown` hoặc Data validation thủ công), (4) tùy chọn: ô bắt đầu nếu khác `A1`. Không lệch merge B:C / B:D ở khối Result.
""".rstrip()


def build_ai_prompt_markdown(tsv_text: str) -> str:
    """Trả 1 chuỗi Markdown: nhiều block ### mục, nối bằng ---; mỗi mục = prompt matrix TSV."""
    text = tsv_text.replace("\r\n", "\n").replace("\r", "\n")
    rows = parse_tsv(text)
    sections = group_sections(rows)
    if not sections:
        return (
            "Không tách được mục/mã nào. Kiểm tra TSV: cột 1 = mã dạng AUTH/TCH/ADM/STU-xxx, "
            "các dòng tên mục ở cột 1, copy đủ 15 cột từ Excel."
        )

    parts: list[str] = []
    for sec in sections:
        b = _section_to_markdown(sec)
        if b:
            parts.append(b)

    if not parts:
        return (
            "Không tách được mục/mã nào. Kiểm tra TSV: cột 1 = mã dạng AUTH/TCH/ADM/STU-xxx, "
            "các dòng tên mục ở cột 1, copy đủ 15 cột từ Excel."
        )
    return "\n\n---\n\n".join(parts)


if __name__ == "__main__":
    import sys

    s = sys.stdin.read()
    sys.stdout.write(build_ai_prompt_markdown(s))
