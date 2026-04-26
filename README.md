# Excel test plan → bảng text (TSV / Markdown)

Dán từ Excel (cột A–O, nhiều dòng trong ô) → **một bảng TSV mỗi ô một dòng** (dán Notion, báo cáo) hoặc **bảng Markdown**.

## Yêu cầu

- **Python 3** có trong `PATH` (lệnh `python` hoặc `py`).
- Trình duyệt (Edge / Chrome) để mở giao diện (localhost).

## Giao diện web (chính)

Trên giao diện **HTML/CSS/JavaScript** (thư mục `web/`), bạn dán dữ liệu **trong trình duyệt** — thân thiện hơn console. Máy phục vụ chạy **chỉ trên máy bạn** (`127.0.0.1`), không tải dữ liệu lên mạng.

1. Từ thư mục này, chạy:
   ```powershell
   python .\serve_ui.py
   ```
2. Trình duyệt tự mở `http://127.0.0.1:8765/` (nếu không, mở link đó bằng tay).
3. Trong Excel: bôi bảng → **Ctrl+C** → trong trang web bấm **Lấy từ clipboard** (hoặc **Ctrl+V** vào ô lớn).
4. Bấm **Chuyển → TSV** hoặc **Chuyển → Markdown** → **Sao chép kết quả**  
5. (Tùy chọn) **Sinh prompt cho AI**: chạy **trong trình duyệt** — mỗi mục một khối copy; prompt yêu cầu AI trả **bảng quyết định (matrix)** dạng TSV: **4 cột A–D** + **cột mã** (E…), dấu **O** ở giao tuyến (mẫu Login 6 mã có sẵn khung). Tùy chọn kèm VBA/Office Script; **không** nhầm với bảng test plan 15 cột «một dòng một case» (kiểu Admin).

**Sau** khi cài lệnh `actr` (xem mục dưới), chỉ cần gõ `actr` — tương đương chạy `TestReportConvert-Menu.ps1` (khởi động `serve_ui.py`).

**Cổng bận?** Một instance khác đang mở, hoặc dùng:
```powershell
$env:TEST_REPORT_UI_PORT = "8766"; python .\serve_ui.py
```

**Chưa tới /api/health (TSV, Markdown thất bại):** cần chạy `python serve_ui.py` và mở đúng cổng. **Sinh prompt cho AI** vẫn dùng được khi tắt server (xử lý trong trình duyệt). Nếu mở `file://` hoặc đổi cổng: thêm `?port=8766` trên URL trỏ tới `index.html` cho khớp.

## Cách vào đúng thư mục (bên nào copy cũng dùng được)

Bạn cần chạy lệnh **trong thư mục** chứa `excel_testplan_to_table.py` (tên thư mục thường là `excel-testplan-to-table`).

| Cách | Thao tác |
|------|----------|
| **1 — Không cần gõ `cd`** | Mở File Explorer, vào thư mục chứa bộ file này. Trên thanh địa chỉ gõ `powershell` rồi Enter → PowerShell mở **sẵn tại thư mục đó**. |
| **2 — Biến môi trường (Windows)** | Mỗi máy một đường dẫn. Thay `PATH_TO` bằng đường dẫn thật tới thư mục (có thể dán từ Explorer: Shift+chọn đường dẫn). |

Dùng biến (tránh dính tên `C:\Users\...` cố định; chỉ cần sửa **một lần** chuỗi cuối cho đúng cây thư mục của bạn):

```powershell
Set-Location -LiteralPath (Join-Path $env:USERPROFILE 'path\to\excel-testplan-to-table')
```

Ví dụ nếu bạn cất project trong `.cursor\projects\empty-window\excel-testplan-to-table` (dưới profile user hiện tại, máy bất kỳ):

```powershell
Set-Location -LiteralPath (Join-Path $env:USERPROFILE '.cursor\projects\empty-window\excel-testplan-to-table')
```

| **3 — Đang ở thư mục cha (repo, Desktop…)** | Vào thư mục con bằng tên, không cần đường dẫn tuyệt đối: |

```powershell
Set-Location .\excel-testplan-to-table
```

**Gợi ý:** Trong VS Code / Cursor, `Terminal → New Terminal` khi thư mục mở đúng workspace thường **đã đúng** `cwd` — khi đó **không cần** `Set-Location` nếu bạn mở terminal từ cạnh cây thư mục file.

## Ghi chú encoding (Windows PowerShell 5.1)

File `TestReportConvert-Menu.ps1` dùng **chữ ASCII** trong mọi chuỗi hiển thị, để khi lưu **UTF-8 (không BOM)** Windows PowerShell vẫn parse đúng. Nếu bạn sửa lại bằng tiếng Việt, hãy lưu file **UTF-8 với BOM** (trong VS Code: chọn encoding ở góc, “Save with Encoding” → **UTF-8 with BOM**), nếu không sẽ lỗi cú pháp giả.

## Cài lệnh nhanh (một lần)

Sau khi `Set-Location` tới thư mục này (xem bảng trên):

```powershell
powershell -ExecutionPolicy Bypass -File .\Register-TestReportCommand.ps1
```

Mở **tab PowerShell mới** hoặc tải lại profile:

```powershell
. $PROFILE
```

Từ đó, ở bất kỳ đâu (sau khi profile đã load), gõ:

- `actr` **hoặc** `auto-convert-test-report` — chạy **giao diện web** (`serve_ui.py`).

Gỡ cài hàm đã thêm (chạy lại từ thư mục này):

```powershell
powershell -ExecutionPolicy Bypass -File .\Register-TestReportCommand.ps1 -Uninstall
```

## Lệnh nhanh `actr` (tùy chọn)

Sau khi chạy `Register-TestReportCommand.ps1`, dùng `actr` / `auto-convert-test-report` để bật **cùng máy chủ** `python serve_ui.py` (cửa sổ PowerShell giữ mở; **Ctrl+C** = tắt server). Dùng giao diện web ở trình duyệt để dán bảng — **không** dán cả bảng vào cửa sổ console.

## Dùng trực tiếp dòng lệnh (không cần web)

Luôn chạy **sau** `Set-Location` vào thư mục này, hoặc dùng **đường dẫn đầy đủ** tới `excel_testplan_to_table.py` (thay bằng path của bạn).

**Bảng TSV — mỗi ô một dòng (không còn xuống dòng bên trong ô):**

```powershell
Get-Clipboard -Raw | Set-Content -Path .\pasted.tsv -Encoding utf8
python .\excel_testplan_to_table.py .\pasted.tsv --format flat-tsv
```

Ghi file:

```powershell
python .\excel_testplan_to_table.py .\pasted.tsv --format flat-tsv -o .\out_flat.tsv
```

**Markdown:**

```powershell
python .\excel_testplan_to_table.py .\pasted.tsv --format markdown -o .\out.md
```

**Cùng máy, không cần `cd`** (dán path đúng tới `excel_testplan_to_table.py` trên máy bạn, một lệnh từ bất kỳ đâu):

```powershell
$py = if (Get-Command python -ErrorAction SilentlyContinue) { "python" } else { "py" }
$script = Join-Path $env:USERPROFILE ".cursor\projects\empty-window\excel-testplan-to-table\excel_testplan_to_table.py"
# hoặc: $script = "D:\dự án\excel-testplan-to-table\excel_testplan_to_table.py"
Get-Clipboard -Raw | & $py $script --format flat-tsv
```

Có thể bỏ file trung gian: script **đọc từ stdin** khi bạn **không** truyền tên file — pipeline phía trên dùng clipboard → stdin.

Cách ổn khi cần lưu lại file gốc:

```powershell
Get-Clipboard -Raw | Set-Content -Path (Join-Path $env:TEMP "pasted_from_excel.tsv") -Encoding utf8
python (Join-Path $env:USERPROFILE "path...\excel_testplan_to_table.py") (Join-Path $env:TEMP "pasted_from_excel.tsv") --format flat-tsv
```

`actr.cmd` (double-click hoặc gõ nếu thư mục có trong `PATH`) cũng mở cùng menu.

## Tùy chọn script Python (tóm tắt)

| Tùy chọn | Mô tả |
|----------|--------|
| `--format flat-tsv` | Một bảng TSV, gộp xuống dòng trong từng ô |
| `--format markdown` | Bảng Markdown (Notion, GitHub) |
| `--multiline br` (mặc định cho markdown) / `space`, `slash`, `keep` | Cách xử lý xuống dòng (markdown) |
| `--skip-section-rows` | Markdown: bỏ dòng cột 1 không phải mã `AUTH-001` / `ADM-001` / … |
| `--bom` + `-o file` | Ghi file UTF-8 có BOM (mở bằng Excel) |
| `--jammed` / `--jammed-out` | Chuỗi dính không tab (cột phụ) |

Xem thêm: `python .\excel_testplan_to_table.py -h`

## Cấu trúc thư mục (tham khảo)

| File / thư mục | Mục đích |
|----------------|----------|
| `excel_testplan_to_table.py` | Script chuyển đổi chính (CLI) |
| `serve_ui.py` | Máy chủ localhost + API; mở trình duyệt |
| `web/` | Giao diện (`index.html`, `style.css`, `app.js`) |
| `TestReportConvert-Menu.ps1` | Bật `serve_ui.py` (dùng với `actr`) |
| `Register-TestReportCommand.ps1` | Cài / gỡ hàm `actr` trong profile |
| `actr.cmd` | Gọi `TestReportConvert-Menu` từ cmd |

---

*Copy-paste thân thiện: dùng `Join-Path $env:USERPROFILE '...'` thay vì dính tên user cụ thể; hoặc mở terminal trực tiếp từ thư mục để bỏ bước `cd`.*
