/**
 * Sinh prompt Markdown (bảng quyết định / matrix TSV) — cùng logic với ai_matrix_prompt.py, chạy 100% trên trình duyệt.
 */
(function (global) {
  "use strict";

  var CASE_ID = /^(?:AUTH|ADM|STU|TCH)-\d{3}\s*$/i;

  function parseTsv(text) {
    var s = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    var rows = [];
    if (!s) {
      return rows;
    }
    var len = s.length;
    var i = 0;
    while (i < len) {
      var row = [];
      var field = "";
      var inQ = false;
      while (true) {
        if (i >= len) {
          row.push(field);
          break;
        }
        var c = s[i];
        if (inQ) {
          if (c === '"') {
            if (i + 1 < len && s[i + 1] === '"') {
              field += '"';
              i += 2;
              continue;
            }
            inQ = false;
            i++;
            continue;
          }
          field += c;
          i++;
          continue;
        }
        if (c === '"') {
          inQ = true;
          i++;
          continue;
        }
        if (c === "\t") {
          row.push(field);
          field = "";
          i++;
          continue;
        }
        if (c === "\n") {
          row.push(field);
          i++;
          break;
        }
        field += c;
        i++;
      }
      rows.push(row);
    }
    return rows;
  }

  function padRow(r, w) {
    if (r.length < w) {
      var a = r.slice();
      while (a.length < w) {
        a.push("");
      }
      return a;
    }
    return r.slice(0, w);
  }

  function isHeaderRow(row) {
    if (!row || !row.length) {
      return false;
    }
    var a = (row[0] || "").trim().toLowerCase();
    return a.indexOf("test case") === 0 || a === "id";
  }

  function rowToTest(row) {
    var r = row.length < 15 ? padRow(row, 15) : row.slice(0, 15);
    var tid = (r[0] || "").trim();
    if (!CASE_ID.test(tid)) {
      return null;
    }
    return {
      id: tid,
      title: r[1] || "",
      steps: r[2] || "",
      expected: r[3] || "",
      pre: r[4] || "",
      r1s: r[5] || "",
      r1d: r[6] || "",
      r1t: r[7] || "",
      r2s: r[8] || "",
      r2d: r[9] || "",
      r2t: r[10] || "",
      r3s: r[11] || "",
      r3d: r[12] || "",
      r3t: r[13] || "",
      note: r[14] || "",
    };
  }

  function isSkipSectionRow(first) {
    var t = (first || "").trim();
    if (!t) {
      return true;
    }
    if (/^function\s+.+$/i.test(t)) {
      return true;
    }
    var lo = t.toLowerCase();
    if (lo === "test cases" || lo === "test case") {
      return true;
    }
    return false;
  }

  function anyNonEmpty(row) {
    for (var j = 0; j < row.length; j++) {
      if ((row[j] || "").trim()) {
        return true;
      }
    }
    return false;
  }

  function groupSections(rows) {
    var order = [];
    var buckets = Object.create(null);
    var cur = "Mục chưa gán";

    for (var r = 0; r < rows.length; r++) {
      var raw = rows[r];
      if (!raw || !anyNonEmpty(raw)) {
        continue;
      }
      if (isHeaderRow(raw)) {
        continue;
      }
      var first = (raw[0] || "").trim();
      var t = rowToTest(raw);
      if (t) {
        if (!buckets[cur]) {
          buckets[cur] = [];
          if (order.indexOf(cur) < 0) {
            order.push(cur);
          }
        }
        buckets[cur].push(t);
      } else {
        if (isSkipSectionRow(first)) {
          continue;
        }
        cur = first;
        if (!buckets[cur]) {
          order.push(cur);
          buckets[cur] = [];
        }
      }
    }

    var out = [];
    for (var o = 0; o < order.length; o++) {
      var n = order[o];
      var list = buckets[n] || [];
      if (list.length > 0) {
        out.push({ name: n, tests: list });
      }
    }
    return out;
  }

  function tsvTableLineTestData(t) {
    return (
      "- **" + t.id + "** | *" + t.title + "* | " +
      "Pre-conditions: " + (t.pre || "(trống)") + " | Kết quả mong đợi: " + (t.expected || "(trống)") + " | " +
      "R1: " + t.r1s + " " + t.r1d + " " + t.r1t + " | R2: " + t.r2s + " " + t.r2d + " " + t.r2t + " | R3: " + t.r3s + " " + t.r3d + " " + t.r3t + " | Ghi chú: " + (t.note || "(trống)")
    );
  }

  function rowExcel4ColMatrix(a, b, c, d, matrix) {
    return [a, b, c, d].concat(matrix).join("\t");
  }

  function repeatEmpties(n) {
    var a = [];
    for (var i = 0; i < n; i++) {
      a.push("");
    }
    return a;
  }

  function skeletonTsvForIds(ids, section) {
    if (!ids.length) {
      return "\t( không có mã thử )";
    }
    var n = ids.length;
    var z = repeatEmpties(n);
    var nIds = ids.join(", ");
    var de = repeatEmpties(n);
    for (var di = 0; di < n; di++) {
      de[di] = "(điền)";
    }
    var R = rowExcel4ColMatrix;
    return [
      R("", "", "", "", ids),
      R("Condition", "Precondition", "", "", z),
      R(
        "",
        "",
        "",
        "(điền từng dòng pre-condition, sau đó " + n + " cột " + nIds + " — dấu O, lấy từ cột pre-conditions ở dữ liệu đã nêu)",
        z
      ),
      R("Confirm", "Return", "", "", z),
      R(
        "",
        "",
        "",
        "(điền từng dòng kết quả/return, " + n + " cột O — từ cột «Kết quả mong đợi / Steps» ở từng mã; AI có thể gộp nếu cùng câu chữ).",
        z
      ),
      R("", "Exception", "", "", z),
      R("", "", "", "(AI: thêm mới các dòng exception phù hợp, O đúng cột " + nIds + ".)", z),
      R("", "Log message", "", "", z),
      R("", "", "", "(AI: thêm mới; có thể lấy ý từ cột ghi chú ở TSV; O đúng cột).", z),
      R("", "Result", "", "", z),
      R("", "Type (N: Normal, A: Abnormal, B: Boundary)", "", "", z),
      R("", "Passed/Failed", "", "", z),
      R("", "Executed Date", "", "", de),
      R("", "Defect ID (để trống)", "", "", z),
    ].join("\n");
  }

  function loginSixcolExample(ids) {
    if (ids.length !== 6) {
      return "";
    }
    var a = ids[0];
    var b = ids[1];
    var c = ids[2];
    var d = ids[3];
    var e = ids[4];
    var f = ids[5];
    var z = ["", "", "", "", "", ""];
    var R = rowExcel4ColMatrix;
    return [
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
        [
          "(điền)",
          "(điền)",
          "(điền)",
          "(điền)",
          "(điền)",
          "(điền)",
        ]
      ),
      R("", "Defect ID (để trống)", "", "", z),
    ].join("\n");
  }

  var TSVCHECK_MSG =
    "Không tách được mục/mã nào. Kiểm tra TSV: cột 1 = mã dạng AUTH/TCH/ADM/STU-xxx, " +
    "các dòng tên mục ở cột 1, copy đủ 15 cột từ Excel.";

  function buildMarkdownForOneSection(sec) {
    if (!sec.tests || !sec.tests.length) {
      return "";
    }
    var idList = sec.tests.map(function (t) {
      return t.id;
    });
    var idLine = idList.map(function (x) {
      return "`" + x + "`";
    }).join(" ");
    var n = idList.length;
    var bullet = sec.tests.map(tsvTableLineTestData).join("\n");
    var isLoginSix = String(sec.name).trim().toLowerCase() === "login" && n === 6;
    var skeleton;
    if (isLoginSix) {
      var skMain = loginSixcolExample(idList);
      var skAlt = skeletonTsvForIds(idList, sec.name);
      skeleton = skMain + "\n\n*— Bản rút gọn tùy biến số cột (nếu cần):*\n\n" + skAlt;
    } else {
      skeleton = skeletonTsvForIds(idList, sec.name);
    }
    var codeFence = "```";
    var block = (
      "### " + sec.name + "\n\n" +
      "**Đầu ra bắt buộc:** Một **khối TSV** (cột cách bằng **tab**) — **Copy → dán Excel** («Text tách cột bằng tab») — dạng **bảng quyết định (test design matrix)**: **A** = nhãn cột dọc lớn (`Condition`, `Confirm`, `Result`); **B** = tiêu đề từng khối: `Precondition` / `Return` / `Exception` / `Log message` (các từ này **chỉ ở cột B**, giống mẫu, không thay bằng cột A). Cột D = mô tả dòng, cột E trở đi = **" + n + " cột mã** (" + idLine + ") và dấu **O**. **Tối thiểu** phải có **bảng TSV hoàn chỉnh** này. (Tùy chọn) kèm VBA/Office Script hoặc ghi **ô bắt đầu** nếu khác `A1`. **Không** trả bảng 15 cột «một dòng — một case» (kiểu Admin) — **output prompt này là matrix**.\n\n" +
      "Bạn tạo **bảng quyết định (decision table)**: từ cột E, **mỗi cột = một mã thử** (" + idLine + ").\n\n" +
      "**Nguồn đáng tin (không tự sửa nội dung cột nghiệp vụ, chỉ sắp lên lưới + đánh dấu O nơi trùng):**\n" +
      "- Từ dữ liệu sau: cột *Pre-conditions* và *Kết quả / Steps* ở từng mã.\n\n" +
      "**Bạn tự tạo thêm (AI):**\n" +
      "- Các dòng *Exception* và *Log message* phù hợp, đánh dấu **O** ở cột mã tương ứng.\n" +
      "- Có thể bổ sung dòng mô tả nếu cần, miễn hợp nghiệp vụ từng mã.\n\n" +
      "**Lấy từ dữ liệu đã cho (sẵn, không tự tạo dữ liệu cột sản phẩm nếu TSV đã có):**\n" +
      "- Các cột mã: **R1 / R2 / R3** (Trạng thái, Ngày, Tester) — dùng cho hàng *Result* / *Passed* / *Executed date* theo từng cột mã.\n" +
      "- Cột *Defect ID* — **để trống** ở mọi ô cột mã (hàng này nếu có).\n\n" +
      "Dữ liệu gốc (theo từng mã):\n\n" +
      bullet +
      "\n\n" +
      "**Quy tắc cột (bắt buộc — dán mới lệch nếu sai):**\n" +
      "- **A / B / D:** Cột A nhãn dọc lớn: `Condition`, `Confirm`. Mô tả ở **D**; **E…** = " + n + " cột mã. Vùng Confirm: `Return` / `Exception` / `Log message` ở **B**.\n" +
      "- **Tiêu đề mục Result (1 dòng):** chữ `Result` ở **B**, **C trống** → dán xong **merge B:C** (Ảnh 2). **A** trống hoặc gộp cột mục dọc.\n" +
      "- **Bốn dòng** *Type* / *Passed/Failed* / *Executed Date* / *Defect ID:* nhãn ở **B**, **C–D trống** → **merge B:D** từng dòng; dữ liệu mã ở **E…**.\n" +
      "- Dòng **tiêu đề mã thử** (E…): A–D trước hàng mã thử trống. **Dùng tab thật.**\n\n" +
      "**Định dạng chữ & màu (Excel):** **Tahoma**; cơ sở **8 pt**. Nền **xanh đậm (dark blue 26 / `RGB(0, 32, 96)` / #002060)** + chữ **trắng** + **Tahoma Bold 8** cho: nhãn cột mục **A** (`Condition`, `Confirm`, `Result`) và **hàng tiêu đề mã** (`AUTH-…`, `TCH-…`, `ADM-…`, `STU-…`, từ cột E). **Precondition**, **Return**, **Exception**, **Log message** (ở B): **Tahoma Bold 8, đen**. Mọi chữ còn lại: **Tahoma 8, đen**, không đậm. **Kẻ viền** toàn bảng; merge khối Result như quy tắc cột. **Ô tick (O):** dùng **Data validation (dropdown)**, nguồn danh sách `O` — không gõ O tay. **Google Sheets:** Data → Data validation → List → `O`.\n\n" +
      codeFence + "vb\n" +
      "Sub ApplyMatrixMergesAndStyle()\n" +
      "    Dim DarkBlue26 As Long, h0 As Long, i As Long, lc As Long\n" +
      "    DarkBlue26 = RGB(0, 32, 96)\n" +
      "    h0 = 26\n" +
      "    On Error Resume Next\n" +
      "    With ActiveSheet.UsedRange\n" +
      "        .Font.Name = \"Tahoma\": .Font.Size = 8: .Font.Color = RGB(0, 0, 0): .Font.Bold = False\n" +
      "    End With\n" +
      "    Range(\"B\" & h0 & \":C\" & h0).Merge\n" +
      "    For i = 1 To 4: Range(\"B\" & (h0 + i) & \":D\" & (h0 + i)).Merge: Next i\n" +
      "    With Range(\"A\" & h0 & \":A\" & (h0 + 4))\n" +
      "        .Merge: .Value = \"Result\"\n" +
      "        .Font.Name = \"Tahoma\": .Font.Size = 8: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)\n" +
      "        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter: .Interior.Color = DarkBlue26\n" +
      "    End With\n" +
      "    lc = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column\n" +
      "    If lc >= 5 Then\n" +
      "        With ActiveSheet.Range(ActiveSheet.Cells(1, 5), ActiveSheet.Cells(1, lc))\n" +
      "            .Font.Name = \"Tahoma\": .Font.Size = 8: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)\n" +
      "            .Interior.Color = DarkBlue26\n" +
      "        End With\n" +
      "    End If\n" +
      "    ' Thủ công: Condition, Confirm, Precondition, Return, Exception, Log message (Bold 8 đen ở B)\n" +
      "    ActiveSheet.UsedRange.Borders.LineStyle = xlContinuous\n" +
      "End Sub\n" +
      "\n" +
      "Sub ApplyOTickDropdown()\n" +
      "    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long\n" +
      "    r1 = 2: r2 = 25: c1 = 5: c2 = 10\n" +
      "    On Error Resume Next\n" +
      "    With ActiveSheet.Range(ActiveSheet.Cells(r1, c1), ActiveSheet.Cells(r2, c2))\n" +
      "        .Validation.Delete\n" +
      "        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _\n" +
      "            Operator:=xlBetween, Formula1:=\"O\"\n" +
      "    End With\n" +
      "End Sub\n" +
      codeFence + "\n" +
      "*(Nếu `Formula1:=\"O\"` lỗi, thử cột tham chiếu chứa O. Google Sheets: list `O`.)* *(ColorIndex 26: theme. Office Script: Tahoma, màu, validation O.)*\n\n" +
      "**Khung bảng TSV mẫu** (dán «Text tách cột bằng tab») — " + n + " cột mã (" + idLine + "):\n\n" +
      codeFence + "text\n" +
      skeleton + "\n" +
      codeFence + "\n\n" +
      "**Gợi ý thứ tự trả lời:** (1) TSV 4 + " + n + " cột, (2) merge/màu/VBA, (3) **dropdown O** (ApplyOTickDropdown hoặc Data validation thủ công), (4) khác A1: chỉnh h0/r1–r2. **Không** bỏ merge B:C / B:D ở khối Result."
    );
    return block.replace(/\s+$/, "");
  }

  /**
   * @returns {{ error: string }} | {{ parts: Array<{ name: string, ids: string[], markdown: string }> }}
   */
  function buildAiPromptParts(tsvText) {
    var text = String(tsvText).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    var rows = parseTsv(text);
    var sections = groupSections(rows);
    if (!sections.length) {
      return { error: TSVCHECK_MSG };
    }
    var out = [];
    for (var s = 0; s < sections.length; s++) {
      var sec = sections[s];
      var md = buildMarkdownForOneSection(sec);
      if (md) {
        var ids = (sec.tests || []).map(function (t) {
          return t.id;
        });
        out.push({ name: String(sec.name), ids: ids, markdown: md });
      }
    }
    if (!out.length) {
      return { error: TSVCHECK_MSG };
    }
    return { parts: out };
  }

  function buildAiPromptMarkdown(tsvText) {
    var r = buildAiPromptParts(tsvText);
    if (r.error) {
      return r.error;
    }
    return r.parts
      .map(function (p) {
        return p.markdown;
      })
      .join("\n\n---\n\n");
  }

  global.buildAiPromptMarkdown = buildAiPromptMarkdown;
  global.buildAiPromptParts = buildAiPromptParts;
})(typeof window !== "undefined" ? window : this);
