/**
 * Test report web UI — cùng origin với trang (http) hoặc 127.0.0.1:PORT nếu mở từ file://
 */
(function () {
  "use strict";

  var $input = document.getElementById("input");
  var $output = document.getElementById("output");
  var $status = document.getElementById("status");
  var $toast = document.getElementById("toast");
  var $btnCopy = document.getElementById("btnCopy");
  var $btnTsv = document.getElementById("btnTsv");
  var $btnMd = document.getElementById("btnMd");
  var $serverBanner = document.getElementById("serverBanner");
  var $errorLogCard = document.getElementById("errorLogCard");
  var $errorLog = document.getElementById("errorLog");
  var $btnCopyErrorLog = document.getElementById("btnCopyErrorLog");
  var $btnClearErrorLog = document.getElementById("btnClearErrorLog");
  var $matrixInput = document.getElementById("matrixInput");
  var $btnMatrixClear = document.getElementById("btnMatrixClear");
  var $btnMatrixXlsx = document.getElementById("btnMatrixXlsx");

  var API_PORT = (function () {
    var m = /[?&]port=(\d+)/.exec(window.location.search || "");
    if (m) {
      return m[1];
    }
    try {
      return window.localStorage.getItem("testReportUiPort") || "8765";
    } catch (e) {
      return "8765";
    }
  })();

  /**
   * Cổng gọi API: ?port= ghi đè; trang http dùng cổng trang; file:// dùng localStorage mặc định.
   */
  function getResolvedApiPort() {
    var m = /[?&]port=(\d+)/.exec(window.location.search || "");
    if (m) {
      return m[1];
    }
    if (location.protocol === "file:") {
      return API_PORT;
    }
    if (location.port) {
      return location.port;
    }
    return location.protocol === "https:" ? "443" : "80";
  }

  function getApiBase() {
    if (location.protocol === "file:") {
      return "http://127.0.0.1:" + API_PORT;
    }
    var m = /[?&]port=(\d+)/.exec(window.location.search || "");
    var h = location.hostname;
    if (h === "127.0.0.1" || h === "localhost" || h === "[::1]") {
      // Cùng origin: trang `localhost` mà gọi API `127.0.0.1` = khác origin — Chrome dễ chặn fetch POST.
      if (m) {
        return location.protocol + "//" + h + ":" + m[1];
      }
      return location.origin;
    }
    return location.protocol + "//" + location.host;
  }

  function setStatus(msg, type) {
    $status.textContent = msg || "";
    $status.className = "status" + (type ? " " + type : "");
  }

  function showToast(msg) {
    $toast.textContent = msg;
    $toast.hidden = false;
    $toast.classList.add("show");
    clearTimeout(showToast._t);
    showToast._t = setTimeout(function () {
      $toast.classList.remove("show");
      $toast.hidden = true;
    }, 2400);
  }

  function apiUrl(path) {
    return getApiBase() + path;
  }

  /** Thông điệp dài khi nút Chuyển TSV/MD thất bại (cần server). */
  function helpServerDown() {
    var p = getResolvedApiPort();
    return (
      "Không tới được API tại " +
      getApiBase() +
      " — bật terminal trong thư mục công cụ, gõ: python serve_ui.py rồi mở trình duyệt tại " +
      "http://localhost:" +
      p +
      "/ (cổng khác: $env:TEST_REPORT_UI_PORT=8766; python serve_ui.py). " +
      "Mở từ file (file://) mà cần chuyển TSV/Markdown hoặc Tải Excel matrix thì cũng phải bật server ở cổng này."
    );
  }

  /** Một dòng: kiểm tra /health thất bại — không cấm «Sinh prompt» (chạy nội bộ trình duyệt). */
  function statusKhiChuaKetNoiServer() {
    return (
      "Chưa tới " +
      getApiBase() +
      " — chỉ cần khi bấm «Chuyển → TSV / Markdown» hoặc «Tải Excel (.xlsx) đúng format». «Sinh prompt cho AI» vẫn dùng được (không cần server)."
    );
  }

  function errDetail(e) {
    if (e == null) {
      return "";
    }
    var s = (e && e.message) || String(e);
    if (s && s !== "TypeError: Failed to fetch" && s !== "Failed to fetch") {
      return " Chi tiết: " + s;
    }
    return "";
  }

  function clearErrorLog() {
    if ($errorLog) {
      $errorLog.textContent = "";
    }
    if ($errorLogCard) {
      $errorLogCard.hidden = true;
    }
    if ($btnCopyErrorLog) {
      $btnCopyErrorLog.disabled = true;
    }
  }

  function showErrorLog(text) {
    if (!$errorLog || !$errorLogCard) {
      return;
    }
    $errorLog.textContent = text || "(trống)";
    $errorLogCard.hidden = false;
    if ($btnCopyErrorLog) {
      $btnCopyErrorLog.disabled = !text;
    }
    try {
      $errorLogCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
    } catch (e2) {}
  }

  function envInfo() {
    return [
      "location.href: " + location.href,
      "protocol: " + location.protocol,
      "navigator.onLine: " + String(navigator.onLine),
      "apiBase: " + getApiBase(),
      "apiPort (resolved): " + getResolvedApiPort(),
      "API_PORT (file/LS): " + API_PORT,
      "userAgent: " + String(navigator.userAgent || ""),
    ].join("\n");
  }

  function formatException(e) {
    if (e == null) {
      return "(exception null)";
    }
    var lines = [
      "name: " + (e.name || "(no name)"),
      "message: " + (e.message || String(e)),
    ];
    if (e.cause !== undefined) {
      lines.push("cause: " + String(e.cause));
    }
    if (e.stack) {
      lines.push("stack:\n" + e.stack);
    }
    return lines.join("\n");
  }

  async function readResponseBody(res) {
    var raw = "";
    try {
      raw = await res.text();
    } catch (e) {
      raw = "(không đọc được body: " + formatException(e) + ")";
    }
    var data = null;
    var jsonErr = "";
    if (raw) {
      try {
        data = JSON.parse(raw);
      } catch (je) {
        jsonErr = je && je.message ? String(je.message) : String(je);
      }
    }
    return { raw: raw, data: data, jsonErr: jsonErr };
  }

  var MAX_ERROR_BODY = 12000;

  function truncate(s, n) {
    if (!s || s.length <= n) {
      return s;
    }
    return s.slice(0, n) + "\n\n... [cắt bớt " + (s.length - n) + " ký tự]";
  }

  function buildHttpErrorLog(op, method, url, res, body) {
    var ct = "";
    try {
      if (res.headers && res.headers.get) {
        ct = res.headers.get("content-type") || "";
      }
    } catch (e0) {
      ct = String(e0);
    }
    var lines = [
      "=== Lỗi HTTP ===",
      "operation: " + op,
      "method: " + method,
      "url: " + url,
      "status: " + res.status + " " + (res.statusText || ""),
      "content-type: " + (ct || "(none)"),
      "",
      "--- body thô (raw) ---",
      truncate(body.raw, MAX_ERROR_BODY),
    ];
    if (body.jsonErr) {
      lines.push("");
      lines.push("--- lỗi parse JSON (body không phải JSON hợp lệ) ---");
      lines.push(body.jsonErr);
    }
    if (body.data !== null && body.data !== undefined) {
      lines.push("");
      lines.push("--- JSON (nếu parse được) ---");
      try {
        lines.push(JSON.stringify(body.data, null, 2));
      } catch (e1) {
        lines.push(String(body.data));
      }
    }
    lines.push("");
    lines.push("--- môi trường ---");
    lines.push(envInfo());
    return lines.join("\n");
  }

  function buildNetworkErrorLog(op, method, url, e) {
    return [
      "=== Lỗi mạng / trình duyệt ===",
      "operation: " + op,
      "method: " + method,
      "url: " + url,
      "",
      formatException(e),
      "",
      "--- môi trường ---",
      envInfo(),
      "",
      helpServerDown(),
    ].join("\n");
  }

  async function checkServer() {
    if (!$serverBanner) {
      return;
    }
    var fromFile = location.protocol === "file:";
    if (fromFile) {
      $serverBanner.className = "server-banner file-hint";
      $serverBanner.innerHTML =
        "Đang mở từ <strong>file (file://)</strong>. " +
        "<strong>«Sinh prompt cho AI»</strong> chạy trên trình duyệt, không cần Python. " +
        "Bấm <strong>«Chuyển → TSV / Markdown»</strong> hoặc <strong>«Tải Excel (.xlsx) đúng format»</strong> thì cần bật <kbd>python serve_ui.py</kbd> " +
        "và mở <a href=\"http://localhost:" + API_PORT + "/\" target=\"_blank\" rel=\"noopener\">http://localhost:" + API_PORT + "/</a> " +
        "hoặc gọi tới cổng <code>" + API_PORT + "</code> từ trang này (CORS). Cổng khác: <code>?port=8766</code>.";
      $serverBanner.hidden = false;
    }
    var healthUrl = apiUrl("/api/health");
    try {
      var r = await fetch(healthUrl, { method: "GET", cache: "no-store" });
      var healthBody = await readResponseBody(r);
      if (r.ok) {
        clearErrorLog();
        $serverBanner.hidden = true;
        setStatus("Đã nối tới server — dán TSV (chuyển / Markdown), matrix → Excel .xlsx, hoặc «Sinh prompt».", "ok");
        return;
      }
      $serverBanner.className = "server-banner file-hint";
      $serverBanner.innerHTML =
        "<strong>Health trả mã " +
        r.status +
        ".</strong> " +
        "Nút <em>Chuyển TSV / Markdown</em> cần <kbd>python serve_ui.py</kbd> tại thư mục công cụ. " +
        "«Sinh prompt cho AI» không phụ thuộc. " +
        "Chi tiết: tải lại trang với F12 &rarr; Network &rarr; <code>health</code>.";
      $serverBanner.hidden = false;
      setStatus(statusKhiChuaKetNoiServer(), "warn");
    } catch (err) {
      $serverBanner.className = "server-banner file-hint";
      $serverBanner.innerHTML =
        "<strong>Chưa bật máy phục vụ ở <code>" + getApiBase() + "</code></strong> (hoặc tường lửa chặn). " +
        "Để dùng <em>Chuyển TSV / Markdown</em>, trong thư mục công cụ chạy: <kbd>python serve_ui.py</kbd> rồi tải lại trang, " +
        "hoặc mở <a href=\"" + getApiBase() + "/\" target=\"_blank\" rel=\"noopener\">trang từ server</a>. " +
        "<strong>«Sinh prompt cho AI»</strong> vẫn bấm được — không cần bước này." +
        errDetail(err);
      $serverBanner.hidden = false;
      setStatus(statusKhiChuaKetNoiServer(), "warn");
    }
  }

  document.getElementById("btnClip").addEventListener("click", async function () {
    setStatus("Đang đọc clipboard…");
    try {
      if (!navigator.clipboard || !navigator.clipboard.readText) {
        setStatus("Trình duyệt không hỗ trợ đọc clipboard. Dùng Ctrl+V trong ô bên dưới.", "err");
        return;
      }
      const t = await navigator.clipboard.readText();
      if (!t) {
        setStatus("Clipboard rỗng. Copy từ Excel (Ctrl+C) trước.", "err");
        return;
      }
      $input.value = t;
      setStatus("Đã tải từ clipboard.", "ok");
    } catch (e) {
      setStatus("Không đọc được clipboard (cần localhost). Dán bằng Ctrl+V.", "err");
    }
  });

  document.getElementById("btnClear").addEventListener("click", function () {
    $input.value = "";
    setStatus("Đã xóa ô nhập.");
  });

  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename || "download.bin";
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      try {
        URL.revokeObjectURL(url);
      } catch (e) {}
    }, 2000);
  }

  if ($btnMatrixClear && $matrixInput) {
    $btnMatrixClear.addEventListener("click", function () {
      $matrixInput.value = "";
      setStatus("Đã xóa ô matrix.");
    });
  }
  async function checkHealthQuick() {
    var healthUrl = apiUrl("/api/health");
    var max = 3;
    for (var i = 0; i < max; i++) {
      try {
        var r = await fetch(healthUrl, { method: "GET", cache: "no-store" });
        if (r && r.ok) {
          return true;
        }
      } catch (e) {
        /* thử lại — Chrome đôi khi thấy net::ERR_EMPTY_RESPONSE lần đầu (keep-alive) */
      }
      if (i < max - 1) {
        await new Promise(function (resolve) {
          setTimeout(resolve, 120);
        });
      }
    }
    return false;
  }

  if ($btnMatrixXlsx && $matrixInput) {
    $btnMatrixXlsx.addEventListener("click", async function () {
      var t = $matrixInput.value;
      if (!t || !String(t).trim()) {
        setStatus("Dán bảng matrix (tab) vào ô phía trên rồi bấm lại.", "err");
        $matrixInput.focus();
        return;
      }
      var op = "matrix-excel";
      var method = "POST";
      var url = apiUrl("/api/matrix-excel");
      clearErrorLog();
      var alive = await checkHealthQuick();
      setStatus(
        alive
          ? "Đang tạo file Excel…"
          : "Không thấy /api/health ổn (ERR_EMPTY_RESPONSE?) — vẫn thử tạo file Excel…",
        alive ? "ok" : "warn"
      );
      try {
        var res = await fetch(url, {
          method: method,
          cache: "no-store",
          headers: { "Content-Type": "application/json; charset=utf-8" },
          body: JSON.stringify({ text: t }),
        });
        var ct = "";
        try {
          if (res.headers && res.headers.get) {
            ct = res.headers.get("content-type") || "";
          }
        } catch (e0) {
          ct = String(e0);
        }
        if (res.ok && /spreadsheet|octet-stream/i.test(ct)) {
          var blob = await res.blob();
          downloadBlob(blob, "decision-matrix.xlsx");
          setStatus("Đã tải decision-matrix.xlsx.", "ok");
          showToast("Đã tải decision-matrix.xlsx");
          return;
        }
        var body = await readResponseBody(res);
        var errText = buildHttpErrorLog(op, method, url, res, body);
        showErrorLog(errText);
        setStatus("Tạo Excel thất bại (HTTP " + res.status + "). Xem log lỗi bên dưới.", "err");
      } catch (e) {
        var errText2 =
          buildNetworkErrorLog(op, method, url, e) +
          "\n\n--- Gợi ý (tạo .xlsx) ---\n" +
          "Nếu trang tải từ " +
          getApiBase() +
          " mà vẫn lỗi: terminal chạy serve_ui đã tắt, cổng bị chặn, hoặc tạm mất kết nối. " +
          "Mở lại terminal, chạy: python serve_ui.py, rồi F5 trang rồi bấm lại.";
        showErrorLog(errText2);
        setStatus("Không tới server khi tạo Excel. Kiểm tra serve_ui. Xem log lỗi.", "err");
      }
    });
  }

  $btnCopy.addEventListener("click", async function () {
    const t = $output.value;
    if (!t) return;
    try {
      await navigator.clipboard.writeText(t);
      showToast("Đã sao chép kết quả");
    } catch (e) {
      $output.select();
      document.execCommand("copy");
      showToast("Đã thử sao chép (Ctrl+C nếu cần)");
    }
  });

  async function convert(format) {
    const text = $input.value;
    if (!text.trim()) {
      setStatus("Ô dữ liệu trống. Dán từ Excel hoặc Lấy từ clipboard.", "err");
      return;
    }
    if (text.indexOf("\t") === -1) {
      const ok = window.confirm("Không thấy ký tự TAB — có thể không phải TSV từ Excel. Vẫn thử chuyển đổi?");
      if (!ok) return;
    }

    $btnTsv.disabled = true;
    $btnMd.disabled = true;
    setStatus("Đang chuyển…");
    clearErrorLog();

    var op = "Chuyển đổi (convert) — format=" + format;
    var url = apiUrl("/api/convert");
    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: JSON.stringify({ text: text, format: format }),
      });
      const body = await readResponseBody(res);

      if (!res.ok) {
        const log = buildHttpErrorLog(op, "POST", url, res, body);
        showErrorLog(log);
        $output.value =
          "Lỗi HTTP " + res.status + " — mở khối «Log lỗi kỹ thuật» bên dưới để xem body phản hồi từ server.";
        $btnCopy.disabled = true;
        setStatus("Lỗi máy chủ: " + res.status, "err");
        return;
      }

      const data = body.data;
      if (data && data.ok) {
        clearErrorLog();
        $output.value = data.output || "";
        $btnCopy.disabled = !($output.value && $output.value.length);
        setStatus("Xong. Bấm «Sao chép kết quả» hoặc Ctrl+C trong ô dưới.", "ok");
        return;
      }

      const appLog = [
        "=== Lỗi ứng dụng (HTTP 200 nhưng ok: false) ===",
        "operation: " + op,
        "url: " + url,
        "body thô: " + truncate(body.raw, MAX_ERROR_BODY),
        data ? "--- JSON ---\n" + JSON.stringify(data, null, 2) : "(không parse được JSON — xem body thô)",
        "",
        "--- môi trường ---",
        envInfo(),
      ].join("\n");
      showErrorLog(appLog);
      $output.value =
        (data && (data.error || data.stderr)) ||
        (data && data.message) ||
        "Chuyển đổi thất bại — xem «Log lỗi kỹ thuật» bên dưới.";
      $btnCopy.disabled = !($output.value && $output.value.length);
      setStatus(data && (data.error || data.stderr) ? String(data.error || data.stderr) : "Chuyển đổi thất bại", "err");
    } catch (e) {
      const log = buildNetworkErrorLog(op, "POST", url, e);
      showErrorLog(log);
      $output.value = "Gọi API thất bại — xem «Log lỗi kỹ thuật» bên dưới.";
      $btnCopy.disabled = true;
      setStatus(helpServerDown() + errDetail(e), "err");
    } finally {
      $btnTsv.disabled = false;
      $btnMd.disabled = false;
    }
  }

  $btnTsv.addEventListener("click", function () {
    convert("flat-tsv");
  });
  $btnMd.addEventListener("click", function () {
    convert("markdown");
  });

  var $aiCard = document.getElementById("aiPromptCard");
  var $aiBlocks = document.getElementById("aiPromptBlocks");
  var $aiToc = document.getElementById("aiPromptToc");
  var $aiHint = document.getElementById("aiPromptHint");
  var $btnAiPrompt = document.getElementById("btnAiPrompt");
  var $btnCopyPrompt = document.getElementById("btnCopyPrompt");
  var aiPromptCombined = "";

  function escapeHtml(s) {
    if (s == null) {
      return "";
    }
    var d = document.createElement("div");
    d.textContent = String(s);
    return d.innerHTML;
  }

  function runAiPrompt() {
    var text = $input.value;
    if (!text || !String(text).trim()) {
      setStatus("Dán bảng TSV (Excel) vào ô trên trước.", "err");
      return;
    }
    if (typeof buildAiPromptParts !== "function") {
      setStatus("Thiếu tệp ai_matrix_prompt.js. Tải lại trang (Ctrl+F5).", "err");
      return;
    }
    $btnAiPrompt.disabled = true;
    setStatus("Đang sinh prompt… (chỉ trên trình duyệt, không gọi mạng)");
    clearErrorLog();
    aiPromptCombined = "";

    try {
      var r = buildAiPromptParts(text);
      if (r.error) {
        if ($aiBlocks) {
          $aiBlocks.innerHTML = "";
        }
        if ($aiToc) {
          $aiToc.innerHTML = "";
          $aiToc.hidden = true;
        }
        if ($aiHint) {
          $aiHint.hidden = true;
        }
        var one = document.createElement("div");
        one.className = "ai-prompt-block ai-prompt-block--err";
        var ta0 = document.createElement("textarea");
        ta0.className = "mono ai-prompt-block-ta";
        ta0.readOnly = true;
        ta0.value = r.error;
        one.appendChild(ta0);
        if ($aiBlocks) {
          $aiBlocks.appendChild(one);
        }
        if ($aiCard) {
          $aiCard.hidden = false;
        }
        $btnCopyPrompt.disabled = true;
        setStatus("Không tách được mục từ TSV — xem hướng dẫn trong khối dưới.", "err");
        try {
          if ($aiCard) {
            $aiCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
          }
        } catch (e) {}
        return;
      }

      var parts = r.parts;
      aiPromptCombined = parts
        .map(function (p) {
          return p.markdown;
        })
        .join("\n\n---\n\n");

      if ($aiToc) {
        $aiToc.innerHTML = "";
        for (var t = 0; t < parts.length; t++) {
          var p = parts[t];
          var li = document.createElement("li");
          var a0 = p.ids[0] || "";
          var a1 = p.ids.length ? p.ids[p.ids.length - 1] : "";
          li.innerHTML = "<strong>" + escapeHtml(p.name) + "</strong> — từ <code>" + escapeHtml(a0) + "</code> tới <code>" + escapeHtml(a1) + "</code> (" + p.ids.length + " mã)";
          $aiToc.appendChild(li);
        }
        $aiToc.hidden = false;
      }
      if ($aiHint) {
        $aiHint.hidden = false;
        $aiHint.textContent = "Có " + parts.length + " mục — bạn có thể dán từng prompt riêng (dưới) hoặc dùng «Sao chép tất cả» nếu muốn một lần.";
      }

      if ($aiBlocks) {
        $aiBlocks.innerHTML = "";
        for (var i = 0; i < parts.length; i++) {
          (function (part) {
            var art = document.createElement("article");
            art.className = "ai-prompt-block";
            var head = document.createElement("div");
            head.className = "ai-prompt-block-head";
            var h3 = document.createElement("h3");
            h3.textContent = part.name;
            var idsEl = document.createElement("div");
            idsEl.className = "ai-prompt-ids";
            idsEl.textContent = part.ids.join("  —  ");
            var btn1 = document.createElement("button");
            btn1.type = "button";
            btn1.className = "btn btn-ghost btn-sm";
            btn1.setAttribute("data-copy-ai-block", "1");
            btn1.textContent = "Sao chép prompt này";
            head.appendChild(h3);
            head.appendChild(idsEl);
            head.appendChild(btn1);
            var ta = document.createElement("textarea");
            ta.className = "mono ai-prompt-block-ta";
            ta.readOnly = true;
            ta.value = part.markdown;
            art.appendChild(head);
            art.appendChild(ta);
            btn1.addEventListener("click", async function () {
              try {
                await navigator.clipboard.writeText(ta.value);
                showToast("Đã sao chép: " + part.name);
              } catch (e) {
                ta.select();
                document.execCommand("copy");
                showToast("Chọn toàn bộ (Ctrl+C)");
              }
            });
            $aiBlocks.appendChild(art);
          })(parts[i]);
        }
      }

      if ($aiCard) {
        $aiCard.hidden = false;
      }
      $btnCopyPrompt.disabled = !aiPromptCombined.length;
      setStatus("Đã tạo " + parts.length + " prompt. Dán từng cái — hoặc sao chép tất cả nếu cần.", "ok");
      try {
        if ($aiCard) {
          $aiCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
        }
      } catch (e) {}
    } catch (e) {
      var errMsg = [
        "=== Lỗi sinh prompt (JavaScript) ===",
        e && e.name ? e.name : "",
        e && e.message ? e.message : String(e),
        e && e.stack ? e.stack : "",
        "",
        "--- môi trường ---",
        envInfo(),
      ]
        .filter(function (x) {
          return x;
        })
        .join("\n");
      showErrorLog(errMsg);
      if ($aiBlocks) {
        $aiBlocks.innerHTML = "";
      }
      if ($aiToc) {
        $aiToc.hidden = true;
        $aiToc.innerHTML = "";
      }
      if ($aiHint) {
        $aiHint.hidden = true;
      }
      var ex = document.createElement("div");
      ex.className = "ai-prompt-block";
      var taE = document.createElement("textarea");
      taE.className = "mono ai-prompt-block-ta";
      taE.readOnly = true;
      taE.value = (e && e.message) || "Lỗi khi tạo prompt — xem log bên dưới.";
      ex.appendChild(taE);
      if ($aiBlocks) {
        $aiBlocks.appendChild(ex);
      }
      if ($aiCard) {
        $aiCard.hidden = false;
      }
      $btnCopyPrompt.disabled = true;
      setStatus("Lỗi khi sinh prompt — xem «Log lỗi kỹ thuật» bên dưới.", "err");
    } finally {
      $btnAiPrompt.disabled = false;
    }
  }

  if ($btnAiPrompt) {
    $btnAiPrompt.addEventListener("click", runAiPrompt);
  }
  if ($btnCopyPrompt) {
    $btnCopyPrompt.addEventListener("click", async function () {
      if (!aiPromptCombined) {
        return;
      }
      try {
        await navigator.clipboard.writeText(aiPromptCombined);
        showToast("Đã sao chép tất cả (" + (aiPromptCombined ? aiPromptCombined.length : 0) + " ký tự)");
      } catch (e) {
        showToast("Không sao chép tự động — từng khối bên dưới (Ctrl+C)");
      }
    });
  }

  if ($btnClearErrorLog) {
    $btnClearErrorLog.addEventListener("click", function () {
      clearErrorLog();
    });
  }
  if ($btnCopyErrorLog) {
    $btnCopyErrorLog.addEventListener("click", async function () {
      if (!$errorLog || !$errorLog.textContent) {
        return;
      }
      try {
        await navigator.clipboard.writeText($errorLog.textContent);
        showToast("Đã sao chép log lỗi");
      } catch (e) {
        try {
          $errorLog.focus();
        } catch (e2) {}
        if (window.getSelection) {
          var r = document.createRange();
          r.selectNodeContents($errorLog);
          var sel = window.getSelection();
          if (sel) {
            sel.removeAllRanges();
            sel.addRange(r);
          }
        }
        showToast("Chọn log (Ctrl+C)");
      }
    });
  }

  // Phím tắt: Ctrl+Enter = TSV
  $input.addEventListener("keydown", function (e) {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      convert("flat-tsv");
    }
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      void checkServer();
    });
  } else {
    void checkServer();
  }
})();
