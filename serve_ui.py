#!/usr/bin/env python3
"""Giao diện web local: phục vụ web/ + /api/convert, /api/matrix-excel. Chạy: python serve_ui.py
Yêu cầu: pip install fastapi uvicorn[standard] openpyxl
"""

from __future__ import annotations

import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import threading
import traceback
import webbrowser
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response
from fastapi.staticfiles import StaticFiles

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
WEB_DIR = ROOT / "web"
CONVERT = ROOT / "excel_testplan_to_table.py"
MATRIX_XLSX = ROOT / "matrix_excel_export.py"
PORT = int(os.environ.get("TEST_REPORT_UI_PORT", "8765"))
HOST = os.environ.get("TEST_REPORT_UI_HOST", "127.0.0.1")

app = FastAPI(
    title="Test report local UI",
    docs_url=None,
    redoc_url=None,
    openapi_url=None,
)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _json_response(data: dict, status_code: int = 200) -> JSONResponse:
    return JSONResponse(
        content=data,
        status_code=status_code,
        # Đồng bộ: escape ASCII cho dữ liệu kỳ lạ từ Excel
        # (FastAPI mặc định vẫn UTF-8; giữ tương đương cũ bằng cách dùng json.dumps tùy biến nếu cần)
    )


def _api_convert(data: dict) -> JSONResponse:
    text = data.get("text", "")
    fmt = data.get("format", "flat-tsv")
    if fmt not in ("flat-tsv", "markdown"):
        return _json_response({"ok": False, "error": "format must be flat-tsv or markdown."}, 400)
    if not (text and str(text).strip()):
        return _json_response(
            {"ok": False, "error": "Empty text.", "output": "", "stderr": ""},
            200,
        )
    if not CONVERT.is_file():
        return _json_response({"ok": False, "error": f"Converter not found: {CONVERT}"}, 500)

    with tempfile.NamedTemporaryFile("w", encoding="utf-8", delete=False, suffix=".tsv", newline="") as f:
        f.write(text)
        tmp = f.name
    try:
        r = subprocess.run(
            [sys.executable, str(CONVERT), tmp, "--format", fmt],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass

    out: dict = {
        "ok": r.returncode == 0,
        "output": r.stdout or "",
        "stderr": (r.stderr or "").strip(),
    }
    if r.returncode != 0:
        out["error"] = (r.stderr or "Conversion failed").strip() or f"exit {r.returncode}"
    return _json_response(out, 200)


def _api_matrix_excel(data: dict) -> Response:
    if not MATRIX_XLSX.is_file():
        return _json_response({"ok": False, "error": f"Missing {MATRIX_XLSX.name}"}, 500)
    text = data.get("text", "")
    if not (text and str(text).strip()):
        return _json_response({"ok": False, "error": "Empty text."}, 200)
    spec = importlib.util.spec_from_file_location("matrix_excel_export", str(MATRIX_XLSX))
    if spec is None or spec.loader is None:  # pragma: no cover
        return _json_response({"ok": False, "error": "Could not load matrix_excel_export."}, 500)
    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
    except Exception as e:  # noqa: BLE001
        return _json_response({"ok": False, "error": f"matrix_excel_export import: {e!s}"}, 500)
    try:
        raw = mod.tsv_to_xlsx_bytes(text)
    except Exception as e:  # noqa: BLE001
        return _json_response({"ok": False, "error": str(e)}, 400)
    return Response(
        content=raw,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="decision-matrix.xlsx"'},
    )


@app.get("/favicon.ico", include_in_schema=False)
def favicon() -> Response:
    return Response(status_code=204)


@app.get("/api/health", include_in_schema=False)
def health() -> dict:
    return {"ok": True, "port": PORT, "url": f"http://localhost:{PORT}/"}


@app.post("/api/convert", include_in_schema=False)
async def api_convert(request: Request) -> JSONResponse:
    if int(request.headers.get("Content-Length", 0) or 0) > 12_000_000:
        return _json_response({"ok": False, "error": "Input too large (max 12 MB)."}, 413)
    try:
        body = (await request.body()).decode("utf-8", errors="replace")
        data = json.loads(body) if body.strip() else {}
    except (json.JSONDecodeError, UnicodeError):
        return _json_response({"ok": False, "error": "Invalid JSON / encoding."}, 400)
    try:
        return _api_convert(data)
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        return _json_response(
            {"ok": False, "error": "Internal server error. See terminal for traceback."},
            500,
        )


@app.post("/api/matrix-excel", include_in_schema=False)
async def api_matrix_excel(request: Request) -> Response:
    if int(request.headers.get("Content-Length", 0) or 0) > 12_000_000:
        return _json_response({"ok": False, "error": "Input too large (max 12 MB)."}, 413)
    try:
        body = (await request.body()).decode("utf-8", errors="replace")
        data = json.loads(body) if body.strip() else {}
    except (json.JSONDecodeError, UnicodeError):
        return _json_response({"ok": False, "error": "Invalid JSON / encoding."}, 400)
    try:
        return _api_matrix_excel(data)
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        return _json_response({"ok": False, "error": "Internal server error (matrix). See terminal."}, 500)


# Thư tự: API trên; static xuống dưới (catch-all)
app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="static")


def main() -> int:
    if sys.platform == "win32":
        try:
            for _stream in (sys.stdout, sys.stderr):
                if hasattr(_stream, "reconfigure"):
                    _stream.reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError, AttributeError):
            pass
    if not WEB_DIR.is_dir() or not (WEB_DIR / "index.html").is_file():
        print("Missing web/ folder with index.html", file=sys.stderr)
        return 1
    try:
        import uvicorn
    except ImportError as e:  # pragma: no cover
        print("Cần: pip install fastapi uvicorn[standard] openpyxl\n" + str(e), file=sys.stderr)
        return 1

    def open_browser() -> None:
        import time

        time.sleep(0.5)
        webbrowser.open(f"http://localhost:{PORT}/")

    threading.Thread(target=open_browser, daemon=True).start()
    if HOST in ("0.0.0.0", "[::]"):
        print("Lưu ý: đang lắng mọi giao diện — chỉ dùng trên mạng tin cậy.")
    print(
        "Test Report (FastAPI): http://localhost:%d/  (host=%s) — bấm Ctrl+C dừng"
        % (PORT, HOST)
    )

    try:
        uvicorn.run(
            app,
            host=HOST,
            port=PORT,
            log_level="warning",
            access_log=False,
        )
    except OSError as e:
        print(
            f"Port {PORT} in use hoặc không bind {HOST!r}. "
            f"Thử: $env:TEST_REPORT_UI_PORT=8766; $env:TEST_REPORT_UI_HOST=127.0.0.1; python serve_ui.py\n{e}",
            file=sys.stderr,
        )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
