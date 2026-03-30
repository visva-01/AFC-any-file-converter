from __future__ import annotations

import cgi
import json
import mimetypes
import os
import shutil
import subprocess
import tempfile
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SCRIPT_PATH = ROOT / "scripts" / "convert-office.ps1"
OFFICE_EXTENSIONS = {".ppt", ".pptx", ".doc", ".docx", ".xls", ".xlsx"}


def powershell_command(input_path: Path, output_path: Path) -> list[str]:
    return [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
        "-InputPath",
        str(input_path),
        "-OutputPath",
        str(output_path),
    ]


def is_office_converter_available() -> bool:
    return SCRIPT_PATH.exists()


class ConvertLabHandler(SimpleHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path == "/favicon.ico":
            self.send_response(HTTPStatus.NO_CONTENT)
            self.end_headers()
            return
        if self.path == "/api/status":
            self.send_json(
                {
                    "available": is_office_converter_available(),
                    "officeExtensions": sorted(OFFICE_EXTENSIONS),
                }
            )
            return

        super().do_GET()

    def do_POST(self) -> None:
        if self.path != "/api/convert":
            self.send_error(HTTPStatus.NOT_FOUND, "Unknown endpoint")
            return

        if not is_office_converter_available():
            self.send_json({"error": "Office converter script is missing."}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )

        uploaded = form["file"] if "file" in form else None
        target = form.getvalue("target", "").lower()

        if uploaded is None or not getattr(uploaded, "filename", ""):
            self.send_json({"error": "No file uploaded."}, HTTPStatus.BAD_REQUEST)
            return

        if target != "pdf":
            self.send_json({"error": "Only PDF output is supported for Office files."}, HTTPStatus.BAD_REQUEST)
            return

        source_name = Path(uploaded.filename).name
        source_extension = Path(source_name).suffix.lower()
        if source_extension not in OFFICE_EXTENSIONS:
            self.send_json({"error": "This desktop helper currently supports Office to PDF only."}, HTTPStatus.BAD_REQUEST)
            return

        with tempfile.TemporaryDirectory(prefix="convertlab-") as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / source_name
            output_path = temp_path / f"{input_path.stem}.pdf"

            with input_path.open("wb") as file_handle:
                shutil.copyfileobj(uploaded.file, file_handle)

            completed = subprocess.run(
                powershell_command(input_path, output_path),
                capture_output=True,
                text=True,
                timeout=180,
                cwd=str(ROOT),
            )

            if completed.returncode != 0 or not output_path.exists():
                error_message = completed.stderr.strip() or completed.stdout.strip() or "Conversion failed."
                self.send_json({"error": error_message}, HTTPStatus.INTERNAL_SERVER_ERROR)
                return

            self.send_file(output_path, download_name=output_path.name)

    def send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, download_name: str) -> None:
        data = path.read_bytes()
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.end_headers()
        self.wfile.write(data)

    def translate_path(self, path: str) -> str:
        raw = super().translate_path(path)
        translated = Path(raw)
        try:
            translated.relative_to(Path.cwd())
        except ValueError:
            return str(ROOT)
        return str(translated)


def main() -> None:
    port = int(os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer(("127.0.0.1", port), ConvertLabHandler)
    print(f"ConvertLab running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()

