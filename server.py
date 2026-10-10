#!/usr/bin/env python3
"""Serve the dashboard and persist shared internal insight reports."""

from __future__ import annotations

import argparse
import json
import os
import tempfile
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit


API_PATH = "/api/internal-insights"
MAX_BODY_BYTES = 2 * 1024 * 1024
REPORT_LOCK = threading.Lock()


def normalize_reports(value):
    if not isinstance(value, dict):
        return {}
    return {
        str(key): report
        for key, report in value.items()
        if str(key).strip() and isinstance(report, dict)
    }


def read_reports(path: Path):
    try:
        return normalize_reports(json.loads(path.read_text(encoding="utf-8")))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def write_reports(path: Path, reports):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(reports, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except OSError:
            pass
        raise


class DashboardHandler(SimpleHTTPRequestHandler):
    reports_path: Path

    def _send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _request_path(self):
        return urlsplit(self.path).path

    def do_GET(self):  # noqa: N802
        if self._request_path() == API_PATH:
            with REPORT_LOCK:
                reports = read_reports(self.reports_path)
            self._send_json(reports)
            return
        super().do_GET()

    def do_POST(self):  # noqa: N802
        if self._request_path() != API_PATH:
            self._send_json({"error": "Not found"}, 404)
            return
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0
        if content_length <= 0 or content_length > MAX_BODY_BYTES:
            self._send_json({"error": "Invalid request body"}, 413)
            return
        try:
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            key = str(payload.get("key", "")).strip()
            report = payload.get("report")
            if not key or not isinstance(report, dict):
                raise ValueError("key and report are required")
            with REPORT_LOCK:
                reports = read_reports(self.reports_path)
                reports[key] = report
                write_reports(self.reports_path, reports)
            self._send_json(reports)
        except (UnicodeDecodeError, json.JSONDecodeError, TypeError, ValueError) as error:
            self._send_json({"error": str(error)}, 400)
        except OSError as error:
            self._send_json({"error": f"Unable to save report: {error}"}, 500)


def main():
    parser = argparse.ArgumentParser(description="Serve the Slot dashboard with shared insight storage")
    parser.add_argument("--bind", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--directory", type=Path, default=Path.cwd())
    args = parser.parse_args()

    directory = args.directory.resolve()
    reports_path = directory / "data" / "internal-test-insights.json"
    handler = lambda *handler_args, **handler_kwargs: DashboardHandler(
        *handler_args, directory=str(directory), **handler_kwargs
    )
    DashboardHandler.reports_path = reports_path
    server = ThreadingHTTPServer((args.bind, args.port), handler)
    print(f"Dashboard server listening on http://{args.bind}:{args.port}/", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
