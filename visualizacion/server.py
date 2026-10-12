"""Local-only server and one-click launcher for the pharmacy market map."""
from __future__ import annotations

import argparse
import functools
import gzip
import hashlib
import json
import logging
import mimetypes
import os
from pathlib import Path
import subprocess
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.error import URLError
from urllib.parse import unquote, urlsplit
from urllib.request import urlopen
import webbrowser

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "app" / "dist"
DATA = ROOT / "data"
APP_ID = "farmacias-potencial-2027"
PROJECT_KEY = hashlib.sha256(str(ROOT).casefold().encode()).hexdigest()[:16]


class DataRepository:
    """Load the compact allocation ledger only when a pharmacy is inspected."""

    def __init__(self, data_dir: Path = DATA):
        self.data_dir = data_dir
        self._ledger = None
        self._store_index = None
        self._names = None
        self._signature = None
        self._lock = threading.Lock()

    def flows(self, store_id: str):
        with self._lock:
            flow_path, meta_path = self.data_dir / "flows.json", self.data_dir / "metadata.json"
            signature = tuple((path.stat().st_mtime_ns, path.stat().st_size) for path in (flow_path, meta_path))
            if self._signature != signature:
                ledger = json.loads(flow_path.read_text(encoding="utf-8"))
                metadata = json.loads(meta_path.read_text(encoding="utf-8"))
                self._ledger = ledger
                self._store_index = {key: i for i, key in enumerate(ledger["store_order"])}
                self._names = metadata.get("sanibrick_names", {})
                self._signature = signature
            ledger, names, store_index = self._ledger, self._names, self._store_index
        if store_id not in store_index:
            raise KeyError(store_id)
        index = store_index[store_id]
        start, end = ledger["indptr"][index:index + 2]
        rows = []
        for offset in range(start, end):
            origin_index = ledger["indices"][offset]
            sb = ledger["origin_order"][origin_index]
            cents = ledger["cents"][offset]
            budget = ledger["origin_budgets_cents"][origin_index]
            rows.append({"sb_code": sb, "sb_name": names.get(sb, sb),
                         "origin_eur": budget / 100, "share": cents / budget if budget else 0,
                         "eur": cents / 100})
        return sorted(rows, key=lambda row: (-row["eur"], row["sb_code"]))


@functools.lru_cache(maxsize=8)
def compressed_file(path: str, modified_ns: int, size: int):
    del modified_ns, size  # part of the cache key: refreshed assets invalidate it
    return gzip.compress(Path(path).read_bytes(), compresslevel=4, mtime=0)


def make_handler(dist_dir: Path = DIST, data_dir: Path = DATA):
    repository = DataRepository(data_dir)

    class Handler(BaseHTTPRequestHandler):
        server_version = "PharmacyMap/1.0"

        def log_message(self, format, *args):
            # Avoid writing private identifiers, search terms or file contents to logs.
            logging.info("HTTP response %s", args[1] if len(args) > 1 else "request")

        def send_body(self, body: bytes, content_type: str, status=200, encoding=None):
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Cache-Control", "no-cache")
            if encoding:
                self.send_header("Content-Encoding", encoding)
            self.end_headers()
            if self.command != "HEAD":
                try:
                    self.wfile.write(body)
                except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                    # A closed or reloaded browser tab may abort a large snapshot.
                    return

        def send_json(self, value, status=200):
            self.send_body(json.dumps(value, ensure_ascii=False, allow_nan=False).encode("utf-8"),
                           "application/json; charset=utf-8", status)

        def do_HEAD(self):
            self.do_GET()

        def do_GET(self):
            request = urlsplit(self.path)
            path = unquote(request.path)
            # No directory listings, arbitrary workspace files, or fallback to private source code.
            if "\x00" in path or "\\" in path or any(part == ".." for part in path.split("/")):
                self.send_json({"error": "Ruta no permitida"}, 400)
                return
            if path == "/api/health":
                self.send_json({"app_id": APP_ID, "project_key": PROJECT_KEY,
                                "ready": (dist_dir / "index.html").exists() and (data_dir / "metadata.json").exists()})
                return
            if path.startswith("/api/flows/"):
                try:
                    self.send_json(repository.flows(path.removeprefix("/api/flows/")))
                except KeyError:
                    self.send_json({"error": "Establecimiento desconocido"}, 404)
                except (OSError, ValueError, TypeError):
                    logging.exception("Allocation ledger could not be loaded")
                    self.send_json({"error": "El detalle de orígenes no está disponible"}, 503)
                return
            base = data_dir if path.startswith("/data/") else dist_dir
            relative = path.removeprefix("/data/") if base == data_dir else path.lstrip("/")
            target = (base / (relative or "index.html")).resolve()
            if not target.is_relative_to(base.resolve()) or not target.is_file():
                self.send_json({"error": "Recurso no encontrado"}, 404)
                return
            if base == data_dir and target.suffix.lower() not in {".json", ".geojson", ".csv"}:
                self.send_json({"error": "Recurso no permitido"}, 403)
                return
            content_type = {".geojson": "application/geo+json", ".json": "application/json", ".csv": "text/csv",
                            ".js": "text/javascript", ".css": "text/css"}.get(target.suffix.lower()) or mimetypes.guess_type(target.name)[0] or "application/octet-stream"
            if target.suffix in {".json", ".geojson", ".html", ".css", ".js", ".csv"}:
                content_type += "; charset=utf-8"
            info = target.stat()
            if "gzip" in self.headers.get("Accept-Encoding", "") and info.st_size > 1024:
                self.send_body(compressed_file(str(target), info.st_mtime_ns, info.st_size), content_type, encoding="gzip")
            else:
                self.send_body(target.read_bytes(), content_type)

    return Handler


def probe(port: int, timeout: float = 0.35):
    try:
        with urlopen(f"http://127.0.0.1:{port}/api/health", timeout=timeout) as response:
            return json.load(response)
    except (URLError, TimeoutError, ValueError, OSError):
        return None


def launch(preferred_port: int):
    if not (DIST / "index.html").exists() or not (DATA / "metadata.json").exists():
        raise RuntimeError("Faltan los datos o la aplicación compilada. Consulta visualizacion/README.md.")
    import socket
    free_port = None
    for port in range(preferred_port, preferred_port + 25):
        health = probe(port)
        if health and health.get("app_id") == APP_ID and health.get("project_key") == PROJECT_KEY:
            webbrowser.open(f"http://127.0.0.1:{port}/")
            print(f"Mapa abierto: http://127.0.0.1:{port}/")
            return
        try:
            with socket.socket() as sock:
                sock.bind(("127.0.0.1", port))
            free_port = port
            break
        except OSError:
            # A busy local server may miss the quick probe while serving the
            # full snapshot. Confirm its identity before spawning another one.
            health = probe(port, timeout=3)
            if health and health.get("app_id") == APP_ID and health.get("project_key") == PROJECT_KEY:
                webbrowser.open(f"http://127.0.0.1:{port}/")
                print(f"Mapa abierto: http://127.0.0.1:{port}/")
                return
            continue
    if free_port is None:
        raise RuntimeError("No hay un puerto local disponible entre los 25 puertos comprobados.")
    log_dir = ROOT / "logs"
    log_dir.mkdir(exist_ok=True)
    python = Path(sys.executable)
    pythonw = python.with_name("pythonw.exe")
    if os.name == "nt" and pythonw.exists():
        python = pythonw
    with (log_dir / "server.log").open("a", encoding="utf-8") as output:
        child = subprocess.Popen([str(python), str(Path(__file__).resolve()), "--port", str(free_port)],
                                 cwd=ROOT, stdout=output, stderr=output,
                                 creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)
    (log_dir / "server.json").write_text(json.dumps({"pid": child.pid, "port": free_port,
                                                      "project_key": PROJECT_KEY}), encoding="utf-8")
    for _ in range(100):
        health = probe(free_port)
        if health and health.get("project_key") == PROJECT_KEY and health.get("ready"):
            webbrowser.open(f"http://127.0.0.1:{free_port}/")
            print(f"Mapa abierto: http://127.0.0.1:{free_port}/")
            return
        if child.poll() is not None:
            break
        time.sleep(0.1)
    raise RuntimeError("El mapa no ha arrancado. Consulta visualizacion/logs/server.log.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--launch", action="store_true", help="Start in background and open the default browser")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    if args.launch:
        launch(args.port)
    else:
        with ThreadingHTTPServer(("127.0.0.1", args.port), make_handler()) as server:
            logging.info("Pharmacy map ready on local port %s", args.port)
            server.serve_forever()


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, OSError) as error:
        print(str(error), file=sys.stderr)
        sys.exit(1)
