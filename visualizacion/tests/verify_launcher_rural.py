"""Read-only rural-delivery verification after two actual launcher executions.

The first execution starts a fresh hidden server after stopping only verified
QA servers of this project. The second execution must reuse that process.
"""
from pathlib import Path
from urllib.request import urlopen
import argparse
import hashlib
import json

ROOT = Path(__file__).resolve().parents[2]

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--expected-pid", type=int, required=True)
    args = parser.parse_args()
    server = json.loads((ROOT / "visualizacion/logs/server.json").read_text())
    base = f"http://127.0.0.1:{server['port']}"
    with urlopen(base + "/api/health", timeout=15) as response:
        health = json.load(response)
    with urlopen(base + "/data/metadata.json", timeout=15) as response:
        metadata = json.load(response)
    with urlopen(base + "/", timeout=60) as response:
        served_sha = hashlib.sha256(response.read()).hexdigest()
    html_sha = hashlib.sha256((ROOT / "visualizacion/app/dist/index.html").read_bytes()).hexdigest()
    snapshot = json.loads((ROOT / "visualizacion/app/src/data.json").read_text(encoding="utf-8"))
    manifest = json.loads((ROOT / "pharmacy/output/current/manifest.json").read_text(encoding="utf-8"))
    book = ROOT / "pharmacy/output/current/UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx"
    book_sha = hashlib.sha256(book.read_bytes()).hexdigest()
    expected_health = {"app_id": "farmacias-potencial-2027", "project_key": "c338604c1b558b7e", "ready": True}
    checks = {
        "own_server_ready": health == expected_health,
        "current_version": metadata.get("version") == manifest["version"] == "2027.3",
        "current_book_hash": metadata["excel"]["sha256"] == book_sha == manifest["files_sha256"][book.name],
        "served_html_matches_local_build": served_sha == html_sha,
        "complete_commercial_app": snapshot["buildStatus"] == "complete" and snapshot["commercialReport"],
        "second_launch_reused_server": server["pid"] == args.expected_pid and server["port"] == 8765,
        "full_universe_retained": metadata["total_stores"] == 22316 and metadata["total_crm"] == 22356,
    }
    report = {
        "date": "2026-10-07", "launcher": "arrancar_app.cmd", "executions": 2,
        "start_method": "Actual Windows CMD; fresh hidden server after stopping only verified project QA processes; browser opened twice; second launch reused server",
        "first_launch_server": {"pid": args.expected_pid, "port": 8765},
        "second_launch_server": server, "health": health, "version": metadata.get("version"),
        "total_stores": metadata["total_stores"], "total_crm": metadata["total_crm"],
        "national_market_eur": metadata["total_market_eur"], "workbook_sha256": book_sha,
        "html_sha256": html_sha, "checks": checks, "passed": all(checks.values()),
        "left_running": True,
    }
    (ROOT / "visualizacion/tests/launcher_rural_controls.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(report, ensure_ascii=False))
    assert report["passed"], report

if __name__ == "__main__":
    main()
