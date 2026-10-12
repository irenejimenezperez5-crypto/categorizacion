"""Audit staged Git blobs before publishing this code-only repository."""
from __future__ import annotations

import os
from pathlib import Path
import re
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
ALLOWED_JSON = {
    "pharmacy/categorization_config.json",
    "visualizacion/app/package.json",
    "visualizacion/app/package-lock.json",
    "visualizacion/app/protected-runtime.json",
}
ALLOWED_SUFFIXES = {".py", ".js", ".jsx", ".mjs", ".css", ".md", ".svg", ".toml", ".lock", ".cmd"}
PRIVATE_PARTS = {"input", "output", "analysis", "rx", "data", "cache", "loop", ".openai",
                 "dist", "build", "node_modules", ".venv", "__pycache__", ".data-app-offline"}
SENSITIVE_PATTERNS = {
    "embedded financial result": re.compile(r"\b\d{1,3}(?:\.\d{3}){2,},\d{2}\b"),
    "private key": re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    "GitHub credential": re.compile(r"(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,})"),
    "OpenAI credential": re.compile(r"\bsk-(?:proj-)?[A-Za-z0-9_-]{30,}"),
    "local user path": re.compile(r"[A-Za-z]:[\\/]+Users[\\/]+[^\s\"'<>]+", re.I),
    "corporate email": re.compile(r"[A-Z0-9._%+-]+@essity\.[A-Z.]+", re.I),
    "literal credential": re.compile(r"(?:api[_-]?key|password|secret|access[_-]?token)\s*[:=]\s*[\"'][A-Za-z0-9_+/=-]{16,}[\"']", re.I),
}


def find_git():
    found = shutil.which("git")
    if found:
        return found
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    for relative in ("CodexGitPortable/cmd/git.exe", "CodexGit/cmd/git.exe", "Programs/Git/cmd/git.exe"):
        candidate = local / relative
        if candidate.is_file():
            return str(candidate)
    raise RuntimeError("Git no está disponible. Añádelo al PATH.")


def main():
    git = find_git()
    records = subprocess.check_output([git, "ls-files", "--stage", "-z"], cwd=ROOT).split(b"\0")
    entries = []
    problems = []
    for record in filter(None, records):
        header, raw_path = record.split(b"\t", 1)
        mode, sha, stage = header.split()
        path = raw_path.decode("utf-8")
        entries.append((path, sha))
        parts = Path(path).parts
        if mode != b"100644" or stage != b"0":
            problems.append((path, "unexpected Git mode/stage"))
        if any(part in PRIVATE_PARTS for part in parts):
            problems.append((path, "private/generated directory"))
        allowed = (path == ".gitignore" or path in ALLOWED_JSON or
                   path == "visualizacion/app/index.html" or
                   (Path(path).suffix in ALLOWED_SUFFIXES) or
                   path == "visualizacion/app/src/content/assets/leaflet/LICENSE")
        if not allowed:
            problems.append((path, "unreviewed file type"))
    if not entries:
        problems.append(("index", "no staged/tracked files"))
    request = b"".join(sha + b"\n" for _, sha in entries)
    response = subprocess.check_output([git, "cat-file", "--batch"], input=request, cwd=ROOT)
    offset = 0
    byte_count = 0
    python_count = 0
    for path, _ in entries:
        end = response.index(b"\n", offset)
        _, kind, raw_size = response[offset:end].split()
        size = int(raw_size)
        blob = response[end + 1:end + 1 + size]
        offset = end + size + 2
        byte_count += size
        if kind != b"blob" or b"\0" in blob:
            problems.append((path, "binary/non-blob content"))
            continue
        try:
            content = blob.decode("utf-8-sig")
        except UnicodeDecodeError:
            problems.append((path, "non-UTF-8 content"))
            continue
        if size > 1_000_000:
            problems.append((path, "large file requires review"))
        for label, pattern in SENSITIVE_PATTERNS.items():
            if pattern.search(content):
                problems.append((path, label))
        if path.endswith(".py"):
            try:
                compile(content, path, "exec")
                python_count += 1
            except SyntaxError:
                problems.append((path, "Python syntax error"))
    # Exercise exclusions using paths that need not exist, including future data.
    private_probes = ["future.xlsx", "new_folder/private.py", "pharmacy/output/new.py",
                      "pharmacy/scripts/credentials.json", "visualizacion/app/src/data.json",
                      "visualizacion/app/dist/index.html", "visualizacion/tests/new.csv",
                      "visualizacion/app/src/.env", "loop/status.md"]
    for path in private_probes:
        code = subprocess.run([git, "check-ignore", "--quiet", "--no-index", path], cwd=ROOT).returncode
        if code != 0:
            problems.append((path, "exclusion probe failed"))
    if problems:
        for path, reason in problems:
            print(f"FAIL: {path}: {reason}")
        return 1
    print(f"OK: {len(entries)} Git blobs; {byte_count:,} bytes; {python_count} Python sources; "
          f"{len(private_probes)} exclusion probes; no detected secrets or private formats.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
