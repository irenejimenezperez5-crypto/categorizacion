"""Download and fingerprint the public, versioned inputs used by categorization.

Private CRM/IQVIA records are never sent to the download endpoints.
"""
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
from pathlib import Path
import re
import time
import zipfile

import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

ROOT = Path(__file__).resolve().parents[2]
CACHE = ROOT / "pharmacy/cache/categorization/official"
MANIFEST = CACHE / "manifest.json"
URLS = {
    "sections_2025.zip": "https://www.ine.es/prodyser/cartografia/seccionado_2025.zip",
    "degurba_2021.zip": "https://ec.europa.eu/eurostat/documents/7116161/15125332/DGURBA_RG_01M_2021_4258.zip/09a07de6-f16e-af72-1178-123259fadefd?t=1748954595744",
    "grid_2021_v3.zip": "https://gisco-services.ec.europa.eu/census/2021/Eurostat_Census-GRID_2021_V3.zip",
}
AGE_URL = "https://www.ine.es/jaxiT3/files/t/es/csv_bdsc/65034.csv"


def session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "PharmacyCategorization/2027.1 public-statistics", "Accept-Encoding": "identity"})
    s.mount("https://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=1.5, status_forcelist=[429, 500, 502, 503, 504])))
    return s


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(2**20), b""):
            h.update(block)
    return h.hexdigest()


def download(item: tuple[str, str]) -> dict:
    name, url = item
    target = CACHE / name
    if target.exists() and zipfile.is_zipfile(target):
        print(f"CACHE {name} {target.stat().st_size / 1e6:.1f} MB", flush=True)
        return {"file": name, "url": url, "sha256": sha256(target), "bytes": target.stat().st_size, "cached": True}
    part = target.with_suffix(target.suffix + ".part")
    with session() as s:
        with s.get(url, stream=True, timeout=(15, 90)) as response:
            response.raise_for_status()
            if response.status_code == 206:
                # A cached probe must not be mistaken for the complete archive.
                raise RuntimeError(f"Unexpected partial archive for {name}: {response.headers.get('Content-Range')}")
            digest = hashlib.sha256()
            count = 0
            last_report = time.monotonic()
            with part.open("wb") as out:
                for block in response.iter_content(2**20):
                    if not block:
                        continue
                    out.write(block)
                    digest.update(block)
                    count += len(block)
                    if time.monotonic() - last_report > 20:
                        print(f"DOWNLOAD {name}: {count / 1e6:.1f} MB", flush=True)
                        last_report = time.monotonic()
            expected = response.headers.get("Content-Length")
            if expected and count != int(expected):
                raise RuntimeError(f"Incomplete download {name}: {count}/{expected}")
            if not zipfile.is_zipfile(part):
                raise RuntimeError(f"Invalid ZIP {name}")
            part.replace(target)
            record = {"file": name, "url": url, "sha256": digest.hexdigest(), "bytes": count, "etag": response.headers.get("ETag"), "last_modified": response.headers.get("Last-Modified")}
    print(f"READY {name} {count / 1e6:.1f} MB", flush=True)
    return record


def filter_population(year: int = 2025) -> dict:
    target = CACHE / f"ine_ages_{year}_v2_filtered.csv.gz"
    if target.exists():
        print(f"CACHE {target.name}", flush=True)
        return {"file": target.name, "url": AGE_URL, "sha256": sha256(target), "filter": f"Periodo={year}; Sexo=Total; ages total,65-69,70-74,75-79,80-84,85+", "bytes": target.stat().st_size}
    selected = []
    consumed = 0
    with session() as s, s.get(AGE_URL, stream=True, timeout=(15, 120)) as response:
        response.raise_for_status()
        response.raw.decode_content = True
        for chunk in pd.read_csv(response.raw, sep=";", encoding="utf-8-sig", dtype=str, chunksize=200_000):
            consumed += len(chunk)
            if not {"Secciones", "Sexo", "Edad", "Periodo", "Total"} <= set(chunk):
                raise RuntimeError(f"Unexpected INE columns: {list(chunk)}")
            age = chunk["Edad"].fillna("")
            keep = (chunk["Periodo"] == str(year)) & (chunk["Sexo"] == "Total")
            # INE has separate 85-89, 90-94, 95-99 and 100+ groups.
            # 85+ MUST include all of them, not just the first quinquennium.
            keep &= age.str.contains(r"Todas las edades|\b65\b|\b70\b|\b75\b|\b80\b|\b85\b|\b90\b|\b95\b|\b100\b", regex=True)
            selected.append(chunk.loc[keep].copy())
            if consumed % 1_000_000 == 0:
                print(f"INE CSV parsed {consumed:,} rows", flush=True)
    data = pd.concat(selected, ignore_index=True)
    if data.empty or data["Secciones"].dropna().nunique() < 30_000:
        raise RuntimeError("INE 2025 section age coverage unexpectedly low")
    data.to_csv(target, index=False, encoding="utf-8", compression="gzip")
    print(f"READY {target.name}: {len(data):,} records from {consumed:,}", flush=True)
    return {"file": target.name, "url": AGE_URL, "sha256": sha256(target), "bytes": target.stat().st_size, "rows": len(data), "source_rows_parsed": consumed, "filter": f"Periodo={year}; Sexo=Total; ages total,65-69,70-74,75-79,80-84,85+", "source_hash_scope": "filtered reproducible extract; not the entire streamed CSV"}


def prepare() -> dict:
    CACHE.mkdir(parents=True, exist_ok=True)
    records = []
    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = [executor.submit(download, item) for item in URLS.items()]
        futures.append(executor.submit(filter_population))
        for future in futures:
            records.append(future.result())
    result = {"reference_date": "2026-10-06", "sources": records}
    MANIFEST.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    for name in URLS:
        with zipfile.ZipFile(CACHE / name) as z:
            print(json.dumps({"archive": name, "members": [{"name": i.filename, "bytes": i.file_size} for i in z.infolist() if not i.is_dir()][:70]}, ensure_ascii=True), flush=True)
    return result


if __name__ == "__main__":
    prepare()
