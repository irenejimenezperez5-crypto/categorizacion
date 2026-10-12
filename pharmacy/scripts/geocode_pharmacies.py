import argparse
import csv
import hashlib
import json
import math
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
import threading

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


PHARMACY_DIR = Path(__file__).resolve().parents[1]
SOURCE_DIR = PHARMACY_DIR
OUTPUT_DIR = PHARMACY_DIR / "output"
CACHE_DIR = PHARMACY_DIR / "cache"
CHECKPOINT_DIR = CACHE_DIR / "checkpoints"

VALID_THRESHOLD = 0.95
WARNING_THRESHOLD = 0.90
MAX_RETRIES = 2
TIMEOUT = 8
REQUEST_DELAY = 0.35
MAX_WORKERS = 6
GEOCODING_BLOCK_SIZE = 100
CACHE_PATH = CACHE_DIR / "geocoding_cache_cartociudad_v7.csv"
OUTPUT_PATH = OUTPUT_DIR / "PHARMACY_MASTER_GEOCODED.xlsx"
LOG_PATH = OUTPUT_DIR / "geocoding_run_log.txt"
PROFILE_PATH = OUTPUT_DIR / "source_profile.json"
NORMALIZED_CHECKPOINT_PATH = CHECKPOINT_DIR / "normalized_pharmacies.csv"
GEOCODED_CHECKPOINT_PATH = CHECKPOINT_DIR / "geocoded_pharmacies_cartociudad_v7.csv"

CARTOCIUDAD_BASE_URL = "https://www.cartociudad.es/geocoder/api/geocoder"
CARTOCIUDAD_CANDIDATES_URL = f"{CARTOCIUDAD_BASE_URL}/candidates"

SOURCE_PATTERNS = ("universo", "pharmacy", "farmacia", "pharmacies")

EXPECTED_SHEETS = ("PHARMACY_MASTER", "GEOCODING_REVIEW", "GEOCODING_QA")
CACHE_LOCK = threading.Lock()
STATS_LOCK = threading.Lock()


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def ensure_dirs():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    CHECKPOINT_DIR.mkdir(parents=True, exist_ok=True)


def log(message):
    ensure_dirs()
    line = f"[{now_iso()}] {message}"
    print(line, flush=True)
    with LOG_PATH.open("a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def clean_text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value)
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s*,\s*", ", ", text)
    return text.strip(" ,")


def repair_mojibake(value):
    text = clean_text(value)
    if not text or not any(mark in text for mark in ("Ã", "Â", "â")):
        return text
    try:
        repaired = text.encode("latin1").decode("utf-8")
        if fold_text(repaired).replace(" ", ""):
            return clean_text(repaired)
    except Exception:
        return text
    return text


def fold_text(value):
    text = clean_text(value).casefold()
    text = "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^0-9a-zñç ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_postal_code(value):
    text = clean_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d+(?:\.0)?", text):
        text = text.split(".")[0]
    digits = re.sub(r"\D", "", text)
    if len(digits) == 4:
        return "0" + digits
    if len(digits) >= 5:
        return digits[:5]
    return digits


def first_non_empty(row, columns):
    for col in columns:
        if col and col in row.index:
            value = clean_text(row[col])
            if value:
                return value
    return ""


def discover_source_file():
    candidates = []
    for path in SOURCE_DIR.iterdir():
        if not path.is_file() or path.suffix.lower() not in (".xlsx", ".xls"):
            continue
        name_fold = fold_text(path.name)
        if any(p in name_fold for p in SOURCE_PATTERNS):
            stat = path.stat()
            try:
                xls = pd.ExcelFile(path)
                sheets = []
                for sheet in xls.sheet_names:
                    preview = pd.read_excel(path, sheet_name=sheet, dtype=str, nrows=1)
                    rows = pd.read_excel(path, sheet_name=sheet, dtype=str, usecols=[0]).shape[0]
                    sheets.append({"sheet": sheet, "rows": int(rows), "columns": int(preview.shape[1])})
            except Exception as exc:
                sheets = [{"error": str(exc)}]
            candidates.append({
                "name": path.name,
                "relative_path": str(path.relative_to(PHARMACY_DIR.parent)),
                "path": str(path),
                "size_bytes": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                "sheets": sheets,
            })

    if not candidates:
        raise RuntimeError("STOPPER: no Excel source candidate found in pharmacy/.")
    candidates = sorted(candidates, key=lambda c: ("universo_pharmacy" not in fold_text(c["name"]), c["name"].lower()))
    if len(candidates) > 1:
        best = candidates[0]
        if "universo_pharmacy" not in fold_text(best["name"]):
            raise RuntimeError("STOPPER: multiple candidates and no unique universo_pharmacy priority.")
    return Path(candidates[0]["path"]), candidates


def profile_workbook(path):
    profile = {"source_file": str(path), "sheets": []}
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False)
        sheet_info = {
            "sheet": sheet,
            "rows": int(df.shape[0]),
            "columns": int(df.shape[1]),
            "headers": [str(c) for c in df.columns],
            "first_10_rows": df.head(10).to_dict(orient="records"),
            "null_percent": {},
            "categorical_uniques": {},
            "possible_ids": [],
            "possible_address_fields": [],
        }
        for col in df.columns:
            series = df[col].map(clean_text)
            null_pct = float((series == "").mean() * 100) if len(series) else 0.0
            sheet_info["null_percent"][str(col)] = round(null_pct, 2)
            unique_count = series[series != ""].nunique(dropna=True)
            col_fold = fold_text(col)
            if unique_count and unique_count <= 25:
                sheet_info["categorical_uniques"][str(col)] = series[series != ""].drop_duplicates().head(25).tolist()
            if any(tok in col_fold for tok in ("id", "account", "crm", "codigo", "code", "number")):
                sheet_info["possible_ids"].append({"column": str(col), "unique": int(unique_count), "null_pct": round(null_pct, 2)})
            if any(tok in col_fold for tok in ("address", "street", "calle", "direccion", "postal", "zip", "municip", "province", "provincia", "localidad", "poblacion", "city")):
                sheet_info["possible_address_fields"].append(str(col))
        profile["sheets"].append(sheet_info)
    PROFILE_PATH.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")
    return profile


def choose_source_sheet(profile):
    scored = []
    for sheet in profile["sheets"]:
        headers_fold = " ".join(fold_text(h) for h in sheet["headers"])
        score = sheet["rows"]
        for token in ("pharmacy", "farmacia", "account", "crm", "direccion", "address", "postal", "sb"):
            if token in headers_fold:
                score += 1000
        scored.append((score, sheet))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1]["sheet"]


def pick_column(columns, patterns, required=False):
    ranked = []
    for col in columns:
        f = fold_text(col)
        for rank, pattern in enumerate(patterns):
            if re.search(pattern, f):
                ranked.append((rank, len(f), col))
                break
    if not ranked:
        if required:
            raise RuntimeError(f"STOPPER: required column not found for patterns {patterns}")
        return None
    ranked.sort()
    return ranked[0][2]


def infer_column_mapping(df):
    columns = list(df.columns)
    mapping = {
        "pharmacy_id": pick_column(columns, [r"^account id$", r"crm.*id", r"account.*id", r"pharmacy.*id", r"farmacia.*id", r"^id$", r"account number", r"customer.*id"]),
        "pharmacy_name": pick_column(columns, [r"pharmacy.*name", r"nombre.*farm", r"farmacia", r"account name", r"name", r"razon social"]),
        "sb_code_2026": pick_column(columns, [r"sb.*2026", r"sb.*code", r"codigo.*sb", r"code.*sb", r"^sb$"]),
        "address_raw": pick_column(columns, [r"full.*address", r"direccion completa", r"address", r"direccion", r"street 1"]),
        "street_type_raw": pick_column(columns, [r"tipo.*via", r"street.*type", r"tip_via"]),
        "street_name_raw": pick_column(columns, [r"street.*name", r"nombre.*via", r"calle", r"via", r"street"]),
        "street_number_raw": pick_column(columns, [r"numero", r"number", r"portal", r"num"]),
        "address_extra_raw": pick_column(columns, [r"address.*2", r"street 2", r"extra", r"complement", r"piso", r"local"]),
        "postal_code_raw": pick_column(columns, [r"postal", r"zip", r"codigo postal", r"cp"]),
        "municipality_raw": pick_column(columns, [r"municip", r"localidad", r"poblacion", r"city", r"ciudad"]),
        "province_raw": pick_column(columns, [r"provincia", r"province"]),
        "autonomous_region_raw": pick_column(columns, [r"ccaa", r"ccaaa", r"autonom", r"region", r"comunidad"]),
        "country_raw": pick_column(columns, [r"country", r"pais"]),
    }
    return mapping


def validate_source_schema(df, mapping):
    stoppers = []
    if not mapping.get("pharmacy_id"):
        stoppers.append("No stable pharmacy_id candidate identified.")
    if not mapping.get("sb_code_2026"):
        stoppers.append("No SB CODE candidate identified.")
    address_cols = [mapping.get(k) for k in ("address_raw", "street_name_raw", "municipality_raw", "postal_code_raw", "province_raw")]
    if not any(address_cols):
        stoppers.append("No address fields identified.")
    if stoppers:
        raise RuntimeError("STOPPER: " + " | ".join(stoppers))

    pid = mapping["pharmacy_id"]
    ids = df[pid].map(clean_text)
    null_ids = int((ids == "").sum())
    dup_ids = int(ids[ids != ""].duplicated().sum())
    if null_ids:
        raise RuntimeError(f"STOPPER: pharmacy_id has {null_ids} null values.")
    if dup_ids:
        dup_examples = ids[ids.duplicated(keep=False)].head(10).tolist()
        raise RuntimeError(f"STOPPER: pharmacy_id has {dup_ids} duplicate rows. Examples: {dup_examples}")

    missing_core = 0
    examples = []
    for _, row in df.iterrows():
        street = first_non_empty(row, [mapping.get("address_raw"), mapping.get("street_name_raw")])
        muni = first_non_empty(row, [mapping.get("municipality_raw")])
        if not street or not muni:
            missing_core += 1
            if len(examples) < 5:
                examples.append(row.to_dict())
    if len(df) and missing_core / len(df) > 0.10:
        raise RuntimeError(f"STOPPER: {missing_core} rows ({missing_core/len(df):.1%}) lack street/address or municipality. Examples: {examples}")


def normalize_street_number(value):
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(r"\s+", "", text.upper())
    return text


def normalize_country(value):
    text = clean_text(value)
    if not text:
        return "España"
    if fold_text(text) in ("espana", "españa", "spain", "es", "011"):
        return "España"
    return text


PROVINCE_FIXES = {
    "alava": "Álava",
    "avila": "Ávila",
    "cadiz": "Cádiz",
    "castellon": "Castellón",
    "cordoba": "Córdoba",
    "coruna": "A Coruña",
    "a coruna": "A Coruña",
    "gerona": "Girona",
    "guipuzcoa": "Gipuzkoa",
    "lerida": "Lleida",
    "malaga": "Málaga",
    "murcia": "Murcia",
    "navarra": "Navarra",
    "orense": "Ourense",
    "rioja": "La Rioja",
    "sevilla": "Sevilla",
    "valencia": "Valencia",
    "vizcaya": "Bizkaia",
}


STREET_PREFIXES = [
    (r"^(c/|cl/|cl\.|calle)\s+", "Calle"),
    (r"^(avda/|avda\.|av/|av\.|avenida)\s+", "Avenida"),
    (r"^(pza/|pza\.|plz/|plz\.|plaza)\s+", "Plaza"),
    (r"^(ps/|pº|paseo)\s+", "Paseo"),
    (r"^(ctra/|ctra\.|carretera)\s+", "Carretera"),
    (r"^(ronda)\s+", "Ronda"),
    (r"^(camino)\s+", "Camino"),
    (r"^(travesia|trva/|trva\.)\s+", "Travesía"),
]


def normalize_province(value):
    text = repair_mojibake(value)
    return PROVINCE_FIXES.get(fold_text(text), text)


def normalize_place(value):
    return repair_mojibake(value)


def split_street_type(street_type, street_name):
    stype = repair_mojibake(street_type)
    sname = repair_mojibake(street_name)
    if stype or not sname:
        return stype, sname
    for pattern, replacement in STREET_PREFIXES:
        if re.search(pattern, sname, flags=re.IGNORECASE):
            cleaned = re.sub(pattern, "", sname, count=1, flags=re.IGNORECASE).strip()
            return replacement, cleaned
    return "", sname


def build_full_address(row):
    street_type = row.get("street_type_normalized", "")
    street_name = row.get("street_name_normalized", "")
    street_number = row.get("street_number_normalized", "")
    address = row.get("address_raw", "")
    parts = []
    if street_name:
        first = " ".join(x for x in (street_type, street_name, street_number) if x)
        parts.append(first)
    elif address:
        parts.append(address)
    locality = " ".join(x for x in (row.get("postal_code_normalized", ""), row.get("municipality_normalized", "")) if x)
    if locality:
        parts.append(locality)
    for field in ("province_normalized", "country_normalized"):
        if row.get(field):
            parts.append(row[field])
    return ", ".join(parts)


def normalize_address_components(df, mapping):
    out = df.copy()
    for std, src in mapping.items():
        if src and src in out.columns:
            out[std] = out[src].map(clean_text)
        else:
            out[std] = ""

    if not out["street_name_raw"].any() and out["address_raw"].any():
        out["street_name_raw"] = out["address_raw"]
    split = out.apply(lambda r: split_street_type(r["street_type_raw"], r["street_name_raw"]), axis=1)
    out["street_type_normalized"] = [x[0] for x in split]
    out["street_name_normalized"] = [x[1] for x in split]
    out["street_number_normalized"] = out["street_number_raw"].map(normalize_street_number)
    out["postal_code_normalized"] = out["postal_code_raw"].map(normalize_postal_code)
    out["municipality_normalized"] = out["municipality_raw"].map(normalize_place)
    out["province_normalized"] = out["province_raw"].map(normalize_province)
    out["autonomous_region_normalized"] = out["autonomous_region_raw"].map(normalize_place)
    out["country_normalized"] = out["country_raw"].map(normalize_country)
    out["full_address_normalized"] = out.apply(build_full_address, axis=1)
    out["full_address_compare"] = out["full_address_normalized"].map(fold_text)
    out.to_csv(NORMALIZED_CHECKPOINT_PATH, index=False, encoding="utf-8-sig")
    return out


def load_geocoding_cache():
    cache = {}
    if CACHE_PATH.exists():
        with CACHE_PATH.open("r", encoding="utf-8-sig", newline="") as fh:
            for row in csv.DictReader(fh):
                cache[row["cache_key"]] = row
    return cache


def append_cache(row):
    with CACHE_LOCK:
        exists = CACHE_PATH.exists()
        fieldnames = list(row.keys())
        with CACHE_PATH.open("a", encoding="utf-8-sig", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames)
            if not exists:
                writer.writeheader()
            writer.writerow(row)


def cache_key(query):
    return hashlib.sha256(fold_text(query).encode("utf-8")).hexdigest()


def parse_lat_lng(candidate):
    lat = candidate.get("lat")
    lng = candidate.get("lng")
    try:
        return float(lat), float(lng)
    except Exception:
        return None, None


def is_in_spain(lat, lon):
    if lat is None or lon is None:
        return False
    boxes = [
        (35.8, 44.3, -9.5, 4.5),      # Peninsula + Baleares
        (27.4, 29.5, -18.5, -13.0),   # Canarias
        (35.7, 35.95, -5.45, -5.20),  # Ceuta
        (35.15, 35.40, -3.10, -2.85), # Melilla
    ]
    return any(a <= lat <= b and c <= lon <= d for a, b, c, d in boxes)


def validate_coordinates(lat, lon):
    try:
        lat = float(lat)
        lon = float(lon)
    except (TypeError, ValueError):
        return False
    if not math.isfinite(lat) or not math.isfinite(lon) or (lat == 0 and lon == 0):
        return False
    return is_in_spain(lat, lon)


def value_match(input_value, result_value):
    a = fold_text(input_value)
    b = fold_text(result_value)
    if not a or not b:
        return ""
    if a == b or a in b or b in a:
        return "MATCH"
    a_tokens = sorted(set(a.split()))
    b_tokens = sorted(set(b.split()))
    if a_tokens == b_tokens:
        return "MATCH"
    if len(a_tokens) >= 2 and set(a_tokens).issubset(set(b_tokens)):
        return "MATCH"
    if len(b_tokens) >= 2 and set(b_tokens).issubset(set(a_tokens)):
        return "MATCH"
    return "MISMATCH"


def number_match(input_number, result_number):
    a = fold_text(input_number).replace(" ", "")
    b = fold_text(result_number).replace(" ", "")
    if not a:
        return ""
    if not b:
        return "MISSING_RESULT"
    return "MATCH" if a == b else "MISMATCH"


def build_attempt_queries(row):
    st = row.get("street_type_normalized", "")
    sn = row.get("street_name_normalized", "") or row.get("address_raw", "")
    num = row.get("street_number_normalized", "")
    cp = row.get("postal_code_normalized", "")
    muni = row.get("municipality_normalized", "")
    prov = row.get("province_normalized", "")
    street_num = " ".join(x for x in (st, sn, num) if x)
    street_num_no_type = " ".join(x for x in (sn, num) if x)
    street_only = " ".join(x for x in (st, sn) if x)
    street_only_no_type = sn
    muni_compact = re.sub(r"\s*-\s*", "-", muni)
    muni_spaced = re.sub(r"\s*-\s*", " ", muni)
    # CartoCiudad's text parser returns better portal candidates with street+number+municipality.
    # CP/province remain mandatory validation fields after the API returns candidates.
    return [
        ("ATTEMPT_1", " ".join(x for x in (street_num_no_type, muni_compact) if x)),
        ("ATTEMPT_2", " ".join(x for x in (street_num, muni_compact) if x)),
        ("ATTEMPT_3", " ".join(x for x in (street_num, cp) if x)),
        ("ATTEMPT_4_DIAGNOSTIC", " ".join(x for x in (street_only_no_type, muni_compact) if x)),
    ]


def request_candidates(query, stats=None):
    params = {"q": query, "limit": 10}
    last_error = ""
    for attempt in range(MAX_RETRIES + 1):
        if stats is not None:
            with STATS_LOCK:
                stats["http_attempts"] += 1
        try:
            response = requests.get(CARTOCIUDAD_CANDIDATES_URL, params=params, timeout=TIMEOUT)
            if response.status_code == 429 or response.status_code >= 500:
                last_error = f"HTTP {response.status_code}"
                with STATS_LOCK:
                    stats["retries"] += 1
                    if response.status_code == 429:
                        stats["http_429"] += 1
                    elif response.status_code >= 500:
                        stats["http_5xx"] += 1
                if attempt < MAX_RETRIES:
                    time.sleep((2 ** attempt) + REQUEST_DELAY)
                continue
            response.raise_for_status()
            data = response.json()
            if isinstance(data, dict) and "response" in data:
                return data.get("response") or [], ""
            if isinstance(data, dict) and "value" in data:
                return data.get("value") or [], ""
            if isinstance(data, list):
                return data, ""
            return [], "Unexpected JSON shape"
        except Exception as exc:
            last_error = str(exc)
            with STATS_LOCK:
                stats["retries"] += 1
            if attempt < MAX_RETRIES:
                time.sleep((2 ** attempt) + REQUEST_DELAY)
    if stats is not None and last_error:
        with STATS_LOCK:
            stats["errors"] += 1
    return [], last_error


def choose_best_candidate(candidates, row):
    if not candidates:
        return None
    preferred = []
    for cand in candidates:
        score = 0
        ctype = fold_text(cand.get("type", ""))
        if ctype == "portal":
            score += 30
        if value_match(row.get("municipality_normalized", ""), repair_mojibake(cand.get("muni", ""))) == "MATCH":
            score += 20
        if value_match(row.get("province_normalized", ""), repair_mojibake(cand.get("province", ""))) == "MATCH":
            score += 15
        if row.get("postal_code_normalized") and clean_text(cand.get("postalCode")) == row.get("postal_code_normalized"):
            score += 15
        if number_match(row.get("street_number_normalized", ""), cand.get("portalNumber", "")) == "MATCH":
            score += 20
        preferred.append((score, cand))
    preferred.sort(key=lambda x: x[0], reverse=True)
    return preferred[0][1]


def score_geocoding_result(row, cand, attempt_name, query, error=""):
    base = {
        "latitude": "",
        "longitude": "",
        "geocoding_source": "CartoCiudad REST Geocoder",
        "geocoding_attempt": attempt_name,
        "geocoding_query": query,
        "geocoding_status": "FAILED" if error else "NO_RESULT",
        "geocoding_quality": "FAILED",
        "geocoding_score": 0,
        "matched_address": "",
        "matched_street": "",
        "matched_number": "",
        "matched_postal_code": "",
        "matched_municipality": "",
        "matched_province": "",
        "address_match": "",
        "number_match": "",
        "postal_code_match": "",
        "municipality_match": "",
        "province_match": "",
        "geocoding_notes": error or "No result",
        "geocoding_timestamp": now_iso(),
        "reason_for_review": error or "No result",
        "suggested_action": "Review source address and geocode manually",
    }
    if not cand:
        return base

    lat, lon = parse_lat_lng(cand)
    ctype = clean_text(cand.get("type"))
    cand_address = repair_mojibake(cand.get("address", ""))
    cand_muni = repair_mojibake(cand.get("muni", ""))
    cand_province = repair_mojibake(cand.get("province", ""))
    street_match = value_match(row.get("street_name_normalized", ""), cand_address)
    num_match = number_match(row.get("street_number_normalized", ""), cand.get("portalNumber", ""))
    cp_match = "MATCH" if row.get("postal_code_normalized") and clean_text(cand.get("postalCode")) == row.get("postal_code_normalized") else ("MISMATCH" if row.get("postal_code_normalized") and clean_text(cand.get("postalCode")) else "")
    muni_match = value_match(row.get("municipality_normalized", ""), cand_muni)
    prov_match = value_match(row.get("province_normalized", ""), cand_province)
    in_spain = validate_coordinates(lat, lon)

    score = 0
    score += 20 if street_match == "MATCH" else 0
    score += 25 if num_match == "MATCH" else 0
    score += 15 if cp_match == "MATCH" else 0
    score += 20 if muni_match == "MATCH" else 0
    score += 10 if prov_match == "MATCH" else 0
    score += 10 if fold_text(ctype) == "portal" else 0
    if not in_spain:
        score = min(score, 40)
    if attempt_name == "ATTEMPT_4_DIAGNOSTIC":
        score = min(score, 69)
    if num_match != "MATCH":
        score = min(score, 79)
    if muni_match == "MISMATCH":
        score = min(score, 49)
    if cp_match == "MISMATCH":
        score = min(score, 79)
    if fold_text(ctype) not in ("portal",):
        score = min(score, 79)

    if score >= 90 and in_spain and num_match == "MATCH" and muni_match == "MATCH" and fold_text(ctype) == "portal":
        quality = "EXACT"
    elif score >= 80 and in_spain and muni_match == "MATCH" and fold_text(ctype) == "portal":
        quality = "HIGH"
    elif score >= 60 and in_spain:
        quality = "REVIEW"
    else:
        quality = "FAILED"

    reasons = []
    if num_match != "MATCH":
        reasons.append("number not confirmed")
    if cp_match == "MISMATCH":
        reasons.append("postcode mismatch")
    if muni_match == "MISMATCH":
        reasons.append("municipality mismatch")
    if prov_match == "MISMATCH":
        reasons.append("province mismatch")
    if not in_spain:
        reasons.append("coordinates outside Spain bounding boxes")
    if fold_text(ctype) != "portal":
        reasons.append(f"non-portal match: {ctype}")
    if attempt_name == "ATTEMPT_4_DIAGNOSTIC":
        reasons.append("diagnostic street-level attempt only")

    base.update({
        "latitude": lat if lat is not None and in_spain else "",
        "longitude": lon if lon is not None and in_spain else "",
        "geocoding_status": "OK" if quality in ("EXACT", "HIGH", "REVIEW") else "FAILED",
        "geocoding_quality": quality,
        "geocoding_score": int(score),
        "matched_address": cand_address,
        "matched_street": cand_address,
        "matched_number": clean_text(cand.get("portalNumber")),
        "matched_postal_code": clean_text(cand.get("postalCode")),
        "matched_municipality": cand_muni,
        "matched_province": cand_province,
        "address_match": street_match,
        "number_match": num_match,
        "postal_code_match": cp_match,
        "municipality_match": muni_match,
        "province_match": prov_match,
        "geocoding_notes": "; ".join(reasons) if reasons else "Accepted portal-level match",
        "reason_for_review": "; ".join(reasons),
        "suggested_action": "" if quality in ("EXACT", "HIGH") else "Review address and confirm portal-level coordinates manually",
    })
    return base


def geocode_address(row, cache, stats):
    for attempt_name, query in build_attempt_queries(row):
        if not query:
            continue
        key = cache_key(query)
        with CACHE_LOCK:
            cached_row = cache.get(key)
        if cached_row:
            with STATS_LOCK:
                stats["cache_hits"] += 1
            cached = dict(cache[key])
            cached.pop("cache_key", None)
            if cached.get("geocoding_quality") in ("EXACT", "HIGH", "REVIEW") or attempt_name == "ATTEMPT_4_DIAGNOSTIC":
                return cached
            continue
        with STATS_LOCK:
            stats["requests_sent"] += 1
        candidates, error = request_candidates(query, stats)
        best = choose_best_candidate(candidates, row)
        result = score_geocoding_result(row, best, attempt_name, query, error)
        result["raw_response"] = json.dumps(candidates[:3], ensure_ascii=False)
        cache_row = {"cache_key": key, **result}
        append_cache(cache_row)
        with CACHE_LOCK:
            cache[key] = cache_row
        time.sleep(REQUEST_DELAY)
        if result["geocoding_quality"] in ("EXACT", "HIGH"):
            return result
        if result["geocoding_quality"] == "REVIEW" and attempt_name != "ATTEMPT_4_DIAGNOSTIC":
            return result
    return score_geocoding_result(row, None, "", "", "No usable geocoding result")


def geocode_dataframe(df):
    cache = load_geocoding_cache()
    stats = Counter()
    stats["cache_entries_loaded"] = len(cache)
    completed = {}
    if GEOCODED_CHECKPOINT_PATH.exists():
        cp = pd.read_csv(GEOCODED_CHECKPOINT_PATH, dtype=str, keep_default_na=False)
        if "pharmacy_id" in cp.columns:
            completed = {clean_text(r["pharmacy_id"]): r for _, r in cp.iterrows()}
            log(f"Loaded geocoded checkpoint rows={len(completed)}")

    result_by_idx = {}
    futures = {}

    def work(item):
        idx, row = item
        row_dict = row.to_dict()
        pid = clean_text(row_dict["pharmacy_id"])
        if pid in completed:
            with STATS_LOCK:
                stats["checkpoint_hits"] += 1
            return idx, {**row_dict, **completed[pid].to_dict()}, False
        result = geocode_address(row_dict, cache, stats)
        return idx, {**row_dict, **result}, True

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        for item in df.iterrows():
            futures[executor.submit(work, item)] = item[0]
        completed_count = 0
        new_results_count = 0
        for future in as_completed(futures):
            idx, combined, is_new = future.result()
            result_by_idx[idx] = combined
            completed_count += 1
            new_results_count += int(is_new)
            if completed_count % GEOCODING_BLOCK_SIZE == 0 and new_results_count:
                ordered = [result_by_idx[i] for i in sorted(result_by_idx)]
                pd.DataFrame(ordered).to_csv(GEOCODED_CHECKPOINT_PATH, index=False, encoding="utf-8-sig")
                log(f"Geocoding checkpoint saved rows={len(ordered)} requests={stats['requests_sent']} cache_hits={stats['cache_hits']}")
    out = pd.DataFrame([result_by_idx[i] for i in sorted(result_by_idx)])
    if new_results_count:
        out.to_csv(GEOCODED_CHECKPOINT_PATH, index=False, encoding="utf-8-sig")
    return out, stats


def detect_suspicious_coordinates(df):
    warnings = []
    coord = df[["latitude", "longitude"]].fillna("").astype(str)
    coordinate_pairs = (
        (row["latitude"], row["longitude"])
        for _, row in coord.iterrows()
        if row["latitude"]
        and row["longitude"]
        and not (row["latitude"] in {"0", "0.0"} and row["longitude"] in {"0", "0.0"})
    )
    counts = Counter(coordinate_pairs)
    suspicious = {k: v for k, v in counts.items() if v >= 5}
    rounded = 0
    outside = 0
    missing = 0
    for _, row in df.iterrows():
        try:
            lat = float(row.get("latitude", ""))
            lon = float(row.get("longitude", ""))
            if not validate_coordinates(lat, lon):
                outside += 1
            if len(str(row.get("latitude")).split(".")[-1]) <= 3 or len(str(row.get("longitude")).split(".")[-1]) <= 3:
                rounded += 1
        except Exception:
            missing += 1
    if suspicious:
        warnings.append(f"Suspicious duplicate coordinates >=5 pharmacies: {len(suspicious)} coordinate pairs")
    if rounded:
        warnings.append(f"{rounded} coordinates appear excessively rounded")
    return {
        "missing_coordinates": missing,
        "outside_spain": outside,
        "suspicious_duplicate_coordinate_pairs": len(suspicious),
        "suspicious_duplicate_rows": sum(suspicious.values()),
        "rounded_coordinates": rounded,
        "warnings": warnings,
    }


def build_qa_report(df, source_file, source_sheet, mapping, stats, source_rows):
    total = len(df)
    quality_counts = df["geocoding_quality"].value_counts(dropna=False).to_dict()
    geo_q = detect_suspicious_coordinates(df)
    valid = quality_counts.get("EXACT", 0) + quality_counts.get("HIGH", 0)
    valid_pct = valid / total if total else 0
    municipality_mismatches = int((df["municipality_match"] == "MISMATCH").sum())
    province_mismatches = int((df["province_match"] == "MISMATCH").sum())
    postcode_mismatches = int((df["postal_code_match"] == "MISMATCH").sum())
    street_present = df["street_name_normalized"].map(clean_text) != ""
    municipality_present = df["municipality_normalized"].map(clean_text) != ""
    addresses_complete = int((street_present & municipality_present).sum())
    addresses_incomplete = total - addresses_complete
    rejected_invalid_or_outside = int(
        df["geocoding_notes"].map(clean_text).str.contains(
            "coordinates outside Spain bounding boxes", regex=False
        ).sum()
    )
    run_status = "PASS"
    stopper_events = []
    if valid_pct < WARNING_THRESHOLD:
        run_status = "STOPPED_NOT_APPROVED"
        stopper_events.append(f"VALID below 90% ({valid_pct:.2%})")
    elif valid_pct < VALID_THRESHOLD:
        run_status = "WARNING"
    if total and municipality_mismatches / total > 0.05:
        run_status = "STOPPED_NOT_APPROVED"
        stopper_events.append("municipality mismatches above 5%")
    if total and geo_q["outside_spain"] / total > 0.02:
        run_status = "STOPPED_NOT_APPROVED"
        stopper_events.append("outside-Spain coordinates above 2%")
    if geo_q["suspicious_duplicate_coordinate_pairs"]:
        run_status = "STOPPED_NOT_APPROVED"
        stopper_events.append("material duplicate coordinate pattern detected")
    if addresses_incomplete:
        geo_q["warnings"].append(
            f"{addresses_incomplete} rows lack both street and municipality completeness"
        )
    if rejected_invalid_or_outside:
        geo_q["warnings"].append(
            f"{rejected_invalid_or_outside} returned invalid/outside coordinate candidates were rejected"
        )

    rows = [
        ("Source file", str(source_file)),
        ("Source sheet", source_sheet),
        ("Total source rows", source_rows),
        ("Total pharmacies", total),
        ("Unique pharmacy IDs", df["pharmacy_id"].nunique()),
        ("Duplicate pharmacy IDs", int(df["pharmacy_id"].duplicated().sum())),
        ("Addresses complete", addresses_complete),
        ("Addresses incomplete", addresses_incomplete),
    ]
    for q in ("EXACT", "HIGH", "REVIEW", "FAILED"):
        n = int(quality_counts.get(q, 0))
        rows.append((f"{q} count", n))
        rows.append((f"{q} %", n / total if total else 0))
    rows.extend([
        ("VALID %", valid_pct),
        ("Missing coordinates", geo_q["missing_coordinates"]),
        ("Coordinates outside Spain", geo_q["outside_spain"]),
        ("Rejected invalid/outside coordinate candidates", rejected_invalid_or_outside),
        ("Municipality mismatches", municipality_mismatches),
        ("Province mismatches", province_mismatches),
        ("Postal code mismatches", postcode_mismatches),
        ("Suspicious duplicate coordinates", geo_q["suspicious_duplicate_rows"]),
        ("Requests sent", stats.get("requests_sent", 0)),
        ("Cache hits", stats.get("cache_hits", 0)),
        ("Checkpoint hits", stats.get("checkpoint_hits", 0)),
        ("Run status", run_status),
        ("Stopper events", "; ".join(stopper_events)),
        ("Column mapping", json.dumps(mapping, ensure_ascii=False)),
        ("Scoring methodology", "100-point score: street 20, portal number 25, postcode 15, municipality 20, province 10, portal geometry 10; caps applied for non-portal, missing/mismatched number, municipality mismatch, outside Spain, and diagnostic attempts."),
        ("Warnings", "; ".join(geo_q["warnings"])),
    ])

    examples = df[df["geocoding_quality"].isin(["REVIEW", "FAILED"])].head(15)
    qa = pd.DataFrame(rows, columns=["Metric", "Value"])
    if not examples.empty:
        err = examples[["pharmacy_id", "pharmacy_name", "full_address_normalized", "geocoding_quality", "geocoding_score", "geocoding_notes"]].copy()
        err.insert(0, "Metric", "Error example")
        err = err.rename(columns={"pharmacy_id": "Value"})
        qa = pd.concat([qa, err], ignore_index=True)
    return qa, run_status


def export_results(df, qa):
    export_df = df.copy()
    for col in ("latitude", "longitude"):
        export_df[col] = pd.to_numeric(export_df[col], errors="coerce")
    export_df["geocoding_score"] = pd.to_numeric(export_df["geocoding_score"], errors="coerce").astype("Int64")
    review_cols = [
        "pharmacy_id", "pharmacy_name", "address_raw", "full_address_normalized", "latitude", "longitude",
        "matched_address", "matched_postal_code", "matched_municipality", "matched_province",
        "geocoding_quality", "geocoding_score", "reason_for_review", "suggested_action",
    ]
    review = export_df[export_df["geocoding_quality"].isin(["REVIEW", "FAILED"])].copy()
    for col in review_cols:
        if col not in review.columns:
            review[col] = ""
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="PHARMACY_MASTER", index=False)
        review[review_cols].to_excel(writer, sheet_name="GEOCODING_REVIEW", index=False)
        qa.to_excel(writer, sheet_name="GEOCODING_QA", index=False)
    wb = load_workbook(OUTPUT_PATH)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        max_col = min(ws.max_column, 80)
        for i in range(1, max_col + 1):
            letter = get_column_letter(i)
            width = min(max(len(clean_text(ws.cell(row=r, column=i).value)) for r in range(1, min(ws.max_row, 100) + 1)) + 2, 45)
            ws.column_dimensions[letter].width = max(width, 10)
    wb.save(OUTPUT_PATH)


def validate_output(source_df, mapping):
    wb = load_workbook(OUTPUT_PATH, read_only=True, data_only=True)
    sheets = wb.sheetnames
    missing = [s for s in EXPECTED_SHEETS if s not in sheets]
    if missing:
        raise RuntimeError(f"Output workbook missing sheets: {missing}")
    ws = wb["PHARMACY_MASTER"]
    if ws.max_row - 1 != len(source_df):
        raise RuntimeError(f"Output row count mismatch: {ws.max_row - 1} vs {len(source_df)}")
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    missing_source_columns = [col for col in source_df.columns if col not in headers]
    if missing_source_columns:
        raise RuntimeError(f"Output missing original source columns: {missing_source_columns}")
    for col in ("pharmacy_id", "sb_code_2026", "latitude", "longitude", "postal_code_normalized"):
        if col not in headers:
            raise RuntimeError(f"Output missing required column {col}")

    header_index = {name: idx for idx, name in enumerate(headers)}
    source_rows = source_df.to_dict(orient="records")
    source_id_col = mapping["pharmacy_id"]
    source_sb_col = mapping["sb_code_2026"]
    original_mismatches = []
    invalid_coordinate_rows = []
    invalid_score_rows = []
    postal_format_errors = []
    for row_number, (source_row, output_row) in enumerate(
        zip(source_rows, ws.iter_rows(min_row=2, values_only=True)), start=2
    ):
        for col in source_df.columns:
            if clean_text(source_row.get(col)) != clean_text(output_row[header_index[col]]):
                original_mismatches.append((row_number, col))
                break
        if clean_text(source_row.get(source_id_col)) != clean_text(output_row[header_index["pharmacy_id"]]):
            raise RuntimeError(f"pharmacy_id not preserved at output row {row_number}")
        if clean_text(source_row.get(source_sb_col)) != clean_text(output_row[header_index["sb_code_2026"]]):
            raise RuntimeError(f"sb_code_2026 not preserved at output row {row_number}")

        lat = output_row[header_index["latitude"]]
        lon = output_row[header_index["longitude"]]
        if lat not in (None, "") or lon not in (None, ""):
            if not isinstance(lat, (int, float)) or not isinstance(lon, (int, float)):
                invalid_coordinate_rows.append(row_number)
            elif not math.isfinite(float(lat)) or not math.isfinite(float(lon)):
                invalid_coordinate_rows.append(row_number)
        score = output_row[header_index["geocoding_score"]]
        if score not in (None, "") and not isinstance(score, (int, float)):
            invalid_score_rows.append(row_number)
        postal = output_row[header_index["postal_code_normalized"]]
        if postal not in (None, "") and not re.fullmatch(r"\d{5}", str(postal)):
            postal_format_errors.append(row_number)

    if original_mismatches:
        raise RuntimeError(f"Original source values changed in {len(original_mismatches)} rows")
    if invalid_coordinate_rows:
        raise RuntimeError(f"Invalid coordinate cell types/values in {len(invalid_coordinate_rows)} rows")
    if invalid_score_rows:
        raise RuntimeError(f"Invalid geocoding_score cell types in {len(invalid_score_rows)} rows")
    if postal_format_errors:
        raise RuntimeError(f"postal_code_normalized format errors in {len(postal_format_errors)} rows")
    wb.close()


def write_profile_report(candidates, profile, source_sheet, mapping):
    report = {
        "generated_at": now_iso(),
        "candidates": candidates,
        "selected_sheet": source_sheet,
        "column_mapping": mapping,
        "profile": profile,
    }
    PROFILE_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile-only", action="store_true")
    args = parser.parse_args()

    ensure_dirs()
    with LOG_PATH.open("a", encoding="utf-8") as fh:
        fh.write(f"\n--- RUN {now_iso()} ---\n")
    log("Starting pharmacy geocoding pipeline")
    log(
        "Configuration: "
        f"VALID_THRESHOLD={VALID_THRESHOLD}; WARNING_THRESHOLD={WARNING_THRESHOLD}; "
        f"MAX_RETRIES={MAX_RETRIES}; TIMEOUT={TIMEOUT}; REQUEST_DELAY={REQUEST_DELAY}; "
        f"MAX_WORKERS={MAX_WORKERS}; CACHE_PATH={CACHE_PATH}; "
        f"CHECKPOINT_PATH={GEOCODED_CHECKPOINT_PATH}"
    )
    log(
        "API/service: CartoCiudad REST Geocoder candidates; "
        f"endpoint={CARTOCIUDAD_CANDIDATES_URL}; params=q,limit; "
        "queries are hierarchical and returned geometry is validated against Spanish territory"
    )
    log(
        "Normalization rules: trim/collapse whitespace, remove line breaks, normalize CP to five digits, "
        "preserve diacritics and portal suffixes, default country to Spain, retain RAW columns, "
        "and never use municipal or postal centroids as pharmacy coordinates"
    )
    source_file, candidates = discover_source_file()
    log(f"Source selected: {source_file}")
    log(f"Candidates: {json.dumps(candidates, ensure_ascii=False)}")
    profile = profile_workbook(source_file)
    source_sheet = choose_source_sheet(profile)
    log(f"Source sheet selected: {source_sheet}")
    df = pd.read_excel(source_file, sheet_name=source_sheet, dtype=str, keep_default_na=False)
    mapping = infer_column_mapping(df)
    write_profile_report(candidates, profile, source_sheet, mapping)
    log(f"Column mapping: {json.dumps(mapping, ensure_ascii=False)}")
    validate_source_schema(df, mapping)
    normalized = normalize_address_components(df, mapping)
    log(f"Normalized checkpoint saved: {NORMALIZED_CHECKPOINT_PATH}")
    if args.profile_only:
        log("Profile-only mode completed; no geocoding executed")
        return 0
    geocoded, stats = geocode_dataframe(normalized)
    qa, run_status = build_qa_report(geocoded, source_file, source_sheet, mapping, stats, len(df))
    export_results(geocoded, qa)
    validate_output(df, mapping)
    log(f"Output validated with openpyxl: {OUTPUT_PATH}")
    log(f"Geocoding statistics: {json.dumps(dict(stats), ensure_ascii=False, sort_keys=True)}")
    log(f"Result distribution: {json.dumps(geocoded['geocoding_quality'].value_counts().to_dict(), ensure_ascii=False)}")
    log(f"QA warnings: {qa.loc[qa['Metric'] == 'Warnings', 'Value'].iloc[0]}")
    log(f"QA stopper events: {qa.loc[qa['Metric'] == 'Stopper events', 'Value'].iloc[0]}")
    log(f"RUN_STATUS={run_status}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        log(str(exc))
        raise
