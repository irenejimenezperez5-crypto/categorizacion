"""Create private, local map inputs from the validated categorization artifacts.

The Excel remains the reviewed deliverable. This exporter uses the compatible
phase caches, records their hashes, and never rewrites an original input.
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import pickle
import re
import shutil
import sys

import geopandas as gpd
import numpy as np
import openpyxl
import pandas as pd
import shapely
from shapely.geometry import mapping

ROOT = Path(__file__).resolve().parents[2]
CACHE = ROOT / "pharmacy/cache/categorization"
DATA = ROOT / "visualizacion/data"
AGE_COLUMNS = ["pop_65_69", "pop_70_74", "pop_75_79", "pop_80_84", "pop_85_mas"]
EXCEL = ROOT / "pharmacy/output/UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx"
CURRENT = ROOT / "pharmacy/output/current"
F8_CONTROLS = ROOT / "pharmacy/output/categorization/F8_controls.json"
CONFIG_PATH = ROOT / "pharmacy/categorization_config.json"
SOURCE_MANIFEST = ROOT / "pharmacy/output/categorization/source_manifest.json"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    before = path.stat()
    with open_source_read(path) as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    after = path.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise ValueError(f"Source changed while reading: {path.name}")
    return h.hexdigest()


def open_source_read(path: Path):
    try:
        return path.open("rb")
    except PermissionError:
        if sys.platform != "win32":
            raise
        # OneDrive may hold a delete-sharing handle while the file is hydrated.
        # This read-only open shares read/write/delete; it does not bypass ACLs
        # and does not rename, modify, or close another application's handle.
        import ctypes
        from ctypes import wintypes
        import msvcrt
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        create = kernel.CreateFileW
        create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        create.restype = wintypes.HANDLE
        handle = create(str(path.resolve()), 0x80000000, 7, None, 3, 0x80, None)
        if handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            fd = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
        except Exception:
            close = kernel.CloseHandle
            close.argtypes = [wintypes.HANDLE]
            close.restype = wintypes.BOOL
            close(handle)
            raise
        return os.fdopen(fd, "rb")


def load(name: str):
    with (CACHE / f"{name}.pickle").open("rb") as f:
        return pickle.load(f)


def string(value) -> str:
    return "" if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)) else str(value).strip()


def code(value, length: int) -> str:
    v = string(value)
    if re.fullmatch(r"\d+\.0", v):
        v = v[:-2]
    return v.zfill(length) if v.isdigit() else v


def clean(value):
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    if isinstance(value, (int, np.integer)):
        return int(value)
    if isinstance(value, (float, np.floating)):
        return float(value) if np.isfinite(value) else None
    if isinstance(value, dict):
        return {str(k): clean(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, np.ndarray)):
        return [clean(v) for v in value]
    return value


def write_json(path: Path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(clean(value), ensure_ascii=False, separators=(",", ":"), allow_nan=False), encoding="utf-8")


def sanibrick_names(budget_codes):
    """RX2026 contains the full descriptive names missing from structure sheet."""
    book = openpyxl.load_workbook(ROOT / "pharmacy/rx/RX 2026.xlsx", read_only=True, data_only=True)
    sheet = book.worksheets[0]
    rows = sheet.iter_rows(values_only=True)
    header = list(next(rows))
    ci, ni = header.index("SB CODE"), header.index("SANIBRICK")
    names = {}
    for row in rows:
        sb, name = code(row[ci], 6), string(row[ni])
        if sb not in budget_codes:
            continue
        if name.startswith(sb):
            name = name[len(sb):].strip(" -")
        if sb in names and names[sb] != name:
            raise ValueError(f"Conflicting source SB names: {sb}")
        names[sb] = name
    book.close()
    if set(names) != set(budget_codes):
        raise ValueError("RX2026 does not cover all modeled SB names")
    return names


def build_stores(stores, crm, territories, names):
    crm_ids = crm.groupby("establishment_id", sort=False)["crm_id"].agg(lambda s: [string(x) for x in s]).to_dict()
    lookup = territories.drop_duplicates("sscc").set_index("sscc")
    municipality_lookup = territories.drop_duplicates("municipio").set_index("municipio")
    sb_lookup = territories.drop_duplicates("sb").set_index("sb")
    result = []
    for row in stores.to_dict("records"):
        sscc = code(row["sscc_modelo"], 10)
        sb = code(row["sb_modelo"], 6)
        municipality_code = code(row["municipio_modelo"], 5)
        territory = lookup.loc[sscc] if sscc in lookup.index else sb_lookup.loc[sb]
        municipality = string(municipality_lookup.loc[municipality_code, "municipio_nombre"]) if municipality_code in municipality_lookup.index else string(row["ciudad"])
        validated = np.isfinite(row["latitude_validada"]) and np.isfinite(row["longitude_validada"])
        lat = row["latitude_validada"] if validated else row["latitude_proxy"]
        lon = row["longitude_validada"] if validated else row["longitude_proxy"]
        record = {
            "id": string(row["establishment_id"]),
            "name": string(row["account_name"]),
            "type": string(row["account_type"]),
            "iqvia": code(row["iqvia_key"], 7),
            "crm_ids": crm_ids[row["establishment_id"]],
            "province": string(territory["provincia"]),
            "province_code": sscc[:2] if sscc else code(territory["sscc"], 10)[:2],
            "municipality": municipality,
            "municipality_code": municipality_code,
            "municipality_method": "conciliado" if municipality_code in municipality_lookup.index else "CRM sin código municipal validado",
            "ccaa": string(row["ccaa_modelo"]),
            "postal_code": code(row["postal_code"], 5),
            "sb_code": sb, "sb_name": names[sb], "sscc": sscc,
            "address": " ".join(x for x in [string(row["calle"]), string(row["num"])] if x),
            "lat": float(lat), "lon": float(lon),
            "location_method": string(row["metodo_ubicacion"]),
            "location_provisional": bool(row["ubicacion_provisional"]) or not validated,
            "provisional": bool(row["provisional"]),
            "territory_provisional": bool(row["territorio_provisional"]),
            "rurality": string(row["ruralidad"]) or "Sin clasificar", "degurba": int(row["degurba"]) if pd.notna(row["degurba"]) else 0,
            "category": string(row["categoria_final"]), "category_base": string(row["categoria_base"]),
            "category_reason": string(row["motivo_ajuste_categoria"]), "sb_high": bool(row["sb_alto"]),
            "potential_eur": float(row["potencial_mercado_eur"]),
            "own_eur": float(row["mercado_desde_sb_propio_eur"]),
            "neighbor_eur": float(row["mercado_desde_vecinos_eur"]),
            "sb_original_eur": float(row["mercado_rx_original_sb_eur"]),
            "sb_captured_eur": float(row["mercado_captado_sb_ubicacion_eur"]),
            "share": float(row["pct_potencial_asignado"]),
            "pop65": float(row["pop_65_mas"]), "pop75": float(row["pop_75_mas"]),
            "pop85": float(row["pop_85_mas"]),
            "weighted_population": float(row["poblacion_ponderada"]),
            "sensitive": bool(row["sensible"]),
            "sensitivity_variation": float(row["max_variacion_relativa"]),
            "crm_row_count": int(row["n_filas_crm"]),
        }
        record.update({c: float(row[c]) for c in AGE_COLUMNS})
        result.append(record)
    return result


def geographic_features(frame, property_columns, tolerance):
    """Simplify in meters; preserve individual polygon topology and record tolerance.

    Overlapping official proxy sections prevent treating the whole national
    section map as a planar coverage. Source geometry is retained in the model;
    these are display geometries, never inputs to allocation.
    """
    simple = frame.copy()
    simple.geometry = simple.geometry.simplify(tolerance, preserve_topology=True)
    simple = simple.to_crs(4326)
    features = []
    for row in simple.to_dict("records"):
        geo = shapely.set_precision(row["geometry"], 0.000001)
        if geo.is_empty:
            geo = row["geometry"]
        features.append({"type": "Feature", "properties": {k: clean(row[k]) for k in property_columns}, "geometry": mapping(geo)})
    return {"type": "FeatureCollection", "features": features}


def build_geometry(geodemography, attribution, market, names, records):
    sections = geodemography["sections"].copy()
    sections["sscc"] = sections["sscc"].map(lambda x: code(x, 10))
    sections["sb_code"] = sections["sb"].map(lambda x: code(x, 6))
    sections["municipality_code"] = sections["municipio"].map(lambda x: code(x, 5))
    sections["municipality"] = sections["municipio_nombre"].astype(str)
    sections["province_code"] = sections["sscc"].str[:2]
    sections["pop65"] = sections[AGE_COLUMNS].sum(axis=1)
    sections["pop75"] = sections[AGE_COLUMNS[2:]].sum(axis=1)
    sections["pop85"] = sections[AGE_COLUMNS[-1]]
    sections["origin_eur"] = np.asarray(attribution["section_budget_cents"]) / 100
    sections["rurality"] = sections["ruralidad"].astype(str)
    sections["provisional"] = sections["provisional_section"].astype(bool)
    sections["centroid_lon"] = sections["centroide_lon"]
    sections["centroid_lat"] = sections["centroide_lat"]
    properties = ["sscc", "sb_code", "municipality", "municipality_code", "pop65", "pop75", "pop85", "origin_eur", "rurality", "provisional", "centroid_lon", "centroid_lat", *AGE_COLUMNS]
    sections_wgs = geographic_features(sections, properties, 5)
    grouped = {}
    for feat in sections_wgs["features"]:
        grouped.setdefault(feat["properties"]["sb_code"], []).append(feat)
    for sb, feats in grouped.items():
        write_json(DATA / "sections" / f"{sb}.geojson", {"type": "FeatureCollection", "features": feats})
    del sections_wgs
    print(f"SSCC display geometry: {len(sections)} sections, {len(grouped)} files", flush=True)

    sb = sections[["sb_code", "geometry"]].dissolve(by="sb_code").reset_index()
    area_lookup = sections.groupby("sb_code", sort=False).agg(province=("provincia", "first"), ccaa=("ccaa", "first"), pop65=("pop65", "sum"), pop75=("pop75", "sum"), pop85=("pop85", "sum"), n_sections=("sscc", "size"))
    rurality = sections.groupby("sb_code")["rurality"].agg(lambda x: x.iloc[0] if x.nunique() == 1 else "Mixto")
    store_frame = pd.DataFrame(records)
    captured = store_frame.groupby("sb_code")["potential_eur"].sum()
    n_stores = store_frame.groupby("sb_code").size()
    for c in area_lookup:
        sb[c] = sb["sb_code"].map(area_lookup[c])
    sb["sb_name"] = sb["sb_code"].map(names)
    sb["original_eur"] = sb["sb_code"].map(market["budgets"]) / 100
    sb["captured_eur"] = sb["sb_code"].map(captured).fillna(0)
    sb["n_stores"] = sb["sb_code"].map(n_stores).fillna(0).astype(int)
    sb["rurality"] = sb["sb_code"].map(rurality)
    p75 = np.percentile(market["budgets"].to_numpy() / 100, 75)
    sb["high"] = sb["original_eur"] >= p75
    write_json(DATA / "sanibricks.geojson", geographic_features(sb, [c for c in sb if c != "geometry"], 35))
    print("SB display geometry written", flush=True)

    # Use all official INE sections for province outlines, including those not
    # linked exactly to the 2026 SB crosswalk, to retain the national coastline.
    raw = geodemography["raw_geometry"].copy()
    raw["code"] = raw["sscc"].map(lambda x: code(x, 10)[:2])
    province_names = sections.drop_duplicates("province_code").set_index("province_code")["provincia"].to_dict()
    provinces = raw[["code", "geometry"]].dissolve(by="code").reset_index()
    provinces["name"] = provinces["code"].map(province_names)
    write_json(DATA / "provinces.geojson", geographic_features(provinces, ["code", "name"], 150))
    print("Province outlines written", flush=True)


def write_geometry_manifest(market, universe):
    parts = []
    section_codes = set()
    for path in sorted((DATA / "sections").glob("*.geojson")):
        geo = json.loads(path.read_text(encoding="utf-8"))
        subtotal = 0.0
        for feat in geo["features"]:
            properties = feat["properties"]
            if properties["sb_code"] != path.stem or properties["sscc"] in section_codes:
                raise ValueError("Duplicate or misplaced display section")
            section_codes.add(properties["sscc"])
            subtotal += properties["origin_eur"]
        if not np.isclose(subtotal, market["budgets"].loc[path.stem] / 100, atol=0.0001, rtol=0):
            raise ValueError("Section display budgets differ from origin")
        parts.append({"file": path.relative_to(DATA).as_posix(), "features": len(geo["features"]), "bytes": path.stat().st_size, "sha256": sha256(path)})
    if section_codes != set(universe["mapping"]["sscc"]):
        raise ValueError("Display geometry sections differ from the territorial crosswalk")
    write_json(DATA / "geometry_manifest.json", {
        "geodemography_sha256": sha256(CACHE / "geodemography.pickle"),
        "attribution_sha256": sha256(CACHE / "attribution.pickle"),
        "market_sha256": sha256(CACHE / "market.pickle"),
        "sections": len(section_codes), "sanibricks": len(parts), "parts": parts,
        "provinces_sha256": sha256(DATA / "provinces.geojson"),
        "sanibricks_sha256": sha256(DATA / "sanibricks.geojson"),
    })


def export_csv(records, universe):
    physical = pd.DataFrame(records)
    physical["crm_ids"] = physical["crm_ids"].map(lambda x: " | ".join(x))
    physical.to_csv(DATA / "establecimientos.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
    # Preserve the fourteen original CRM fields and append model fields.
    crm = universe["crm"].copy()
    joined = crm[universe["original_columns"] + ["establishment_id"]].merge(
        physical, left_on="establishment_id", right_on="id", how="left", validate="many_to_one", suffixes=("", "_model")
    )
    joined["contabilizar_modelo"] = ~joined["establishment_id"].duplicated()
    joined["mercado_sumable_crm_eur"] = joined["potential_eur"].where(joined["contabilizar_modelo"], 0)
    joined.to_csv(DATA / "universo_crm.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")


def export_flows(attribution, market, stores):
    by_store = attribution["flow_cents"].tocsc()
    write_json(DATA / "flows.json", {
        "store_order": stores["establishment_id"].astype(str).tolist(),
        "origin_order": [code(x, 6) for x in market["budgets"].index],
        "origin_budgets_cents": market["budgets"].to_numpy(),
        "indptr": by_store.indptr, "indices": by_store.indices, "cents": by_store.data,
    })


def validate(records, universe, market, attribution):
    frame = pd.DataFrame(records)
    assert len(frame) == 22316 and frame["id"].is_unique
    assert len(universe["crm"]) == 22356
    assert round(frame["potential_eur"].sum() * 100) == int(market["budgets"].sum())
    assert np.allclose(frame["own_eur"] + frame["neighbor_eur"], frame["potential_eur"], atol=1e-8)
    assert ((frame["pop65"] >= frame["pop75"]) & (frame["pop75"] >= frame["pop85"])).all()
    assert frame["lat"].between(27, 45).all() and frame["lon"].between(-19, 5).all()
    assert frame["sscc"].str.fullmatch(r"\d{10}|").all()
    assert frame["sb_code"].str.fullmatch(r"\d{6}").all()
    assert frame["municipality_code"].str.fullmatch(r"\d{5}|").all()
    assert not ((frame["sb_high"]) & (frame["category"] == "D")).any()
    shares = frame.groupby("sb_code")["share"].sum()
    assert np.allclose(shares, 1, atol=1e-9)
    assert np.array_equal(np.asarray(attribution["flow_cents"].sum(axis=1)).ravel(), market["budgets"].to_numpy())
    assert np.array_equal(np.asarray(attribution["flow_cents"].sum(axis=0)).ravel(), np.rint(frame["potential_eur"].to_numpy() * 100).astype(np.int64))


def prepare_current():
    """Publish the signed V2 tables locally; keep V1 spatial artifacts as history."""
    manifest=json.loads((CURRENT/'manifest.json').read_text(encoding='utf-8'))
    controls=json.loads((CURRENT/'controls.json').read_text(encoding='utf-8'))
    for name in ['stores.json','sanibricks.json','universo_crm.csv','establecimientos.csv','sanibricks.csv',
                 'sb_prioritarios.csv','cortes_categorias.csv','thresholds.json','priority_policy.json',EXCEL.name]:
        source=CURRENT/name
        if sha256(source)!=manifest['files_sha256'][name]:
            raise ValueError(f'La entrega V2 no coincide con su manifiesto: {name}')
        if name!=EXCEL.name:
            shutil.copy2(source,DATA/name)
    records=json.loads((DATA/'stores.json').read_text(encoding='utf-8'))
    sb=json.loads((DATA/'sanibricks.json').read_text(encoding='utf-8'))
    sb_by_code={r['sb_code']:r for r in sb}
    geometry=json.loads((DATA/'sanibricks.geojson').read_text(encoding='utf-8'))
    for feature in geometry['features']:
        p=feature['properties'];r=sb_by_code[p['sb_code']]
        p.update({'captured_v1_eur':p.get('captured_v1_eur',p.get('captured_eur')),
                  'captured_eur':r['market_eur'],'sb_priority_national':r['sb_priority_national'],
                  'national_AB_n':r['national_AB_n'],'national_AB_eur':r['national_AB_eur'],
                  'visitable_national_AB_n':r['visitable_national_AB_n'],'visitable_national_AB_eur':r['visitable_national_AB_eur']})
    write_json(DATA/'sanibricks.geojson',geometry)
    # CCAA boundaries union the existing official province display shapes. No third-party tiles.
    provinces=gpd.read_file(DATA/'provinces.geojson')
    province_ccaa=pd.DataFrame(records).groupby('province_code').ccaa.agg(lambda x:x.mode().iloc[0]).to_dict()
    provinces['ccaa']=provinces['code'].map(province_ccaa)
    if provinces.ccaa.isna().any():raise ValueError('Provincia sin CCAA validada')
    ccaa=provinces[['ccaa','geometry']].dissolve(by='ccaa').reset_index()
    totals=pd.DataFrame(records).groupby('ccaa').agg(market_eur=('potential_eur','sum'),n_stores=('id','size'))
    for col in totals:ccaa[col]=ccaa.ccaa.map(totals[col])
    write_json(DATA/'ccaa.geojson',{'type':'FeatureCollection','features':[{'type':'Feature',
         'properties':{c:clean(row[c]) for c in ['ccaa','market_eur','n_stores']},'geometry':mapping(row['geometry'])}
         for row in ccaa.to_dict('records')]})
    previous=json.loads((DATA/'metadata.json').read_text(encoding='utf-8'))
    meta={**previous,'version':manifest['version'],'generated_at':manifest['generated_at'],
       'total_stores':controls['physical_unique'],'total_crm':controls['crm_rows'],'total_market_eur':manifest['market_eur'],
       'categories':controls['national_counts'],'categories_national':controls['national_counts'],'categories_ccaa':controls['ccaa_counts'],
       'priority_policy':manifest['priority_policy'],'thresholds':json.loads((DATA/'thresholds.json').read_text(encoding='utf-8')),
       'excel':{'path':str((CURRENT/EXCEL.name).relative_to(ROOT)).replace('\\','/'),
                'sha256':manifest['files_sha256'][EXCEL.name],'bytes':(CURRENT/EXCEL.name).stat().st_size},
       'model':{'version':manifest['version'],'age_weights':[1,2,3],'market_allocation':'Calibrado por SB de ubicación sobre demanda demográfica atribuida V1',
                'categories':'Cortes naturales en euros, nacional y por CCAA, sin restricciones por SB',
                'rural':'Peso igual de demanda conjunta entre puntos rurales del grupo municipio–SB, sin distancia; población original conservada','non_rural':'DEGURBA 1/2 visitable, peso espacial individual heredado y cercanía gradual','scope_counts':controls['scope_counts'],'cuts_reference_version':manifest['cuts_reference_version']},
       'share_denominator':'Mercado RX observado agregado del SB de ubicación; fracción 0–1.',
       'flows_version':'2027.1 historical only; no descomposición de los euros V2',
       'v2_validation':controls,'sources':[{'path':p,'sha256':digest} for p,digest in manifest['inputs_sha256'].items()],
       'f8_validation_historical':previous.get('f8_validation_historical',previous.get('f8_validation'))}
    meta.pop('f8_validation',None)
    for r in meta['sources']:
        if sha256(ROOT/r['path'])!=r['sha256']:raise ValueError('Input V2 cambiado: '+r['path'])
    meta['geometry']['ccaa_source']='Unión de provincias INE 2025 ya revisadas; geometría de visualización, sin modificar atribución.'
    meta['geometry']['sb_properties_version']=manifest['version']
    meta['geometry']['sanibricks_sha256']=sha256(DATA/'sanibricks.geojson')
    write_json(DATA/'export_controls.json',controls)
    meta['output_files']=[{'path':p.name,'bytes':p.stat().st_size,'sha256':sha256(p)} for p in sorted(DATA.iterdir())
                          if p.is_file() and p.name not in {'metadata.json','README.md'}]
    write_json(DATA/'metadata.json',meta)
    print(json.dumps({'version':manifest['version'],'stores':len(records),'ccaa_shapes':len(ccaa),'priority_sb':controls['priority_sb_count']},ensure_ascii=False),flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-geometry", action="store_true", help="Refresh tabular inputs using unchanged existing display geometry")
    parser.add_argument('--legacy-v1',action='store_true',help='Explicitly export the historical model; not the current default.')
    args = parser.parse_args()
    if not args.legacy_v1 and (CURRENT/'manifest.json').exists():
        prepare_current()
        return
    controls = json.loads(F8_CONTROLS.read_text(encoding="utf-8"))
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    if config != json.loads(SOURCE_MANIFEST.read_text(encoding="utf-8"))["config"]:
        raise ValueError("Model configuration differs from categorization provenance; rerun and validate categorization")
    if controls.get("state") != "complete" or sha256(EXCEL) != controls.get("xlsx_sha256"):
        raise ValueError("The reviewed Excel differs from the F8 controls; regenerate and validate categorization before exporting map inputs")
    DATA.mkdir(parents=True, exist_ok=True)
    print("Loading validated categorization tables", flush=True)
    categorized, universe, market, attribution = [load(n) for n in ["categorized", "universe", "market", "attribution"]]
    stores = categorized["stores"]
    if not np.array_equal(stores["store_index"].to_numpy(), np.arange(len(stores))):
        raise ValueError("Store order differs from the attribution sparse matrix")
    names = sanibrick_names(set(market["budgets"].index))
    records = build_stores(stores, universe["crm"], universe["mapping"], names)
    validate(records, universe, market, attribution)
    if len(records) != controls["physical_establishments"] or len(universe["crm"]) != controls["original_rows"] or int(market["budgets"].sum()) != round(controls["national_market_eur"] * 100):
        raise ValueError("Cache totals differ from the reviewed F8 Excel controls")
    if args.skip_geometry:
        previous_path = DATA / "geometry_manifest.json"
        if not previous_path.exists() or not (DATA / "sanibricks.geojson").exists() or not (DATA / "provinces.geojson").exists():
            raise ValueError("No complete previous display geometry exists; run without --skip-geometry")
        previous = json.loads(previous_path.read_text(encoding="utf-8"))
        if previous.get("geodemography_sha256") != sha256(CACHE / "geodemography.pickle") or previous.get("attribution_sha256") != sha256(CACHE / "attribution.pickle") or previous.get("market_sha256") != sha256(CACHE / "market.pickle"):
            raise ValueError("Geodemography changed; regenerate display geometry without --skip-geometry")
        if {p.stem for p in (DATA / "sections").glob("*.geojson")} != set(market["budgets"].index):
            raise ValueError("Section geometry files do not cover exactly the modeled SB")
    write_json(DATA / "stores.json", records)
    export_csv(records, universe)
    export_flows(attribution, market, stores)
    print(f"Tables written: {len(records)} establishments; {len(universe['crm'])} CRM records", flush=True)
    if not args.skip_geometry:
        print("Loading geodemography to build display geometries", flush=True)
        geodemography = load("geodemography")
        build_geometry(geodemography, attribution, market, names, records)
        del geodemography
        write_geometry_manifest(market, universe)
    sources = [EXCEL, F8_CONTROLS, CONFIG_PATH, SOURCE_MANIFEST, ROOT / "pharmacy/rx/RX 2026.xlsx", ROOT / "pharmacy/input/iqvia/Estructura SB_No Onekey_2026.xlsx", *[CACHE / f"{n}.pickle" for n in ["categorized", "universe", "market", "attribution", "geodemography"]]]
    source_records = [{"path": p.relative_to(ROOT).as_posix(), "sha256": sha256(p), "bytes": p.stat().st_size} for p in sources]
    metadata = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total_stores": len(records), "total_crm": len(universe["crm"]),
        "total_market_eur": int(market["budgets"].sum()) / 100,
        "total_sanibricks": len(market["budgets"]), "total_sections": len(universe["mapping"]),
        "categories": dict(Counter(r["category"] for r in records)),
        "provisional_locations": sum(r["location_provisional"] for r in records),
        "provisional_stores": sum(r["provisional"] for r in records),
        "sensitive_stores": sum(r["sensitive"] for r in records),
        "age_totals": {c: sum(r[c] for r in records) for c in AGE_COLUMNS + ["pop65", "pop75", "pop85"]},
        "period": "Julio 2025 – junio 2026", "market": "RX ESSITY + RESTO · sell-in IQVIA",
        "excel": source_records[0], "sources": source_records,
        "f8_validation": {"path": F8_CONTROLS.relative_to(ROOT).as_posix(), "excel_sha256_matches": True, "national_difference_cents": controls["national_difference_cents"], "rows_match": True},
        "sanibrick_names": names,
        "sb_name_source": "pharmacy/rx/RX 2026.xlsx · SANIBRICK, first sheet",
        "model": {**config, "rural": "demanda municipal sin ponderación por distancia"},
        "interpretation": "Potencial estimado y población atribuida por el modelo; no son ventas observadas ni clientes observados.",
        "location_warning": "Ubicaciones provisionales: el punto mostrado es representativo de una distribución espacial, no una coordenada precisa del establecimiento.",
        "share_denominator": "Mercado total captado por establecimientos del SB de ubicación; fracción 0–1.",
        "geometry": {"crs": "EPSG:4326", "source": "INE 2025, crosswalk IQVIA 2026", "display_simplification_m": {"provinces": 150, "sanibricks": 35, "sections": 5}, "coordinate_rounding_degrees": 0.000001, "model_geometry_unchanged": True},
    }
    metadata["output_files"] = [{"path": p.relative_to(DATA).as_posix(), "bytes": p.stat().st_size, "sha256": sha256(p)} for p in sorted(DATA.glob("*")) if p.is_file() and p.name not in {"metadata.json", "README.md"}]
    write_json(DATA / "metadata.json", metadata)
    print(json.dumps({"stores": len(records), "crm": len(universe["crm"]), "eur": metadata["total_market_eur"], "categories": metadata["categories"], "provisional_locations": metadata["provisional_locations"], "excel_sha256": metadata["excel"]["sha256"]}), flush=True)


if __name__ == "__main__":
    main()
