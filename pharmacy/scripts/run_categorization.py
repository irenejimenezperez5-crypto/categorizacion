"""Resumable national attribution pipeline; private input files remain local.

Usage: python scripts/run_categorization.py --through F8
Each phase is invalidated by changes to source fingerprints, code or configuration.
"""
from __future__ import annotations

import argparse
from collections import defaultdict
import hashlib
import json
from pathlib import Path
import pickle
import re
import sqlite3
import sys
import time
import zipfile

import geopandas as gpd
import numpy as np
import pandas as pd
from pyproj import Transformer
from rapidfuzz.fuzz import ratio, token_set_ratio
import shapely
from scipy import sparse
from scipy.spatial import cKDTree

from categorization_core import AGE_COLUMNS, age_demand, allocate_cents, categories, cents, code, distance_shares, haversine_matrix, text_key

ROOT = Path(__file__).resolve().parents[2]
PHARMACY = ROOT / "pharmacy"
CACHE = PHARMACY / "cache/categorization"
OFFICIAL = CACHE / "official"
OUT = PHARMACY / "output/categorization"
CONFIG_PATH = PHARMACY / "categorization_config.json"
STRUCTURE = ROOT / "pharmacy/input/iqvia/Estructura SB_No Onekey_2026.xlsx"
TO_XY = Transformer.from_crs(4326, 3035, always_xy=True)
TO_LL = Transformer.from_crs(3035, 4326, always_xy=True)
PHASE_NAMES = {"F1": "Universo", "F2": "Mercado", "F3": "Geodemografia", "F4": "Localizacion", "F5": "Atribucion", "F6": "Piloto y sensibilidad", "F7": "Categorias", "F8": "Entrega"}


def log(s):
    print(s, flush=True)


def sha(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda: f.read(2**20), b""):
            h.update(block)
    return h.hexdigest()


def write_json(path, data):
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2, default=lambda v: v.item() if isinstance(v, np.generic) else str(v)), encoding="utf-8")


def dump(name, value):
    target = CACHE / f"{name}.pickle"
    temporary = target.with_suffix(".part")
    with temporary.open("wb") as f:
        pickle.dump(value, f, protocol=5)
    temporary.replace(target)


def load(name):
    # Only artifacts produced locally by this pipeline are loaded.
    with (CACHE / f"{name}.pickle").open("rb") as f:
        return pickle.load(f)


def checkpoint(phase, state, detail=""):
    path = ROOT / "loop/status.md"
    doc = path.read_text(encoding="utf-8")
    replacement = f"| {phase} | {state} | {detail} |"
    doc = re.sub(rf"^\| {phase} \|.*$", replacement, doc, flags=re.M)
    doc += f"\n- 2026-10-06: {phase} {state.lower()}; {detail}.\n"
    path.write_text(doc, encoding="utf-8")


def f1(config):
    crm = pd.read_excel(PHARMACY / "univero_pharmacy_crm.xlsx", sheet_name="Sheet1", dtype=object)
    assert len(crm) == 22356 and crm.crm_id.is_unique
    original_columns = list(crm)
    crm["crm_key"] = crm.crm_id.astype(str).str.replace(r"\.0$", "", regex=True)
    crm["iqvia_key"] = crm.iqvia_id.map(lambda v: code(v, 7))
    crm["sb_crm"] = crm.sbrick_code.map(lambda v: code(v, 6))
    mapping = pd.read_excel(STRUCTURE, sheet_name="Sanibrick-SSCC-Municipios", header=2, dtype=object)
    mapping = mapping.rename(columns={"SSCC": "sscc", "CODIGO MUNICIPIO": "municipio", "CODIGO SANIBRICK 2026": "sb", "CCAA": "ccaa", "PROVINCIA": "provincia", "MUNICIPIO": "municipio_nombre"})
    mapping["sscc"] = mapping.sscc.map(lambda v: code(v, 10))
    mapping["sb"] = mapping.sb.map(lambda v: code(v, 6))
    mapping["municipio"] = mapping.municipio.map(lambda v: code(v, 5))
    mapping = mapping.dropna(subset=["sscc", "sb"]).drop_duplicates("sscc").sort_values("sscc").reset_index(drop=True)
    assert len(mapping) == 36477
    structure = pd.read_excel(STRUCTURE, sheet_name="Farmacias", header=2, dtype=object)
    structure["iqvia_key"] = structure["CODIGO IQVIA"].map(lambda v: code(v, 7))
    structure["crm_key"] = structure["CRM ID"].map(lambda v: str(v).strip().removesuffix(".0") if pd.notna(v) else None)
    iq = structure.dropna(subset=["iqvia_key"]).set_index("iqvia_key")
    assert iq.index.is_unique
    cm = structure.dropna(subset=["crm_key"])
    cm = cm[~cm.crm_key.duplicated(keep=False)].set_index("crm_key")
    rows = []
    for _, r in crm.iterrows():
        match = iq.loc[r.iqvia_key] if r.iqvia_key in iq.index else (cm.loc[r.crm_key] if r.crm_key in cm.index else None)
        found = match is not None
        section = code(match["SSCC"], 10) if found else None
        rows.append({"sscc_iqvia": section,
                     "sb_iqvia": code(match["CODIGO UBICACIÓN SANIBRICK 2026"], 6) if found else None,
                     "join_method": "IQVIA" if r.iqvia_key in iq.index else "CRM" if found else "SIN_CORRESPONDENCIA",
                     "estructura_direccion": match["DIRECCIÓN"] if found else None})
    crm = pd.concat([crm, pd.DataFrame(rows)], axis=1)
    crm["establishment_id"] = crm.iqvia_key.map(lambda v: f"IQVIA_{v}" if v else None).fillna("CRM_" + crm.crm_key)
    # A unique IQVIA establishment is an independent confirmation of repeated CRM identities.
    # Unconfirmed repetitions require an exact address match, otherwise retain CRM identity.
    for key, group in crm[crm.iqvia_key.notna() & crm.iqvia_key.duplicated(keep=False)].groupby("iqvia_key"):
        if key not in iq.index:
            address_keys = group.apply(lambda r: text_key(f"{r.ciudad} {r.calle} {r.num}"), axis=1)
            if address_keys.nunique() > 1:
                crm.loc[group.index, "establishment_id"] = "CRM_" + group.crm_key
    duplicate_audit = crm[crm.establishment_id.duplicated(keep=False)].copy()
    duplicate_audit["resolucion"] = "Identidad fisica unica confirmada por codigo IQVIA de estructura 2026; direcciones CRM alternativas conservadas"
    duplicate_audit.to_csv(OUT / "duplicados_crm.csv.gz", index=False, compression="gzip")
    stores = crm.drop_duplicates("establishment_id").copy().reset_index(drop=True)
    counts = crm.establishment_id.value_counts()
    stores["n_filas_crm"] = stores.establishment_id.map(counts)
    sec = mapping.set_index("sscc")
    stores["sb_modelo"] = stores.sscc_iqvia.map(sec.sb).fillna(stores.sb_iqvia).fillna(stores.sb_crm)
    stores["sscc_modelo"] = stores.sscc_iqvia.where(stores.sscc_iqvia.isin(sec.index))
    stores["municipio_modelo"] = stores.sscc_modelo.map(sec.municipio)
    stores["ccaa_modelo"] = stores.sscc_modelo.map(sec.ccaa).fillna(stores.ccaaa)
    stores["conflicto_sb_crm_iqvia"] = stores.sb_iqvia.notna() & stores.sb_crm.ne(stores.sb_iqvia)
    # Resolve municipality names only when unique within the known SB, not nationally.
    mun = mapping.drop_duplicates(["sb", "municipio"])
    mun_names = defaultdict(list)
    for r in mun.itertuples():
        mun_names[(r.sb, text_key(r.municipio_nombre))].append(r.municipio)
    for i in stores.index[stores.municipio_modelo.isna()]:
        candidates = mun_names.get((stores.at[i, "sb_modelo"], text_key(stores.at[i, "ciudad"])), [])
        if len(candidates) == 1:
            stores.at[i, "municipio_modelo"] = candidates[0]
    geo = pd.read_excel(PHARMACY / "output/PHARMACY_MASTER_GEOCODED_V2.xlsx", sheet_name="PHARMACY_MASTER", usecols=["crm_id", "latitude_final", "longitude_final", "coordinate_source_final", "coordinate_quality_final"])
    geo["crm_key"] = geo.crm_id.astype(str).str.replace(r"\.0$", "", regex=True)
    candidates = crm[["crm_key", "establishment_id"]].merge(geo.drop(columns="crm_id"), on="crm_key", validate="one_to_one")
    stores = stores.sort_values("establishment_id").reset_index(drop=True)
    stores["store_index"] = stores.index
    report = {"crm_rows": len(crm), "physical_establishments": len(stores), "duplicate_extra_rows": len(crm)-len(stores), "duplicate_groups": int((counts > 1).sum()), "matched_sections_rows": int(crm.sscc_iqvia.notna().sum()), "territorial_sb_conflicts": int(stores.conflicto_sb_crm_iqvia.sum()), "non_numeric_iqvia": int(crm.iqvia_key.isna().sum()), "account_types": crm.account_type.value_counts().to_dict()}
    dump("universe", {"crm": crm, "original_columns": original_columns, "stores": stores, "mapping": mapping, "geo_candidates": candidates})
    write_json(OUT / "F1_controls.json", report)
    log(f"F1 {report}")
    return report


def f2(config):
    data = load("universe")
    history = pd.read_excel(PHARMACY / "output/RX_2024_2026_SB2026.xlsx", sheet_name="RX_DATA")
    raw = pd.read_excel(PHARMACY / "rx/RX 2026.xlsx", sheet_name=0)
    periods = pd.date_range(config["period_start"], config["period_end"], freq="MS")
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    hcols = [f"{month_names[d.month-1]}-{d.strftime('%y')}" for d in periods]
    rcols = [f"{month_names[d.month-1]} {d.strftime('%y')}" for d in periods]
    def extract(df, columns):
        filtered = df[df.Manufacturer.isin(["ESSITY", "RESTO"]) & df.Measure.eq("Euros")].copy() if "Manufacturer" in df else df[df.MANUFACTURER.isin(["ESSITY", "RESTO"]) & df.Measure.eq("Euros")].copy()
        filtered["sb"] = filtered["SB CODE"].map(lambda v: code(v, 6))
        filtered = filtered[filtered.sb.isin(data["mapping"].sb)]
        # IQVIA exports '-' for zero/no recorded monthly amount. This is also
        # interpreted as zero by the existing history builder; count it explicitly.
        dash_count = int(filtered[columns].eq("-").sum().sum())
        values = filtered[columns].replace("-", 0).apply(pd.to_numeric, errors="raise")
        assert values.notna().all().all()
        money = values.map(cents)
        filtered["budget_cents"] = money.sum(axis=1)
        return filtered.groupby("sb").budget_cents.sum().astype("int64").sort_index(), int((values < 0).sum().sum()), dash_count
    budgets, negative, history_dash = extract(history, hcols)
    raw_budgets, raw_negative, raw_dash = extract(raw, rcols)
    assert budgets.equals(raw_budgets), "MAT does not reconcile between RX2026 and consolidated history"
    assert len(budgets) == 2576 and (budgets > 0).all()
    assert data["stores"].sb_modelo.isin(budgets.index).all(), "CRM establishment without a valid SB budget"
    pop = pd.read_excel(STRUCTURE, sheet_name="Habitantes por SB", header=2)
    pop["sb"] = pop["CODIGO SANIBRICK 2026"].map(lambda v: code(v, 6))
    age_inputs = [next(c for c in pop if "PobT" in c and term in c) for term in ["65-69", "70-74", "75-79", "80-84", "85"]]
    pop = pop.set_index("sb").loc[budgets.index]
    sb_population = pop[age_inputs].rename(columns=dict(zip(age_inputs, AGE_COLUMNS))).astype(float)
    sb_population["pop_total"] = pop[" POB 2026"].astype(float)
    sb_population["ccaa"] = pop.CCAA
    dump("market", {"budgets": budgets, "sb_population": sb_population})
    report = {"sb_count": len(budgets), "national_market_eur": int(budgets.sum())/100, "raw_consolidated_difference_cents": int((budgets-raw_budgets).abs().sum()), "negative_monthly_cells_preserved": negative, "raw_negative_monthly_cells": raw_negative, "raw_dash_cells_interpreted_as_zero": raw_dash, "age_population": sb_population[AGE_COLUMNS].sum().to_dict()}
    write_json(OUT / "F2_controls.json", report)
    log(f"F2 {report}")
    return report


def read_official_geometry():
    return gpd.read_file(extract_shape("sections_2025")).to_crs(3035)


def extract_shape(name):
    """Recreate the geometry files from the fingerprinted archive (ASCII directory)."""
    import shutil
    target = OFFICIAL/name
    target.mkdir(exist_ok=True)
    with zipfile.ZipFile(OFFICIAL/f"{name}.zip") as z:
        for member in z.namelist():
            if Path(member).suffix.lower() in {".shp",".dbf",".prj",".shx",".cpg"}:
                with z.open(member) as src,(target/Path(member).name).open("wb") as dst:
                    shutil.copyfileobj(src,dst)
    return next(target.glob("*.shp"))


def read_grid():
    target = OFFICIAL / "grid_spain_2021.csv.gz"
    metadata = target.with_suffix(".manifest.json")
    source_info = json.loads((OFFICIAL/"manifest.json").read_text(encoding="utf-8"))
    source_hash = next(r["sha256"] for r in source_info["sources"] if r["file"]=="grid_2021_v3.zip")
    cache_info = json.loads(metadata.read_text(encoding="utf-8")) if metadata.exists() else {}
    valid = target.exists() and cache_info.get("source_sha256")==source_hash and cache_info.get("extract_sha256")==sha(target)
    if not valid:
        with zipfile.ZipFile(OFFICIAL / "grid_2021_v3.zip") as z:
            member = next(n for n in z.namelist() if n.endswith("_V3.csv"))
            selected = []
            with z.open(member) as f:
                for chunk in pd.read_csv(f, chunksize=200000, usecols=["GRD_ID", "T", "Y_GE65", "CNTR_ID"]):
                    selected.append(chunk.loc[chunk.CNTR_ID.fillna("").str.contains(r"(?:^|-)ES(?:-|$)")])
            grid = pd.concat(selected, ignore_index=True)
            grid.to_csv(target, index=False, compression="gzip")
            write_json(metadata,{"source_sha256":source_hash,"extract_sha256":sha(target),"filter":"CNTR_ID contains country ES; T,Y_GE65"})
    grid = pd.read_csv(target)
    ids = grid.GRD_ID.str.extract(r"N(\d+)E(\d+)").astype(float)
    grid["y"] = ids[0]
    grid["x"] = ids[1]
    grid = grid[grid["T"].gt(0) | grid.Y_GE65.gt(0)].copy()
    grid = grid.dropna(subset=["x", "y"]).reset_index(drop=True)
    return grid


def mesh_for_section(geometry, grid, grid_tree, size=250.):
    hits = grid_tree.query(geometry, predicate="intersects")
    if len(hits) == 0:
        p = geometry.representative_point()
        return np.array([[p.x, p.y]]), np.ones(1), "SIN_MALLA: punto interior", False
    cells = grid.iloc[hits]
    all_boxes, masses65, masses_total = [], [], []
    for r in cells.itertuples():
        for dx in np.arange(0., 1000., size):
            for dy in np.arange(0., 1000., size):
                all_boxes.append(shapely.box(r.x+dx, r.y+dy, r.x+dx+size, r.y+dy+size))
                masses65.append(max(0., r.Y_GE65))
                masses_total.append(max(0., getattr(r, "T")))
    clipped = shapely.intersection(np.array(all_boxes, dtype=object), geometry)
    areas = shapely.area(clipped)
    weights = areas * np.array(masses65)
    method = "Malla 250m; densidad 65+ Eurostat 2021 1km"
    fallback = False
    if weights.sum() <= 0:
        weights = areas * np.array(masses_total)
        method = "Malla 250m; poblacion total 2021 (65+ nula/confidencial)"
        fallback = True
    if weights.sum() <= 0:
        weights = areas
        method = "Malla 250m; superficie (malla sin poblacion utilizable)"
        fallback = True
    keep = weights > 0
    if not keep.any():
        p = geometry.representative_point()
        return np.array([[p.x, p.y]]), np.ones(1), "SIN_INTERSECCION: punto interior", True
    p = shapely.point_on_surface(clipped[keep])
    coordinates = np.column_stack([shapely.get_x(p), shapely.get_y(p)])
    weights = weights[keep]
    return coordinates, weights / weights.sum(), method, fallback


def f3(config):
    universe, market = load("universe"), load("market")
    sections = universe["mapping"].copy()
    raw_geo = read_official_geometry()
    raw_geo["sscc"] = raw_geo.CUSEC.astype(str).str.zfill(10)
    raw_geo["municipio"] = raw_geo.CUMUN.astype(str).str.zfill(5)
    raw_geo["district"] = raw_geo.sscc.str[:7]
    raw_geo.geometry = shapely.make_valid(raw_geo.geometry.values)
    geo = raw_geo.set_index("sscc").geometry
    mun_geo = raw_geo.dissolve("municipio").geometry
    district_geo = raw_geo.dissolve("district").geometry
    sections["geometry"] = sections.sscc.map(geo)
    sections["geometry_method"] = "INE SSCC 2025 exacta"
    for i in sections.index[sections.geometry.isna()]:
        sid, mid = sections.at[i, "sscc"], sections.at[i, "municipio"]
        if sid[:7] in district_geo.index:
            sections.at[i, "geometry"] = district_geo[sid[:7]]
            sections.at[i, "geometry_method"] = "IMPUTADA: distrito INE 2025; SSCC IQVIA sin geometria"
        elif mid in mun_geo.index:
            sections.at[i, "geometry"] = mun_geo[mid]
            sections.at[i, "geometry_method"] = "IMPUTADA: municipio INE 2025"
        else:
            # Municipal changes require an explicitly recorded provincial fallback.
            candidates = raw_geo[raw_geo.sscc.str.startswith(sid[:2])]
            name_hits = candidates[candidates.NMUN.map(text_key).eq(text_key(sections.at[i, "municipio_nombre"]))]
            if len(name_hits):
                sections.at[i, "geometry"] = name_hits.geometry.union_all()
                sections.at[i, "geometry_method"] = "IMPUTADA: municipio por nombre y provincia"
            else:
                same_sb = sections.loc[sections.sb.eq(sections.at[i,"sb"]) & sections.geometry.notna(), "geometry"]
                if same_sb.empty:
                    raise RuntimeError(f"No geographic fallback for section {sid}")
                sections.at[i, "geometry"] = shapely.union_all(same_sb.values)
                sections.at[i, "geometry_method"] = "IMPUTADA: territorio SB"
    sections = gpd.GeoDataFrame(sections, geometry="geometry", crs=3035)
    ages = pd.read_csv(OFFICIAL / "ine_ages_2025_v2_filtered.csv.gz", dtype=str)
    ages["sscc"] = ages.Secciones.str.extract(r"^(\d{10})")
    ages["mid"] = ages.Municipios.str.extract(r"^(\d{5})")
    ages["value"] = pd.to_numeric(ages.Total.str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
    def age_group(s):
        if "Todas" in s: return "pop_total"
        lower = int(re.search(r"\d+", s).group())
        return {65: AGE_COLUMNS[0], 70: AGE_COLUMNS[1], 75: AGE_COLUMNS[2], 80: AGE_COLUMNS[3]}.get(lower, AGE_COLUMNS[4] if lower >= 85 else None)
    ages["age_col"] = ages.Edad.map(age_group)
    needed = ["pop_total"] + AGE_COLUMNS
    by_sec = ages.dropna(subset=["sscc"]).pivot_table(index="sscc", columns="age_col", values="value", aggfunc=lambda x: x.sum(min_count=1))[needed]
    # Check inclusion of every 85+ quinquennium against INE national totals.
    national = ages[ages.Provincias.isna()].groupby("age_col").value.sum()
    assert abs(by_sec[AGE_COLUMNS[4]].sum() - national[AGE_COLUMNS[4]]) < 1., "INE 85+ aggregation failed"
    by_mun = ages[ages.sscc.isna() & ages.mid.notna()].pivot_table(index="mid", columns="age_col", values="value", aggfunc=lambda x: x.sum(min_count=1))[needed]
    municipal_section_count = raw_geo.groupby("municipio").size()
    sections["population_method"] = "INE 2025 SSCC; reescalado por edad a total IQVIA SB2026"
    for c in needed:
        sections["ine_"+c] = sections.sscc.map(by_sec[c])
    missing = sections[["ine_"+c for c in needed]].isna().any(axis=1)
    for i in sections.index[missing]:
        mid = sections.at[i, "municipio"]
        if mid in by_mun.index:
            estimated = by_mun.loc[mid] / max(1, municipal_section_count.get(mid, 1))
            method = "IMPUTADA: media de secciones del municipio INE 2025"
        else:
            same = sections.loc[sections.sb.eq(sections.at[i, "sb"]), ["ine_"+c for c in needed]].mean().fillna(0.)
            estimated = pd.Series(same.values, index=needed)
            method = "IMPUTADA: media de secciones del SB"
        for c in needed:
            if pd.isna(sections.at[i, "ine_"+c]):
                sections.at[i, "ine_"+c] = estimated[c]
        sections.at[i, "population_method"] = method + "; reescalado IQVIA SB2026"
    zero_denominators = 0
    for sb, group in sections.groupby("sb", sort=True):
        for c in needed:
            target = market["sb_population"].at[sb, c]
            proportions = group["ine_"+c].to_numpy(float)
            if proportions.sum() <= 0:
                zero_denominators += 1
                proportions = group.ine_pop_total.to_numpy(float)
                if proportions.sum() <= 0: proportions = np.ones(len(group))
                sections.loc[group.index, "population_method"] += "; respaldo por poblacion total/neutral"
            sections.loc[group.index, c] = target * proportions / proportions.sum()
    deg = gpd.read_file(extract_shape("degurba_2021"), where="CNTR_CODE = 'ES'").to_crs(3035)
    deg["municipio"] = deg.GISCO_ID.str.replace("ES_", "", regex=False).str.zfill(5)
    deg_codes = deg.drop_duplicates("municipio").set_index("municipio").DGURBA
    sections["degurba"] = sections.municipio.map(deg_codes)
    sections["rurality_method"] = "Eurostat DEGURBA LAU 2021; codigo municipal"
    deg_tree = shapely.STRtree(deg.geometry.values)
    for i in sections.index[sections.degurba.isna()]:
        point = sections.at[i, "geometry"].representative_point()
        hits = deg_tree.query(point, predicate="intersects")
        hit = hits[0] if len(hits) else deg_tree.nearest(point)
        sections.at[i, "degurba"] = int(deg.iloc[hit].DGURBA)
        sections.at[i, "rurality_method"] = "IMPUTADA: enlace espacial DEGURBA 2021 (cambio municipal)"
    sections["degurba"] = sections.degurba.astype(int)
    sections["ruralidad"] = sections.degurba.map({1: "Ciudad", 2: "Area intermedia", 3: "Rural"})
    grid = read_grid()
    boxes = shapely.box(grid.x.values, grid.y.values, grid.x.values+1000., grid.y.values+1000.)
    tree = shapely.STRtree(boxes)
    points = []
    for i, r in enumerate(sections.itertuples()):
        # Even rural sections get a population-supported representative centre.
        xy, w, method, fallback = mesh_for_section(r.geometry, grid, tree, config["demand_grid_m"])
        center_xy = (xy*w[:,None]).sum(axis=0)
        center = shapely.Point(center_xy)
        if not r.geometry.covers(center):
            # Select closest supported point to the weighted centre, inside the section.
            center_xy = xy[np.argmin(((xy-center_xy)**2).sum(axis=1))]
        lon, lat = TO_LL.transform(*center_xy)
        sections.at[i, "centroide_lon"] = lon
        sections.at[i, "centroide_lat"] = lat
        sections.at[i, "mesh_method"] = method
        sections.at[i, "mesh_fallback"] = bool(fallback)
        sections.at[i, "n_puntos_malla"] = len(xy)
        points.append((xy, w))
        if (i+1) % 2500 == 0: log(f"F3 demand mesh: {i+1:,}/{len(sections):,}")
    sections["provisional_section"] = sections.geometry_method.str.startswith("IMPUTADA") | sections.population_method.str.startswith("IMPUTADA") | sections.rurality_method.str.startswith("IMPUTADA")
    sb_geo = sections.dissolve("sb").geometry
    sbtree = shapely.STRtree(sb_geo.values)
    neighbors = {}
    for sb, geom in sb_geo.items():
        # Small tolerance permits coincident boundary rounding differences, not arbitrary proximity.
        hits = sbtree.query(geom.buffer(2.), predicate="intersects")
        neighbors[sb] = sorted(set(sb_geo.index[hits]) | {sb})
    report = {"sections": len(sections), "official_geometry_count": len(raw_geo), "unmapped_official_sections": int((~raw_geo.sscc.isin(sections.sscc)).sum()), "imputed_geometry_sections": int(sections.geometry_method.str.startswith("IMPUTADA").sum()), "imputed_age_sections": int(missing.sum()), "zero_age_denominators": zero_denominators, "degurba_counts": sections.ruralidad.value_counts().to_dict(), "degurba_imputed": int(sections.rurality_method.str.startswith("IMPUTADA").sum()), "grid_spain_populated_cells": len(grid), "mesh_points": sum(len(x[0]) for x in points), "ine_85plus_vs_national_difference": float(by_sec[AGE_COLUMNS[4]].sum()-national[AGE_COLUMNS[4]])}
    for c in AGE_COLUMNS:
        error = (sections.groupby("sb")[c].sum() - market["sb_population"][c]).abs().max()
        assert error < 1e-6
        report[c+"_max_sb_error"] = float(error)
    dump("geodemography", {"sections": sections, "raw_geometry": raw_geo, "municipal_geometry": mun_geo, "points": points, "neighbors": neighbors})
    sections.drop(columns="geometry").to_csv(OUT / "sscc_geodemografia.csv.gz", index=False, compression="gzip")
    # Explicit crosswalk exceptions, not a silent assertion that 2025 == 2026.
    sections[sections.provisional_section].drop(columns="geometry").to_csv(OUT / "correspondencias_territoriales_provisionales.csv.gz", index=False, compression="gzip")
    write_json(OUT / "F3_controls.json", report)
    log(f"F3 {report}")
    return report


def sample_distribution(xy, probabilities, maximum=16):
    """Deterministic equal-mass quadrature; not a fabricated exact position."""
    xy, p = np.asarray(xy), np.asarray(probabilities, dtype=float)
    p = p / p.sum()
    if len(p) <= maximum:
        return xy, p
    order = np.lexsort((xy[:, 1], xy[:, 0]))
    selected = np.searchsorted(np.cumsum(p[order]), (np.arange(maximum)+.5)/maximum)
    unique, count = np.unique(order[np.minimum(selected, len(order)-1)], return_counts=True)
    return xy[unique], count/maximum


def f4(config):
    universe, geo = load("universe"), load("geodemography")
    stores = universe["stores"].copy()
    sections, raw = geo["sections"], geo["raw_geometry"]
    sec_index = sections.set_index("sscc")
    tree = shapely.STRtree(raw.geometry.values)
    candidates = universe["geo_candidates"].dropna(subset=["latitude_final", "longitude_final"]).copy()
    x, y = TO_XY.transform(candidates.longitude_final.to_numpy(float), candidates.latitude_final.to_numpy(float))
    pts = shapely.points(x, y)
    hits = tree.query(pts, predicate="intersects")
    coordinate_sections = {int(a): raw.iloc[int(b)].sscc for a, b in zip(*hits)}
    by_establishment = defaultdict(list)
    for j, r in enumerate(candidates.itertuples()):
        sid = coordinate_sections.get(j)
        if sid: by_establishment[r.establishment_id].append((sid, r))
    stores["latitude_validada"] = np.nan
    stores["longitude_validada"] = np.nan
    stores["metodo_ubicacion"] = ""
    stores["revision_identidad"] = stores.n_filas_crm.gt(1)
    stores["incidencia_ubicacion"] = ""
    rejects = 0
    accepted = 0
    for i, r in stores.iterrows():
        options = []
        for sid, candidate in by_establishment.get(r.establishment_id, []):
            if pd.notna(r.municipio_modelo) and sid[:5] != r.municipio_modelo:
                rejects += 1
                continue
            if pd.isna(r.municipio_modelo):
                matches_sb = sid in sec_index.index and sec_index.at[sid, "sb"] == r.sb_modelo
                actual_name = raw.loc[raw.sscc.eq(sid), "NMUN"].iloc[0]
                if not matches_sb and token_set_ratio(text_key(r.ciudad), text_key(actual_name)) < 92:
                    rejects += 1
                    continue
            preference = (100 if sid == r.sscc_iqvia else 0) + (10 if candidate.coordinate_quality_final == "EXACT" else 0)
            options.append((preference, sid, candidate))
        if options:
            _, sid, candidate = sorted(options, key=lambda o: (-o[0], o[1], o[2].crm_key))[0]
            stores.at[i, "latitude_validada"] = candidate.latitude_final
            stores.at[i, "longitude_validada"] = candidate.longitude_final
            stores.at[i, "metodo_ubicacion"] = f"{candidate.coordinate_source_final}; validada en municipio INE"
            stores.at[i, "municipio_modelo"] = sid[:5]
            if sid in sec_index.index:
                stores.at[i, "sscc_modelo"] = sid
                stores.at[i, "sb_modelo"] = sec_index.at[sid, "sb"]
                stores.at[i, "ccaa_modelo"] = sec_index.at[sid, "ccaa"]
            else:
                stores.at[i, "incidencia_ubicacion"] = "SSCC del punto INE2025 no existe en mapa IQVIA2026; se conserva enlace territorial IQVIA"
            accepted += 1
        elif by_establishment.get(r.establishment_id):
            stores.at[i, "incidencia_ubicacion"] = "Coordenada V2 rechazada: municipio distinto de estructura IQVIA"
    # Address evidence can recover a point even when commercial and titular names differ.
    # Require explicit pharmacy, point geometry, exact house number, municipality and nonambiguous street.
    con = sqlite3.connect(f"file:{(PHARMACY/'cache/poi_match_cache_v2.sqlite').as_posix()}?mode=ro", uri=True)
    osm = pd.read_sql_query("SELECT osm_type,osm_id,name,lat,lon,is_point,tags_json FROM osm_pharmacies WHERE is_point=1", con)
    con.close()
    from prepare_pharmacy_addresses import prepare as prepare_addresses
    full_address_file, address_source = prepare_addresses()
    full_addresses = pd.read_csv(full_address_file)
    osm = pd.concat([full_addresses,osm],ignore_index=True).drop_duplicates(["osm_type","osm_id"],keep="first").reset_index(drop=True)
    ox, oy = TO_XY.transform(osm.lon.values, osm.lat.values)
    osm_hits = tree.query(shapely.points(ox, oy), predicate="intersects")
    osm_sid = {int(a): raw.iloc[int(b)].sscc for a,b in zip(*osm_hits)}
    addresses = defaultdict(list)
    def house(v):
        s = text_key(v).removesuffix(" 0")
        return str(int(s)) if s.isdigit() else s
    def street(v):
        s = text_key(v)
        return re.sub(r"^(?:C|CL|CALLE|AV|AVDA|AVENIDA|PL|PLAZA|PZ|PS|PASEO|PG|PGO|RUA|CARRETERA|CTRA)\s+", "", s)
    for j, r in enumerate(osm.itertuples()):
        sid = osm_sid.get(j)
        tags = json.loads(r.tags_json)
        if not sid or tags.get("amenity") != "pharmacy" or not tags.get("addr:street") or not tags.get("addr:housenumber"):
            continue
        addresses[(sid[:5], house(tags["addr:housenumber"]))].append((street(tags["addr:street"]), sid, r))
    recovered = 0
    used_osm = set()
    for i in stores.index[stores.latitude_validada.isna()]:
        r = stores.loc[i]
        choices = []
        for st, sid, o in addresses.get((r.municipio_modelo, house(r.num)), []):
            score = ratio(street(r.calle), st)
            if score >= 94 and (o.osm_type, o.osm_id) not in used_osm:
                choices.append((score, sid, o))
        if not choices: continue
        choices.sort(key=lambda v: (-v[0], v[1], v[2].osm_id))
        if len(choices) > 1:
            ll = np.array([[o[2].lon, o[2].lat] for o in choices])
            if haversine_matrix(ll[:1], ll).max() > 75.:
                continue
        _, sid, o = choices[0]
        stores.at[i, "latitude_validada"] = o.lat
        stores.at[i, "longitude_validada"] = o.lon
        stores.at[i, "metodo_ubicacion"] = f"OSM {o.osm_type}/{o.osm_id}; direccion + numero + municipio"
        if sid in sec_index.index:
            stores.at[i, "sscc_modelo"] = sid
            stores.at[i, "sb_modelo"] = sec_index.at[sid, "sb"]
            stores.at[i, "ccaa_modelo"] = sec_index.at[sid, "ccaa"]
        used_osm.add((o.osm_type, o.osm_id))
        recovered += 1
    placements_xy, placements_ll, placements_p, offsets = [], [], [], []
    presence = defaultdict(dict)
    sections_by_mun = {k: np.array(v, dtype=int) for k,v in sections.groupby("municipio").groups.items()}
    sections_by_sb = {k: np.array(v, dtype=int) for k,v in sections.groupby("sb").groups.items()}
    sections_index = {sid: i for i,sid in enumerate(sections.sscc)}
    for i, r in stores.iterrows():
        offsets.append(len(placements_p))
        if pd.notna(r.latitude_validada):
            lonlat = np.array([[r.longitude_validada, r.latitude_validada]])
            xx, yy = TO_XY.transform(lonlat[:,0], lonlat[:,1])
            xy, p = np.column_stack([xx, yy]), np.ones(1)
            presence[r.municipio_modelo][i] = 1.
        else:
            if r.sscc_modelo in sections_index:
                idx = np.array([sections_index[r.sscc_modelo]])
                method = "PROVISIONAL: distribucion en SSCC conocida"
            elif r.municipio_modelo in sections_by_mun:
                idx = sections_by_mun[r.municipio_modelo]
                method = "PROVISIONAL: distribucion en municipio"
            else:
                idx = sections_by_sb[r.sb_modelo]
                method = "PROVISIONAL: distribucion en SB"
            mass = age_demand(sections.loc[idx, AGE_COLUMNS].to_numpy())
            if mass.sum() <= 0: mass = np.ones(len(idx))
            mass = mass / mass.sum()
            xy_all, p_all = [], []
            for j, weight in zip(idx, mass):
                xy_j, p_j = geo["points"][j]
                xy_all.append(xy_j)
                p_all.append(weight*p_j)
                mid = sections.at[j, "municipio"]
                presence[mid][i] = presence[mid].get(i, 0.) + float(weight)
            xy, p = sample_distribution(np.vstack(xy_all), np.concatenate(p_all), maximum=16)
            lon, lat = TO_LL.transform(xy[:,0], xy[:,1])
            lonlat = np.column_stack([lon, lat])
            stores.at[i, "metodo_ubicacion"] = method + "; cuadratura hasta 16 puntos de masa"
        placements_xy.extend(xy)
        placements_ll.extend(lonlat)
        placements_p.extend(p)
        stores.at[i, "n_ubicaciones_modeladas"] = len(p)
        stores.at[i, "longitude_proxy"] = float((lonlat[:,0]*p).sum())
        stores.at[i, "latitude_proxy"] = float((lonlat[:,1]*p).sum())
    offsets.append(len(placements_p))
    mun_deg = sections.drop_duplicates("municipio").set_index("municipio").degurba
    stores["degurba"] = stores.municipio_modelo.map(mun_deg)
    stores["ruralidad"] = stores.degurba.map({1: "Ciudad", 2: "Area intermedia", 3: "Rural"}).fillna("Mixto/no resuelto")
    stores["ubicacion_provisional"] = stores.latitude_validada.isna()
    stores["territorio_provisional"] = stores.sscc_modelo.isna() | stores.sscc_modelo.map(sec_index.provisional_section).fillna(True) | stores.incidencia_ubicacion.ne("")
    stores["provisional"] = stores.ubicacion_provisional | stores.territorio_provisional | stores.revision_identidad
    coord = stores.dropna(subset=["latitude_validada"])
    shared = coord.groupby(["latitude_validada", "longitude_validada"]).establishment_id.transform("size")
    stores["coordenada_compartida"] = False
    stores.loc[coord.index, "coordenada_compartida"] = shared.gt(1)
    stores["sb_modificado_vs_crm"] = stores.sb_modelo.ne(stores.sb_crm)
    report = {"physical_establishments": len(stores), "validated_v2_establishments": accepted, "recovered_osm_address": recovered, "coordinates_rejected_municipal_conflict": rejects, "validated_coordinates": int(stores.latitude_validada.notna().sum()), "uncertain_location_establishments": int(stores.ubicacion_provisional.sum()), "provisional_establishments": int(stores.provisional.sum()), "territorial_location_changes_vs_crm": int(stores.sb_modificado_vs_crm.sum()), "modeled_placements": len(placements_p), "shared_coordinate_establishments": int(stores.coordenada_compartida.sum()),"osm_full_address_source":address_source}
    dump("locations", {"stores": stores, "placement_xy": np.array(placements_xy), "placement_ll": np.array(placements_ll), "placement_p": np.array(placements_p), "offsets": np.array(offsets, dtype=int), "municipal_presence": dict(presence)})
    stores.to_csv(OUT / "establecimientos_localizados.csv.gz", index=False, compression="gzip")
    write_json(OUT / "F4_controls.json", report)
    log(f"F4 {report}")
    return report


def aggregate_scenario(matrix, sections, stores, market, weights):
    """Preserve origin budgets independently of final store/location-SB totals."""
    budgets = market["budgets"]
    demands = age_demand(sections[AGE_COLUMNS].to_numpy(float), weights)
    sec_money = np.zeros(len(sections))
    for sb, indices in sections.groupby("sb").groups.items():
        idx = np.array(indices, dtype=int)
        d = demands[idx]
        if d.sum() <= 0: d = sections.loc[idx,"pop_total"].to_numpy(float)
        if d.sum() <= 0: d = np.ones(len(idx))
        sec_money[idx] = int(budgets[sb]) * d/d.sum()
    origin_index = {sb: i for i,sb in enumerate(budgets.index)}
    group = sparse.csr_matrix((np.ones(len(sections)), (sections.sb.map(origin_index).values, np.arange(len(sections)))), shape=(len(budgets), len(sections)))
    flows_unrounded = (group @ matrix.multiply(sec_money[:,None])).tocsr()
    flows_unrounded.sort_indices()  # Stable establishment-ID order for leftover-cent ties.
    flows = flows_unrounded.copy()
    for i in range(len(budgets)):
        left,right = flows.indptr[i:i+2]
        if right == left: raise RuntimeError(f"No eligible store for SB {budgets.index[i]}")
        flows.data[left:right] = allocate_cents(int(budgets.iloc[i]), flows.data[left:right])
    flows.data = flows.data.astype(np.int64)
    potential = np.asarray(flows.sum(axis=0)).ravel().astype(np.int64)
    attributed_population = matrix.T @ sections[AGE_COLUMNS].to_numpy(float)
    return {"potential_cents": potential, "flow_cents": flows, "flows_unrounded": flows_unrounded, "population": attributed_population, "section_budget_cents": sec_money}


def f5(config):
    geo, loc, market = load("geodemography"), load("locations"), load("market")
    sections, stores = geo["sections"], loc["stores"]
    ns, nf = len(sections), len(stores)
    points_tree = cKDTree(loc["placement_xy"])
    placement_store = np.repeat(np.arange(nf), np.diff(loc["offsets"]))
    stores_by_sb = {sb: np.array(v,dtype=int) for sb,v in stores.groupby("sb_modelo").groups.items()}
    rows, cols = [], []
    values = {name: [] for name in ["base", "beta_1_5", "beta_3", "centroide"]}
    rural_fallbacks, candidate_fallbacks = [], []
    max_candidates = 0
    section_own_shares = np.zeros(ns)
    for i,r in enumerate(sections.itertuples()):
        if r.degurba == 3:
            candidates = loc["municipal_presence"].get(r.municipio, {})
            if candidates:
                c = np.array(sorted(candidates), dtype=int)
                w = np.array([candidates[j] for j in c], dtype=float)
                w /= w.sum()
            else:
                c = stores_by_sb.get(r.sb, np.array([], dtype=int))
                if len(c) == 0:
                    raise RuntimeError(f"Rural municipality and SB have no eligible stores: {r.municipio}/{r.sb}")
                w = np.ones(len(c))/len(c)
                rural_fallbacks.append(i)
            shares = {name: w for name in values}
        else:
            initial = set()
            for sb in geo["neighbors"].get(r.sb, [r.sb]):
                initial.update(stores_by_sb.get(sb, []))
            xy, dw = geo["points"][i]
            # Every point within 3 km of the section is eligible, in addition to own/adjacent SB.
            bounds = r.geometry.bounds
            midpoint = np.array([(bounds[0]+bounds[2])/2,(bounds[1]+bounds[3])/2])
            search_radius = np.linalg.norm(midpoint-np.array(bounds[:2])) + config["candidate_radius_m"]
            near = np.array(points_tree.query_ball_point(midpoint, search_radius), dtype=int)
            if len(near):
                dist = shapely.distance(shapely.points(loc["placement_xy"][near]), r.geometry)
                initial.update(placement_store[near[dist <= config["candidate_radius_m"]]])
            if not initial:
                # Geographical holes must still retain the origin budget; nearest candidates flagged.
                k = min(8,len(loc["placement_xy"]))
                _, nearest = points_tree.query(midpoint, k=k)
                initial.update(placement_store[np.atleast_1d(nearest)])
                candidate_fallbacks.append(i)
            c = np.array(sorted(initial),dtype=int)
            starts, ends = loc["offsets"][c], loc["offsets"][c+1]
            placements = np.concatenate([np.arange(a,b) for a,b in zip(starts,ends)])
            offsets = np.r_[0, np.cumsum(ends-starts)[:-1]]
            pp = loc["placement_p"][placements]
            lon, lat = TO_LL.transform(xy[:,0],xy[:,1])
            ll = np.column_stack([lon,lat])
            # Bound temporary matrices for extensive intermediate municipalities.
            shares = {name: np.zeros(len(c)) for name in ["base", "beta_1_5", "beta_3"]}
            chunk_size = max(1, min(128, int(2e6/max(1,len(placements)))))
            for left in range(0,len(ll),chunk_size):
                right = min(len(ll),left+chunk_size)
                d = haversine_matrix(ll[left:right],loc["placement_ll"][placements])
                prob = dw[left:right]
                mass = prob.sum()
                for name, beta in [("base",config["distance_exponent"]),("beta_1_5",1.5),("beta_3",3.)]:
                    shares[name] += mass * distance_shares(d,offsets,pp,prob/mass,beta,config["distance_offset_m"])
            d = haversine_matrix(np.array([[r.centroide_lon,r.centroide_lat]]),loc["placement_ll"][placements])
            shares["centroide"] = distance_shares(d,offsets,pp,np.ones(1),config["distance_exponent"],config["distance_offset_m"])
        max_candidates = max(max_candidates,len(c))
        rows.extend([i]*len(c))
        cols.extend(c)
        for name in values: values[name].extend(shares[name])
        section_own_shares[i] = shares["base"][stores.loc[c,"sb_modelo"].eq(r.sb).values].sum()
        if (i+1)%1500 == 0: log(f"F5 attribution: {i+1:,}/{ns:,} sections; {len(rows):,} edges")
    matrices = {name: sparse.csr_matrix((np.array(v), (np.array(rows),np.array(cols))),shape=(ns,nf)) for name,v in values.items()}
    for name,m in matrices.items():
        err = np.abs(np.asarray(m.sum(axis=1)).ravel()-1.).max()
        assert err < 1e-10, f"Nonconserved shares: {name}"
        sparse.save_npz(CACHE/f"shares_{name}.npz",m)
    result = aggregate_scenario(matrices["base"],sections,stores,market,config["age_weights"])
    flow = result["flow_cents"]
    row_error = np.asarray(flow.sum(axis=1)).ravel()-market["budgets"].values
    assert np.max(np.abs(row_error)) == 0
    age_error = np.abs(result["population"].sum(axis=0)-sections[AGE_COLUMNS].sum().values)
    assert age_error.max() < 1e-5
    own = np.zeros(nf,dtype=np.int64)
    for j,sb in enumerate(market["budgets"].index):
        same = stores_by_sb.get(sb,[])
        if len(same): own[same] = np.asarray(flow[j,same].todense()).ravel()
    result["own_cents"] = own
    result["neighbor_cents"] = result["potential_cents"]-own
    result["section_own_share"] = section_own_shares
    result["rural_fallback_sections"] = rural_fallbacks
    result["candidate_fallback_sections"] = candidate_fallbacks
    dump("attribution",result)
    sparse.save_npz(OUT/"flujos_centimos_base.npz",flow)
    report = {"national_market_eur": int(result["potential_cents"].sum())/100, "max_origin_sb_difference_cents": int(np.abs(row_error).max()), "population_absolute_error_by_age": age_error.tolist(), "section_store_edges": len(rows), "max_candidates_per_section": max_candidates, "rural_no_municipal_store_sections": len(rural_fallbacks), "nonrural_nearest_fallback_sections": len(candidate_fallbacks), "cross_sb_market_eur": int(result["neighbor_cents"].sum())/100, "nonrural_cross_sb_sections": int(((section_own_shares < 1.-1e-9)&sections.degurba.ne(3)).sum())}
    write_json(OUT/"F5_controls.json",report)
    log(f"F5 {report}")
    return report


def f6(config):
    geo,loc,market,base = load("geodemography"),load("locations"),load("market"),load("attribution")
    sections,stores = geo["sections"],loc["stores"]
    base_stores = stores.copy()
    base_stores["potential_cents"] = base["potential_cents"]
    base_categories = categories(base_stores,market["budgets"],config["high_sb_quantile"],config["category_cutoffs"])
    scenarios = [("centroide",config["age_weights"]),("beta_1_5",config["age_weights"]),("beta_3",config["age_weights"]),("edades_1_1_1",[1.,1.,1.]),("edades_1_2_4",[1.,2.,4.])]
    detail = stores[["establishment_id","sb_modelo","ccaa_modelo","ruralidad","provisional"]].copy()
    detail["base_eur"] = base["potential_cents"]/100.
    detail["categoria_base_escenario"] = base_categories.categoria_final
    report_rows = []
    sensitive = np.zeros(len(stores),dtype=bool)
    max_variation = np.zeros(len(stores))
    for name,weights in scenarios:
        matrix_name = "base" if name.startswith("edades") else name
        matrix = sparse.load_npz(CACHE/f"shares_{matrix_name}.npz")
        result = aggregate_scenario(matrix,sections,stores,market,weights)
        scenario_stores = stores.copy()
        scenario_stores["potential_cents"] = result["potential_cents"]
        cats = categories(scenario_stores,market["budgets"],config["high_sb_quantile"],config["category_cutoffs"])
        denom = base["potential_cents"].astype(float)
        variation = np.divide(result["potential_cents"]-denom,denom,out=np.zeros(len(stores)),where=denom!=0)
        variation[(denom==0)&(result["potential_cents"]>0)] = np.inf
        changed = cats.categoria_final.ne(base_categories.categoria_final).values
        sensitive |= changed | (np.abs(variation)>config["sensitivity_relative_threshold"])
        max_variation = np.maximum(max_variation,np.abs(variation))
        detail[name+"_eur"] = result["potential_cents"]/100.
        detail[name+"_variacion"] = variation
        detail[name+"_categoria"] = cats.categoria_final
        report_rows.append({"escenario":name,"mercado_eur":int(result["potential_cents"].sum())/100,"categorias_cambiadas":int(changed.sum()),"variacion_mayor_20pct":int((np.abs(variation)>.2).sum())})
    detail["sensible"] = sensitive
    detail["max_variacion_relativa"] = max_variation
    # Pilot: stratify urban/intermediate/rural/mixed and SB/CCAA capture boundaries.
    sb_degurba_count = sections.groupby("sb").degurba.nunique()
    strata = {"urbano":stores.degurba.eq(1),"intermedio":stores.degurba.eq(2),"rural":stores.degurba.eq(3),"sb_mixto":stores.sb_modelo.map(sb_degurba_count).gt(1),"frontera_sb":base["neighbor_cents"]>base["potential_cents"]*.2,"ubicacion_incierta":stores.ubicacion_provisional}
    rng = np.random.default_rng(config["random_seed"])
    selected = []
    for stratum, mask in strata.items():
        indices = np.flatnonzero(np.asarray(mask))
        take = min(4,len(indices))
        for index in rng.choice(indices,take,replace=False): selected.append((stratum,int(index)))
    pilot = pd.DataFrame([{"estrato":stratum,**detail.loc[i].to_dict(),"sscc":stores.at[i,"sscc_modelo"],"municipio":stores.at[i,"municipio_modelo"],"metodo_ubicacion":stores.at[i,"metodo_ubicacion"]} for stratum,i in selected])
    detail.to_csv(OUT/"sensibilidad_establecimientos.csv.gz",index=False,compression="gzip")
    pilot.to_csv(OUT/"piloto.csv",index=False,encoding="utf-8-sig")
    from audit_walking_pilot import audit_walking
    walking = audit_walking(stores,sections,pilot,CACHE,OUT)
    dump("sensitivity",{"detail":detail,"scenarios":pd.DataFrame(report_rows),"pilot":pilot,"walking":walking})
    report = {"scenarios":report_rows,"sensitive_establishments":int(sensitive.sum()),"pilot_cases":len(pilot),"pilot_seed":config["random_seed"],"walking_routes":len(walking),"walking_route_states":walking.estado.value_counts().to_dict(),"walking_max_distance_ratio":float(walking.ratio_pie_geografica.max()) if "ratio_pie_geografica" in walking else None}
    write_json(OUT/"F6_controls.json",report)
    log(f"F6 {report}")
    return report


def f7(config):
    stores,market,base,sensitivity = load("locations")["stores"].copy(),load("market"),load("attribution"),load("sensitivity")
    stores["potential_cents"] = base["potential_cents"]
    stores = categories(stores,market["budgets"],config["high_sb_quantile"],config["category_cutoffs"])
    for j,c in enumerate(AGE_COLUMNS): stores[c] = base["population"][:,j]
    stores["pop_65_mas"] = stores[AGE_COLUMNS].sum(axis=1)
    stores["pop_75_mas"] = stores[AGE_COLUMNS[2:]].sum(axis=1)
    stores["poblacion_ponderada"] = age_demand(stores[AGE_COLUMNS].values,config["age_weights"])
    stores["potencial_mercado_eur"] = stores.potential_cents/100.
    stores["mercado_desde_sb_propio_eur"] = base["own_cents"]/100.
    stores["mercado_desde_vecinos_eur"] = base["neighbor_cents"]/100.
    stores["mercado_rx_original_sb_eur"] = stores.sb_modelo.map(market["budgets"])/100.
    stores["mercado_captado_sb_ubicacion_eur"] = stores.groupby("sb_modelo").potencial_mercado_eur.transform("sum")
    stores["pct_potencial_asignado"] = np.divide(stores.potencial_mercado_eur,stores.mercado_captado_sb_ubicacion_eur,out=np.zeros(len(stores)),where=stores.mercado_captado_sb_ubicacion_eur.ne(0))
    stores["sensible"] = sensitivity["detail"].sensible.values
    stores["max_variacion_relativa"] = sensitivity["detail"].max_variacion_relativa.values
    stores["periodo_mercado"] = "Julio 2025 - junio 2026"
    assert (stores.pop_65_mas+1e-8>=stores.pop_75_mas).all()
    assert (stores.pop_75_mas+1e-8>=stores.pop_85_mas).all()
    positive = stores.mercado_captado_sb_ubicacion_eur.gt(0)
    assert (stores[positive].groupby("sb_modelo").pct_potencial_asignado.sum()-1.).abs().max()<1e-10
    distributions = []
    for version in ["base","final"]:
        for (ccaa,cat),group in stores.groupby(["ccaa_modelo",f"categoria_{version}"]):
            distributions.append({"ccaa":ccaa,"version":version,"categoria":cat,"establecimientos":len(group),"mercado_eur":int(group.potential_cents.sum())/100.,"porcentaje_mercado_ccaa":group.potential_cents.sum()/stores.loc[stores.ccaa_modelo.eq(ccaa),"potential_cents"].sum()})
    dump("categorized",{"stores":stores,"category_distribution":pd.DataFrame(distributions)})
    stores.to_csv(OUT/"establecimientos_categorizados.csv.gz",index=False,compression="gzip")
    report = {"category_base_counts":stores.categoria_base.value_counts().to_dict(),"category_final_counts":stores.categoria_final.value_counts().to_dict(),"category_adjustments":int(stores.motivo_ajuste_categoria.ne("").sum()),"high_sb_threshold_eur":float(market["budgets"].quantile(config["high_sb_quantile"]))/100.,"high_sb_count":int(market["budgets"].ge(market["budgets"].quantile(config["high_sb_quantile"])).sum()),"high_sb_D_establishments":int((stores.sb_alto & stores.categoria_final.eq("D")).sum()),"high_sb_entirely_A":int(stores[stores.sb_alto].groupby("sb_modelo").categoria_final.apply(lambda g:g.eq("A").all()).sum())}
    write_json(OUT/"F7_controls.json",report)
    log(f"F7 {report}")
    return report


def f8(config):
    from export_categorization import export_and_validate
    return export_and_validate(ROOT,config,load("universe"),load("market"),load("geodemography"),load("categorized"),load("attribution"),load("sensitivity"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--through", choices=list(PHASE_NAMES), default="F8")
    parser.add_argument("--force", action="store_true")
    parser.add_argument('--legacy-v1',action='store_true',help='Explicitly reproduce the historical F1–F8 categories. Current national V2 is the default final delivery.')
    args = parser.parse_args()
    if not args.legacy_v1 and args.through in {'F7','F8'}:
        # The signed spatial foundation stays separate from the current economic policy.
        # Do not accidentally replace V2 with the old SB-protected F7/F8 output.
        import subprocess
        for filename in ['market.pickle','universe.pickle','attribution.pickle']:
            if not (CACHE/filename).exists():
                raise RuntimeError('Falta la base espacial: ejecutar primero --legacy-v1 y exportar su visualización para disponer de la referencia V1 revisada.')
        subprocess.run([sys.executable,str(Path(__file__).with_name('national_categorization.py'))],check=True,cwd=ROOT)
        if args.through == 'F8':
            subprocess.run([sys.executable,str(ROOT/'visualizacion/scripts/prepare_data.py')],check=True,cwd=ROOT)
        return
    OUT.mkdir(parents=True, exist_ok=True)
    CACHE.mkdir(parents=True, exist_ok=True)
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    if config["population_year"]!=2025 or config["degurba_year"]!=2021:
        raise ValueError("This release supports INE2025 and DEGURBA2021; obtain and reconcile new editions before changing their years")
    if len(config["age_weights"])!=3 or min(config["age_weights"])<0 or sum(config["age_weights"])<=0:
        raise ValueError("Age weights require three nonnegative values with positive total")
    if config["distance_exponent"]<=0 or config["distance_offset_m"]<=0 or config["candidate_radius_m"]<0:
        raise ValueError("Invalid distance parameters")
    if config["demand_grid_m"]<=0 or abs(1000/config["demand_grid_m"]-round(1000/config["demand_grid_m"]))>1e-9:
        raise ValueError("Demand-grid width must divide the official1000m cell exactly")
    inputs = [PHARMACY / "univero_pharmacy_crm.xlsx", STRUCTURE, PHARMACY / "output/PHARMACY_MASTER_GEOCODED_V2.xlsx", PHARMACY / "output/RX_2024_2026_SB2026.xlsx", PHARMACY / "rx/RX 2026.xlsx", PHARMACY / "cache/poi_match_cache_v2.sqlite", PHARMACY / "cache/osm/spain-latest.osm.pbf", ROOT / "pharmacy/input/iqvia/WELCOME PACK SANIBRICKS 2026.pdf"]
    source_records = [{"file": str(p.relative_to(ROOT)), "sha256": sha(p), "bytes": p.stat().st_size} for p in inputs]
    public = json.loads((OFFICIAL / "manifest.json").read_text(encoding="utf-8"))
    write_json(OUT / "source_manifest.json", {"private_local_sources": source_records, "public_sources": public, "config": config})
    run_sig = hashlib.sha256(json.dumps({"sources": source_records, "public": public, "config": config}, sort_keys=True).encode()).hexdigest()
    upstream = run_sig
    for phase in PHASE_NAMES:
        # Phase-scoped signatures: code edits to export do not rerun earlier computations.
        function = globals().get(phase.lower())
        if function is None:
            raise RuntimeError(f"Phase {phase} has not been implemented")
        import inspect
        implementation = inspect.getsource(function)
        helper_dependencies = {"F3":[read_official_geometry,extract_shape,read_grid,mesh_for_section],"F4":[sample_distribution],"F5":[aggregate_scenario],"F6":[aggregate_scenario]}
        implementation += "".join(inspect.getsource(h) for h in helper_dependencies.get(phase,[]))
        companion = {"F4":"prepare_pharmacy_addresses.py","F6":"audit_walking_pilot.py","F8":"export_categorization.py"}.get(phase)
        if companion: implementation += sha(Path(__file__).with_name(companion))
        core_hash = sha(Path(__file__).with_name("categorization_core.py"))
        signature = hashlib.sha256((upstream + implementation + core_hash).encode()).hexdigest()
        manifest_file = OUT / f"{phase}_manifest.json"
        old = json.loads(manifest_file.read_text(encoding="utf-8")) if manifest_file.exists() else {}
        if not args.force and old.get("signature") == signature and old.get("state") == "complete":
            log(f"CACHE {phase}: compatible inputs and parameters")
        else:
            checkpoint(phase, "En curso", PHASE_NAMES[phase])
            started = time.monotonic()
            try:
                report = function(config)
                write_json(manifest_file, {"phase": phase, "signature": signature, "state": "complete", "elapsed_seconds": round(time.monotonic()-started, 2), "controls": report})
                checkpoint(phase, "Completada", f"Controles: pharmacy/output/categorization/{phase}_controls.json")
            except Exception as e:
                write_json(manifest_file, {"phase": phase, "signature": signature, "state": "failed", "error": str(e)})
                with (ROOT / "loop/failure.md").open("a", encoding="utf-8") as f:
                    f.write(f"\n- 2026-10-06: ejecucion {phase}: `{type(e).__name__}: {e}`. Pendiente de correccion y nueva validacion.\n")
                raise
        upstream = signature
        if phase == args.through: break


if __name__ == "__main__":
    main()
