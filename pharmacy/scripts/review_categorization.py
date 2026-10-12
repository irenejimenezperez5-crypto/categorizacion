"""Read-only business review; scenarios never replace the approved F8 model."""
from __future__ import annotations
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import pickle

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "pharmacy/analysis/categorization_review_20261006"
CACHE = ROOT / "pharmacy/cache/categorization"
AGES = ["pop_65_69", "pop_70_74", "pop_75_79", "pop_80_84", "pop_85_mas"]


def load(name):
    with (CACHE / f"{name}.pickle").open("rb") as file:
        return pickle.load(file)


def plain(value):
    if isinstance(value, dict):
        return {str(k): plain(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, np.ndarray)):
        return [plain(v) for v in value]
    if isinstance(value, np.generic):
        return value.item()
    return value


def save_json(name, value):
    (OUT / name).write_text(json.dumps(plain(value), ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")


def save_csv(name, frame):
    frame.to_csv(OUT / name, index=False, sep=";", decimal=",", encoding="utf-8-sig")


def rank_grades(values, groups=None):
    """Same cumulative-euro policy, independently implemented without production helpers."""
    groups = pd.Series("España", index=values.index) if groups is None else groups
    result = pd.Series("D", index=values.index)
    for group in pd.unique(groups):
        part = values.loc[groups.eq(group)].sort_values(ascending=False, kind="stable")
        previous = (part.cumsum() - part) / part.sum()
        result.loc[part.index] = previous.map(lambda q: "A" if q < .2 else "B" if q < .5 else "C" if q < .75 else "D")
    return result


def cents_allocation(values, budgets, sb_codes):
    """Allocate actual location-SB budgets with largest remainders, stable ID tie-breaks."""
    result = pd.Series(0, index=values.index, dtype="int64")
    for sb, group in values.groupby(sb_codes, sort=True):
        money = int(round(budgets.loc[sb] * 100))
        weights = group.to_numpy(float)
        if weights.sum() <= 0:
            weights = np.ones(len(group))
        raw = money * weights / weights.sum()
        rounded = np.floor(raw).astype("int64")
        remainder = money - int(rounded.sum())
        order = np.argsort(-(raw - rounded), kind="stable")
        rounded[order[:remainder]] += 1
        result.loc[group.index] = rounded
    return result


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    stores_path = ROOT / "pharmacy/output/archive/2027.1/stores.json"
    s = pd.DataFrame(json.loads(stores_path.read_text(encoding="utf-8"))).set_index("id").sort_index()
    market, attribution, sensitivity = load("market"), load("attribution"), load("sensitivity")
    budgets = market["budgets"] / 100
    pop = market["sb_population"].copy()
    pop["pop65"] = pop[AGES].sum(axis=1)
    pop["pop75"] = pop[AGES[2:]].sum(axis=1)
    pop["weighted"] = pop[AGES[:2]].sum(axis=1) + 2 * pop[AGES[2:4]].sum(axis=1) + 3 * pop[AGES[4]]
    sb = s.groupby("sb_code").agg(n=("potential_eur", "size"), captured_eur=("potential_eur", "sum"),
        provisional_share=("location_provisional", "mean"), name=("sb_name", "first"), ccaa=("ccaa", "first"),
        category_pattern=("category", lambda x: "".join(sorted(set(x)))), all_CD=("category", lambda x: x.isin(["C", "D"]).all()))
    sb["original_eur"] = budgets
    sb["original_per_store_eur"] = sb.original_eur / sb.n
    sb["captured_per_store_eur"] = sb.captured_eur / sb.n
    sb["net_change_eur"] = sb.captured_eur - sb.original_eur
    sb["net_change_pct"] = sb.net_change_eur / sb.original_eur
    sb["pop65_origin"] = pop.pop65
    sb["weighted_origin"] = pop.weighted
    sb["original_eur_per_65"] = sb.original_eur / sb.pop65_origin
    sb["original_eur_per_weighted"] = sb.original_eur / sb.weighted_origin
    sb["high_sb"] = sb.original_eur.ge(budgets.quantile(.75))
    save_csv("sanibricks_diagnostic.csv", sb.reset_index())

    # Independent reconstruction of the approved base classification.
    reconstructed = rank_grades(s.potential_eur, s.ccaa)
    assert reconstructed.eq(s.category_base).all(), "Independent base classification differs"
    national_cumulative = rank_grades(s.potential_eur)
    s["category_national_cumulative"] = national_cumulative

    # Scenarios isolate allocation and classification changes. Neither is a forecast or validated sales estimate.
    s["location_calibrated_eur"] = cents_allocation(s.potential_eur, budgets, s.sb_code) / 100
    s["demographic_calibrated_eur"] = cents_allocation(s.weighted_population, budgets, s.sb_code) / 100
    for fraction in [.25, .5]:
        signal = s.demographic_calibrated_eur * (1 - fraction) + s.sb_code.map(budgets / sb.n) * fraction
        column = f"demographic_shrunk_{int(fraction*100)}_eur"
        s[column] = cents_allocation(signal, budgets, s.sb_code) / 100
        s[f"category_shrunk_{int(fraction*100)}"] = np.select([s[column] >= limit for limit in (40000, 25000, 15000)], ["A", "B", "C"], default="D")
    for column in ["location_calibrated_eur", "demographic_calibrated_eur", "demographic_shrunk_25_eur", "demographic_shrunk_50_eur"]:
        calculated = (s[column] * 100).round().astype("int64").groupby(s.sb_code).sum()
        assert calculated.equals(market["budgets"].sort_index()), column
    s["category_calibrated_regional"] = rank_grades(s.location_calibrated_eur, s.ccaa)
    s["category_demographic_regional"] = rank_grades(s.demographic_calibrated_eur, s.ccaa)

    # Illustrative common EUR bands, not selected thresholds. Sensitivity explicitly compares alternatives.
    for name, cutoffs in [("bands_40_25_15", (40000, 25000, 15000)),
                          ("bands_35_22_12", (35000, 22000, 12000)),
                          ("bands_50_30_15", (50000, 30000, 15000))]:
        s[name] = np.select([s.potential_eur >= limit for limit in cutoffs], ["A", "B", "C"], default="D")
        s[f"calibrated_{name}"] = np.select([s.demographic_calibrated_eur >= limit for limit in cutoffs], ["A", "B", "C"], default="D")
    scenario_columns = ["category", "category_base", "category_national_cumulative", "category_calibrated_regional", "category_demographic_regional", "bands_40_25_15", "bands_35_22_12", "bands_50_30_15", "calibrated_bands_40_25_15", "category_shrunk_25", "category_shrunk_50"]
    scenario_value = {column: "potential_eur" for column in scenario_columns}
    scenario_value.update({"category_calibrated_regional": "location_calibrated_eur", "category_demographic_regional": "demographic_calibrated_eur", "calibrated_bands_40_25_15": "demographic_calibrated_eur", "category_shrunk_25": "demographic_shrunk_25_eur", "category_shrunk_50": "demographic_shrunk_50_eur"})
    for column in scenario_columns:
        distribution = s.groupby(column).agg(n=("potential_eur", "size"), market_eur_scenario=(scenario_value[column], "sum"), market_eur_v1_same_group=("potential_eur", "sum")).reset_index()
        save_csv(f"distribution_{column}.csv", distribution)
    save_csv("establishments_review.csv", s.reset_index())
    cp = s[s.postal_code.eq("28002")].copy()
    save_csv("cp_28002.csv", cp.sort_values("potential_eur", ascending=False).reset_index())
    affected = sb.loc[cp.sb_code.unique()].copy()
    affected["n_stores_cp"] = cp.groupby("sb_code").size()
    affected["market_stores_cp"] = cp.groupby("sb_code").potential_eur.sum()
    save_csv("cp_28002_sanibricks.csv", affected.reset_index())

    # Region thresholds and policy overrides. Base bands are actually different currencies of meaning by region.
    region_rows = []
    for region, group in s.groupby("ccaa"):
        row = {"ccaa": region, "n": len(group), "market_eur": group.potential_eur.sum(), "median_eur": group.potential_eur.median(),
               "provisional_share": group.location_provisional.mean(), "sensitive_share": group.sensitive.mean()}
        for category in "ABCD":
            points = group[group.category_base.eq(category)]
            row[f"base_{category}_n"] = len(points)
            row[f"base_{category}_min_eur"] = points.potential_eur.min() if len(points) else None
            row[f"base_{category}_max_eur"] = points.potential_eur.max() if len(points) else None
        region_rows.append(row)
    regions = pd.DataFrame(region_rows)
    save_csv("ccaa_thresholds.csv", regions)
    overrides = s[s.category.ne(s.category_base)]
    save_csv("policy_overrides.csv", overrides.reset_index())
    save_csv("category_overlap_by_quality.csv", s.groupby(["rurality", "location_provisional", "category"]).agg(n=("potential_eur", "size"), mean=("potential_eur", "mean"), median=("potential_eur", "median"), eur=("potential_eur", "sum")).reset_index())

    # Attach sensitivity scenarios already calculated with the approved origin budgets.
    detail = sensitivity["detail"].set_index("establishment_id")
    save_csv("cp_28002_sensitivity.csv", detail.loc[cp.index].reset_index())
    save_csv("sensitivity_scenarios.csv", sensitivity["scenarios"])
    # The national catchment-to-sales relationship is generated by the same model, not external evidence.
    correlations = {"national_pearson_eur_weighted_population": s.potential_eur.corr(s.weighted_population),
                    "national_spearman_eur_weighted_population": s.potential_eur.corr(s.weighted_population, method="spearman"),
                    "cp28002_pearson_eur_pop65": cp.potential_eur.corr(cp.pop65)}

    # Origin concentration informs whether geographic precision is justified.
    flow = attribution["flow_cents"].tocsc()
    concentration = []
    for index in range(flow.shape[1]):
        values = flow.data[flow.indptr[index]:flow.indptr[index + 1]]
        concentration.append(float(values.max() / values.sum()) if len(values) else 0)
    uncertainty = {
        "provisional_count": int(s.location_provisional.sum()), "provisional_share": float(s.location_provisional.mean()),
        "provisional_market_share": float(s.loc[s.location_provisional, "potential_eur"].sum() / s.potential_eur.sum()),
        "sensitive_count": int(s.sensitive.sum()), "sensitive_share": float(s.sensitive.mean()),
        "urban_provisional_share": float(s[s.degurba.eq(1)].location_provisional.mean()),
        "median_dominant_origin_share": float(np.median(concentration)),
        "sensitive_by_quality": s.groupby("location_provisional").sensitive.mean().to_dict()}
    region_amounts = pop.groupby("ccaa").sum(numeric_only=True)
    region_amounts["original_eur"] = budgets.groupby(pop.ccaa).sum()
    region_amounts["original_eur_per_65"] = region_amounts.original_eur / region_amounts.pop65
    region_amounts["original_eur_per_weighted"] = region_amounts.original_eur / region_amounts.weighted
    save_csv("ccaa_market_population.csv", region_amounts.reset_index())
    controls = {"base_category_reconstruction": True, "location_sb_conserved_cents": True,
                "demographic_sb_conserved_cents": True, "physical_stores": len(s),
                "national_market_cents": int(np.rint(s.potential_eur * 100).sum()),
                "original_market_cents": int(market["budgets"].sum())}
    assert controls["national_market_cents"] == controls["original_market_cents"]
    input_hashes = {str(stores_path.relative_to(ROOT)): hashlib.sha256(stores_path.read_bytes()).hexdigest(),
        **{str((CACHE / f"{name}.pickle").relative_to(ROOT)): hashlib.sha256((CACHE / f"{name}.pickle").read_bytes()).hexdigest() for name in ["market", "attribution", "sensitivity"]}}
    metadata = json.loads((ROOT / "pharmacy/output/archive/2027.1/metadata.json").read_text(encoding="utf-8"))
    source_hashes = {item["path"]: item["sha256"] for item in metadata["sources"]}
    for name in ["market", "attribution"]:
        key = str((CACHE / f"{name}.pickle").relative_to(ROOT))
        assert input_hashes[key] == source_hashes[key.replace("\\", "/")]
    expected_stores = next(item["sha256"] for item in metadata["output_files"] if item["path"] == "stores.json")
    assert input_hashes[str(stores_path.relative_to(ROOT))] == expected_stores
    controls["export_manifest_matches_inputs"] = True
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(), "scope": "Review and unapproved scenarios; original model unchanged",
        "controls": controls, "cp28002": {"n": len(cp), "sb_count": cp.sb_code.nunique(),
            "category_counts": cp.category.value_counts().to_dict(), "base_counts": cp.category_base.value_counts().to_dict(),
            "market_eur": cp.potential_eur.sum(), "min_eur": cp.potential_eur.min(), "max_eur": cp.potential_eur.max(),
            "mean_eur": cp.potential_eur.mean(), "median_eur": cp.potential_eur.median(),
            "median_regional_percentile": float(cp.potential_eur.median() >= 0 and (s[s.ccaa.eq("COMUNIDAD DE MADRID")].potential_eur <= cp.potential_eur.median()).mean()),
            "provisional": int(cp.location_provisional.sum()), "sensitive": int(cp.sensitive.sum()),
            "scenarios": {column: cp[column].value_counts().to_dict() for column in scenario_columns},
            "scenario_market_eur": {column: float(cp[scenario_value[column]].sum()) for column in scenario_columns}},
        "allocation": {"gross_cross_sb_eur": float(s.neighbor_eur.sum()),
            "net_territorial_redistribution_eur": float(sb.net_change_eur.abs().sum() / 2),
            "sb_more_than_20pct_net_change": int(sb.net_change_pct.abs().gt(.2).sum()),
            "sb_more_than_50pct_net_change": int(sb.net_change_pct.abs().gt(.5).sum()),
            "sb_all_CD": int(sb.all_CD.sum())},
        "policy": {"base_counts": s.category_base.value_counts().to_dict(), "final_counts": s.category.value_counts().to_dict(),
            "overrides": len(overrides), "high_sb_threshold_eur": budgets.quantile(.75),
            "high_sb_count": int(sb.high_sb.sum()), "low_value_promotions_below_10000": int((overrides.category_base.eq("D") & overrides.potential_eur.lt(10000)).sum()),
            "national_cumulative_changes_vs_base": int(s.category_national_cumulative.ne(s.category_base).sum())},
        "scenarios": {column: {"n": s[column].value_counts().to_dict(), "market_by_category": s.groupby(column)[scenario_value[column]].sum().to_dict(), "changes_vs_current": int(s[column].ne(s.category).sum()), "changes_vs_base": int(s[column].ne(s.category_base).sum())} for column in scenario_columns},
        "uncertainty": uncertainty, "correlations": correlations,
        "types": s.groupby("type").agg(n=("potential_eur", "size"), eur=("potential_eur", "sum")).reset_index().to_dict("records"),
        "input_hashes": input_hashes,
    }
    save_json("summary.json", summary)
    print(json.dumps(plain({"cp28002": summary["cp28002"], "allocation": summary["allocation"], "policy": summary["policy"], "uncertainty": uncertainty, "correlations": correlations, "types": summary["types"]}), ensure_ascii=True, indent=2), flush=True)


if __name__ == "__main__":
    main()
