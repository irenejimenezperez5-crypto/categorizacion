"""Madrid-only, data-derived double categorization. Production artifacts stay unchanged.

Run with pharmacy/.venv/Scripts/python.exe pharmacy/scripts/madrid_category_pilot.py
The national/regional reference distributions are read nationally; exported points
and original CRM rows are restricted to Comunidad de Madrid.
"""
from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import pickle
import sys

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

from economic_breaks import natural_breaks, assign
from review_categorization import cents_allocation, rank_grades, plain
from export_categorization import Writer, create_previews, numeric_formats, EUR, PCT, POP

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "pharmacy/analysis/madrid_pilot_20261006"
REVIEW = ROOT / "pharmacy/analysis/categorization_review_20261006"
CACHE = ROOT / "pharmacy/cache/categorization"
MADRID = "COMUNIDAD DE MADRID"
AGES = ["pop_65_69", "pop_70_74", "pop_75_79", "pop_80_84", "pop_85_mas"]
CATEGORIES = list("ABCD")
BOOK = OUT / "PILOTO_MADRID_DOBLE_CATEGORIZACION_2027.xlsx"
PRODUCTION = ROOT / "pharmacy/output/archive/2027.1/UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx"
sys.path.insert(0, str(ROOT / "visualizacion/scripts"))
from prepare_data import sha256


def csv(name, frame):
    frame.to_csv(OUT / name, index=False, sep=";", decimal=",", encoding="utf-8-sig")


def save_json(name, data):
    (OUT / name).write_text(json.dumps(plain(data), ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")


def load(name):
    with (CACHE / f"{name}.pickle").open("rb") as file:
        return pickle.load(file)


def distribution(frame, column, money="potential_calibrated_eur", scope="Comunidad de Madrid"):
    part = frame.groupby(column).agg(establishments=("id", "size"), market_eur=(money, "sum"),
                                    min_eur=(money, "min"), median_eur=(money, "median"), max_eur=(money, "max"))
    part = part.reindex(CATEGORIES, fill_value=0).rename_axis("category").reset_index()
    part.insert(0, "scenario", column)
    part.insert(0, "scope", scope)
    part["establishments_pct"] = part.establishments / len(frame)
    part["market_pct"] = part.market_eur / frame[money].sum()
    return part


def matrix(frame, old, new):
    result = pd.crosstab(frame[old], frame[new]).reindex(index=CATEGORIES, columns=CATEGORIES, fill_value=0)
    return result.rename_axis("category_from").reset_index().rename_axis(None, axis=1)


def grouped_bootstrap(frame, base_cuts, iterations, seed):
    """SB cluster-resampling diagnostics of thresholds, NOT forecast confidence intervals."""
    rng = np.random.default_rng(seed)
    groups = [g.potential_calibrated_eur.to_numpy() for _, g in frame.groupby("sb_code", sort=True)]
    values = frame.potential_calibrated_eur.to_numpy()
    original = assign(values, base_cuts)
    draws, counts = [], np.zeros((len(frame), 4), dtype=int)
    for _ in range(iterations):
        sample = np.concatenate([groups[i] for i in rng.integers(0, len(groups), len(groups))])
        fitted = natural_breaks(sample)["cuts_ascending_eur"]
        draws.append(fitted)
        membership = np.searchsorted(fitted, values, side="right")
        for k in range(4):
            counts[:, k] += membership == k
    positions = np.searchsorted(base_cuts, values, side="right")
    match = counts[np.arange(len(frame)), positions] / iterations
    return {"iterations": iterations, "seed": seed, "resampling_unit": "SB de ubicación",
            "p05_cuts_eur": np.quantile(draws, .05, axis=0).tolist(),
            "p50_cuts_eur": np.quantile(draws, .5, axis=0).tolist(),
            "p95_cuts_eur": np.quantile(draws, .95, axis=0).tolist(),
            "interpretation": "Dispersión de cortes al remuestrear SB; no intervalo de ventas ni validación comercial.",
            "draws": draws}, match


def reviewed_bootstrap(frame, base_cuts, iterations, seed, scope):
    cachefile=OUT/"bootstrap_cuts.json"
    manifest=OUT/"sources.json"
    if cachefile.exists() and manifest.exists():
        sources=json.loads(manifest.read_text(encoding="utf-8"))["input_hashes"]
        # The original reference is immutable, even after the local app receives V2.
        # Resolve its historical path through the verified archive without altering
        # the original analytical provenance or recomputing compatible bootstrap draws.
        valid=all(sha256(ROOT/('pharmacy/output/archive/2027.1/stores.json'
            if p.replace('\\','/')=='visualizacion/data/stores.json' else p))==digest
            for p,digest in sources.items())
        info=json.loads(cachefile.read_text(encoding="utf-8"))[scope]
        if valid and info["iterations"]==iterations and info["seed"]==seed:
            values=frame.potential_calibrated_eur.to_numpy()
            original=assign(values,base_cuts)
            match=np.mean([assign(values,draw)==original for draw in info["draws"]],axis=0)
            return info,match
    return grouped_bootstrap(frame,base_cuts,iterations,seed)


def boundary_validation():
    """Check exact global optimization against an independent O(k*n²) DP on small samples."""
    rng = np.random.default_rng(6102026)
    for log in (False, True):
        for _ in range(25):
            raw = np.round(rng.lognormal(7, .8, 30), 2)
            raw = np.r_[raw, raw[:5]]  # duplicate values must never straddle classes
            info = natural_breaks(raw, logarithmic=log)
            x, w = np.unique(raw, return_counts=True)
            x = np.log(x) if log else x
            x = (x - x.mean()) / x.std()
            sw, sx, sq = np.r_[0, np.cumsum(w)], np.r_[0, np.cumsum(w*x)], np.r_[0, np.cumsum(w*x*x)]
            n = len(x)
            dp = np.full((5, n + 1), np.inf)
            dp[0, 0] = 0
            for k in range(1, 5):
                for j in range(k, n + 1):
                    candidates = []
                    for i in range(k - 1, j):
                        cost = sq[j]-sq[i]-(sx[j]-sx[i])**2/(sw[j]-sw[i])
                        candidates.append(dp[k-1, i]+max(float(cost), 0))
                    dp[k, j] = min(candidates)
            expected = 1-dp[4,n]/(sq[-1]-sx[-1]**2/sw[-1])
            assert abs(info["goodness_variance_fit"]-expected) < 1e-10
            cuts = np.asarray(info["cuts_ascending_eur"])
            assert assign(cuts, cuts).tolist() == ["C", "B", "A"]
            assert assign(cuts-.01, cuts).tolist() == ["D", "C", "B"]
    return {"independent_optimization_samples": 50, "inclusive_threshold_cases": 300}


def export_workbook(m, crm, cuts, summary, sb, distribution_frame, matrices, sensitivity, controls):
    writer = Writer()
    explanation = [
        ("Ámbito", "Comunidad de Madrid: 2.947 establecimientos físicos y 2.953 filas CRM; mantener tiendas sanitarias."),
        ("Referencia", "MAT julio 2025–junio 2026, ESSITY + RESTO RX, sell-in por ubicación; sin proyección a 2027."),
        ("Objetivo", "Clasificar potencial de mercado. No es clasificación de rutas, rentabilidad, crecimiento Essity ni ventas observadas por farmacia."),
        ("Mercado piloto", summary["madrid_calibrated_eur"]),
        ("Atribución", "Presupuesto de cada SB de ubicación × peso de población atribuida del establecimiento / suma de pesos del SB. Céntimos por mayores restos, desempate ID."),
        ("Señal demográfica", "65–74 ×1 + 75–84 ×2 + 85+ ×3. Población atribuida y competencia espacial heredadas; residentes fraccionarios modelados."),
        ("Nacional", "Cortes naturales en euros aprendidos sobre 22.316 establecimientos. Letra comparable en toda España usando referencia congelada."),
        ("CCAA", "Mismo algoritmo ajustado a cada CCAA. Letra relativa; los 19 juegos de cortes son referencia analítica, no cambio de producción."),
        ("Método", "Cuatro grupos contiguos ordenados, mínimo global de variación cuadrática en euros dentro de grupo. Sin cuotas de población, farmacias o mercado."),
        ("Reglas SB", "Eliminadas en el piloto: no mínimo C en SB altos ni prohibición de todos A. La V1 original conserva su categoría histórica."),
        ("Cortes editables", "Cortes € contiene límites inferiores inclusivos C/B/A. Editarlos recalcula letras en Puntos/CRM; distribuciones, migraciones, CP y robustez son instantáneas del cálculo entregado. Reentrenar requiere ejecutar el script."),
        ("Incertidumbre", "Ubicaciones provisionales y sensibilidad espacial V1 conservadas. Estabilidad de cortes M1 mide remuestreo de SB; no precisión comercial."),
        ("Euros desde vecinos", "Los importes propio/vecino pertenecen a V1. No sumarlos como descomposición del nuevo potencial calibrado; flujos nuevos no calculados."),
        ("CRM sumable", "Una sola fila CRM por establecimiento tiene contabilizar_modelo=TRUE y mercado_sumable_crm_eur; las otras repiten sus resultados para consulta."),
        ("Alcance", "Piloto separado; Excel nacional, CSV de la aplicación y mapa originales permanecen sin cambio."),
        ("Lectura recomendada", "Usar nacional para valor comparable y CCAA para posición regional. Comparar euros antes de interpretar una letra y revisar puntos sensibles."),
    ]
    writer.table("LEEME", pd.DataFrame(explanation, columns=["tema", "explicacion_o_valor"]),
                 "Madrid · dos lecturas del mismo potencial", "Piloto M1 · 6 octubre 2026 · categorías basadas en datos", widths={"tema":25,"explicacion_o_valor":70},formats={"explicacion_o_valor":EUR},freeze="A5")
    writer.table("Cortes €", cuts, "Cortes económicos aprendidos", "Límites inferiores inclusivos: A ≥ corte B/A; B ≥ corte C/B; C ≥ corte D/C. Referencias congeladas; no son umbrales de rentabilidad.",
                 formats=numeric_formats(cuts), widths={"scope":33,"method":40}, freeze="B5")
    cutcols = {c:get_column_letter(i+1) for i,c in enumerate(cuts.columns)}
    rows = {row.scope: i+5 for i,row in cuts.iterrows()}
    def category_formula(money_col, scope, row):
        cr=rows[scope]
        return f'IF({money_col}{row}>=\'Cortes €\'!${cutcols["A_from_eur"]}${cr},"A",IF({money_col}{row}>=\'Cortes €\'!${cutcols["B_from_eur"]}${cr},"B",IF({money_col}{row}>=\'Cortes €\'!${cutcols["C_from_eur"]}${cr},"C","D")))'

    leading=["id","name","potential_calibrated_eur","category_national","category_ccaa","category_v1","potential_v1_eur",
             "delta_potential_eur","delta_potential_pct","sb_share_calibrated_pct","pop65","pop75","pop85","weighted_population",
             "municipality","postal_code","sb_code","sb_name","sscc","province","ccaa","type","address"]
    rest=[c for c in m if c not in leading]
    points=m[leading+rest].sort_values(["potential_calibrated_eur","id"], ascending=[False,True]).reset_index(drop=True)
    cols={c:get_column_letter(i+1) for i,c in enumerate(points.columns)}
    formulas={
        "category_national":lambda r,c: "="+category_formula(cols["potential_calibrated_eur"],"ESPAÑA",r),
        "category_ccaa":lambda r,c: "="+category_formula(cols["potential_calibrated_eur"],MADRID,r),
        "delta_potential_eur":lambda r,c:f"={cols['potential_calibrated_eur']}{r}-{cols['potential_v1_eur']}{r}",
        "delta_potential_pct":lambda r,c:f"={cols['delta_potential_eur']}{r}/{cols['potential_v1_eur']}{r}",
        "sb_share_calibrated_pct":lambda r,c:f"={cols['potential_calibrated_eur']}{r}/{cols['sb_original_eur']}{r}",
        "pop65":lambda r,c:"="+"+".join(f"{cols[a]}{r}" for a in AGES),
        "pop75":lambda r,c:"="+"+".join(f"{cols[a]}{r}" for a in AGES[2:]),
        "pop85":lambda r,c:f"={cols[AGES[4]]}{r}",
    }
    fmt=numeric_formats(points)
    for c in ["pop65","pop75","pop85","weighted_population"]: fmt[c]=POP
    for c in ["national_break_stability_pct","ccaa_break_stability_pct","sb_share_demographic_pct","share_v1"]:fmt[c]=PCT
    labels={"category_national":"Categoría nacional · M1","category_ccaa":"Categoría CCAA · M1","category_v1":"Categoría final · V1",
            "category_base_v1":"Categoría base sin reglas SB · V1","potential_calibrated_eur":"Potencial calibrado M1 · €",
            "potential_v1_eur":"Potencial original V1 · €","sb_share_calibrated_pct":"% mercado del SB · M1",
            "sensitive_v1":"Sensibilidad espacial · V1","national_break_stability_pct":"% remuestreos conserva letra nacional",
            "ccaa_break_stability_pct":"% remuestreos conserva letra CCAA","pop65":"Población atribuida 65+","pop75":"Población atribuida 75+","pop85":"Población atribuida 85+",
            "own_eur_v1":"Euros propio · V1 exclusivamente","neighbor_eur_v1":"Euros vecinos · V1 exclusivamente"}
    writer.table("Puntos Madrid", points, "Madrid · establecimientos únicos", "Una fila por punto físico. Comparar V1 y M1 sobre el mismo universo. Potencial estimado, población atribuida; códigos como texto.",
                 formulas=formulas,labels=labels,formats=fmt,widths={"name":44,"id":22,"category_national":20,"category_ccaa":20,"category_v1":19,"potential_calibrated_eur":24})

    # CRM retains every original field, with lookup formulas for the comparable metrics.
    crm_formulas={}
    crmcols={c:get_column_letter(i+1) for i,c in enumerate(crm.columns)}
    point_end=len(points)+4
    for c in ["potential_calibrated_eur","category_national","category_ccaa","sb_share_calibrated_pct","pop65","pop75","pop85"]:
        position=cols[c]
        crm_formulas[c]=lambda r,j,p=position:f'=INDEX(\'Puntos Madrid\'!${p}$5:${p}${point_end},MATCH({crmcols["establishment_id"]}{r},\'Puntos Madrid\'!$A$5:$A${point_end},0))'
    crm_formulas["mercado_sumable_crm_eur"]=lambda r,c:f'=IF({crmcols["contabilizar_modelo"]}{r},{crmcols["potential_calibrated_eur"]}{r},0)'
    crmfmt=numeric_formats(crm)
    crmfmt.update({c:POP for c in ["pop65","pop75","pop85","weighted_population"]})
    crmfmt.update({c:PCT for c in ["national_break_stability_pct","ccaa_break_stability_pct","sb_share_demographic_pct","share_v1"]})
    writer.table("CRM Madrid", crm, "Madrid · filas CRM originales", "2.953 filas CRM; preservar los 14 campos originales. Sumar mercado_sumable_crm_eur para evitar doble conteo; fórmulas enlazadas a puntos.",
                 formulas=crm_formulas,labels=labels,formats=crmfmt,widths={"account_name":44,"establishment_id":23})
    writer.table("Distribución",distribution_frame,"Distribución y cobertura de mercado", "Cada escenario conserva el mismo conjunto de establecimientos. V1 se valora con sus euros y M1 con los calibrados.",formats=numeric_formats(distribution_frame),widths={"scenario":40,"scope":29})
    writer.table("Migraciones",matrices,"Cómo cambia cada establecimiento", "Matrices por número de puntos: categoría de origen en filas y nueva en columnas; A/B/C/D ordenadas.",widths={"comparison":40,"scope":30})
    writer.table("SB Madrid",sb,"Conciliación por SB de ubicación", "M1 conserva el sell-in registrado en cada SB. El delta respecto a V1 es una transformación del reparto, no crecimiento observado.",formats=numeric_formats(sb),widths={"sb_name":40})
    cp=points[points.postal_code.eq("28002")].copy()
    # This slice is explicitly a snapshot; category calculations live in Puntos Madrid.
    writer.table("CP 28002",cp,"CP 28002 · detalle comparativo", "40 puntos de tres SB; instantánea M1. Las categorías editables están en Puntos Madrid.",labels=labels,formats=fmt,widths={"name":44})
    writer.table("Robustez",sensitivity,"Sensibilidad de criterios y cortes", "Remuestreo por SB (100 Madrid / 40 nacional), transformación logarítmica y exclusión del 1 % superior. No son intervalos de venta real.",formats=numeric_formats(sensitivity),widths={"scope":30,"scenario":40})
    controls_frame=pd.DataFrame([{"control":k,"result":json.dumps(plain(v),ensure_ascii=False) if isinstance(v,(dict,list)) else v} for k,v in controls.items()])
    writer.table("Controles",controls_frame,"Controles del piloto", "Conservación a céntimos; registros y población; cortes monotónicos; referencia nacional y producción inalterada.",widths={"control":45,"result":65},freeze="B5")
    method=pd.DataFrame([
        {"tema":"Entradas","detalle":"stores.json + market/universe caches F8; extracción ESSITY/RESTO del RX original para control independiente."},
        {"tema":"Presupuesto","detalle":"P_i = M_SB × H_i / sum(H_j en SB ubicación); H_i es población ponderada atribuida. No se reconstruye el origen de euros nuevo."},
        {"tema":"Partición","detalle":"Minimizar Σ_g Σ_i∈g (P_i-media_g)² con cuatro grupos contiguos de potencial; solución exacta DP divide-and-conquer."},
        {"tema":"Empates y cortes","detalle":"Valores iguales siempre misma letra. Corte = punto medio entre máximo de clase inferior y mínimo superior, redondeado hacia arriba al céntimo."},
        {"tema":"Elección de escala","detalle":"Euros como base: el objetivo expresa diferencias absolutas de mercado. Log-euros compara diferencias proporcionales y queda como sensibilidad, sin mezclar objetivos."},
        {"tema":"4 grupos","detalle":"Cuatro clases son una decisión del usuario; los tres importes son aprendidos. No acredita cuatro segmentos comerciales naturales ni retornos distintos."},
        {"tema":"Estabilidad","detalle":"Remuestrear SB con reposición y reaprender cortes; evaluar cada punto original. Los intervalos describen la variación de los cortes, no errores de ventas."},
        {"tema":"IQVIA","detalle":"WELCOME PACK SANIBRICKS 2026, pág. física 4: agrupación de sell-in por ubicación de farmacia; conservar el ámbito del canal original."},
        {"tema":"Método oficial","detalle":"https://doc.esri.com/en/arcgis-pro/latest/help/mapping/layer-properties/data-classification-methods.html · cortes naturales dependen del conjunto de referencia."},
        {"tema":"Referencias pequeñas","detalle":"Ceuta/Melilla y CCAA pequeñas: los cortes son descriptivos y más susceptibles a composición; no interpretar estabilidad no medida como validada."},
    ])
    writer.table("Método y fuentes",method,"Método, fuentes y límites", "Datos privados locales. Los umbrales describen este MAT y universo; congelarlos para comparar periodos sin confundir cambios de corte con cambios de mercado.",widths={"tema":26,"detalle":70},freeze="B5")
    for name in ["Puntos Madrid","CRM Madrid","CP 28002"]:
        frame=writer.frames[name]; ws=writer.book[name]
        for field in ["category_national","category_ccaa","category_v1"]:
            c=get_column_letter(frame.columns.get_loc(field)+1)
            for letter,color in [("A","C6E0D5"),("B","DCEAF3"),("C","FFF0D2"),("D","E8DFEC")]:
                ws.conditional_formatting.add(f"{c}5:{c}{len(frame)+4}",CellIsRule(operator="equal",formula=[f'"{letter}"'],fill=PatternFill("solid",fgColor=color)))
        for field in ["location_provisional","sensitive_v1","near_national_cut_5pct","near_ccaa_cut_5pct"]:
            c=get_column_letter(frame.columns.get_loc(field)+1)
            ws.conditional_formatting.add(f"{c}5:{c}{len(frame)+4}",CellIsRule(operator="equal",formula=["TRUE"],fill=PatternFill("solid",fgColor="FFF0D2")))
    writer.save(BOOK)
    create_previews(BOOK,writer,OUT/"previews")
    # Read actual exported caches and formulas. Every formula must have a cache.
    data=openpyxl.load_workbook(BOOK,read_only=True,data_only=True)
    formula=openpyxl.load_workbook(BOOK,read_only=True,data_only=False)
    count=0
    for name in data.sheetnames:
        for cached, authored in zip(data[name].iter_rows(min_row=5),formula[name].iter_rows(min_row=5)):
            for a,b in zip(cached,authored):
                if b.data_type=="f":
                    count+=1; assert a.value is not None,(name,b.coordinate)
                assert a.data_type != "e",(name,a.coordinate,a.value)
        assert data[name].freeze_panes if not data[name].__class__.__name__.startswith("ReadOnly") else True
    exported=pd.DataFrame(data["Puntos Madrid"].iter_rows(min_row=5,values_only=True),columns=points.columns)
    assert exported.category_national.equals(points.category_national)
    assert exported.category_ccaa.equals(points.category_ccaa)
    assert int(np.rint(exported.potential_calibrated_eur*100).sum())==controls["madrid_market_cents"]
    crm_actual=pd.DataFrame(data["CRM Madrid"].iter_rows(min_row=5,values_only=True),columns=crm.columns)
    assert int(np.rint(crm_actual.mercado_sumable_crm_eur*100).sum())==controls["madrid_market_cents"]
    assert all(isinstance(x,str) for x in exported.sb_code)
    data.close(); formula.close()
    return {"formula_caches_checked":count,"no_excel_error_cells":True,"physical_and_crm_totals":True,"categories_match_csv":True,"workbook_sha256":sha256(BOOK)}


def main():
    OUT.mkdir(parents=True,exist_ok=True)
    production_before=sha256(PRODUCTION)
    stores_path=ROOT/"pharmacy/output/archive/2027.1/stores.json"
    original_stores_hash=sha256(stores_path)
    s=pd.DataFrame(json.loads(stores_path.read_text(encoding="utf-8"))).set_index("id").sort_index()
    market, universe=load("market"),load("universe")
    budgets=market["budgets"]/100
    # Fresh re-computation, checked against raw manufacturer extraction and prior review.
    s["potential_calibrated_eur"]=cents_allocation(s.weighted_population,budgets,s.sb_code)/100
    raw=pd.read_csv(REVIEW/"market_manufacturer_sb.csv",sep=";",decimal=",",dtype={"sb":str}).set_index("sb")
    observed=(raw[["ESSITY_Euros","RESTO_Euros"]].sum(axis=1,min_count=1)*100).round().astype("int64").sort_index()
    assert observed.equals(market["budgets"].sort_index())
    actual=(s.potential_calibrated_eur*100).round().astype("int64").groupby(s.sb_code).sum().sort_index()
    assert actual.equals(observed)
    previous=pd.read_csv(REVIEW/"establishments_review.csv",sep=";",decimal=",",dtype={"id":str}).set_index("id").sort_index()
    assert np.array_equal(np.rint(s.potential_calibrated_eur*100),np.rint(previous.demographic_calibrated_eur*100))
    assert rank_grades(s.potential_eur,s.ccaa).equals(s.category_base)
    optimization_controls=boundary_validation()
    national=natural_breaks(s.potential_calibrated_eur)
    national_log=natural_breaks(s.potential_calibrated_eur,logarithmic=True)
    all_cuts={"ESPAÑA":national}
    for ccaa,g in s.groupby("ccaa",sort=True): all_cuts[ccaa]=natural_breaks(g.potential_calibrated_eur)
    cutrows=[]
    for scope,info in all_cuts.items():
        g=s if scope=="ESPAÑA" else s[s.ccaa.eq(scope)]
        c,b,a=info["cuts_ascending_eur"]
        cutrows.append({"scope":scope,"reference_establishments":len(g),"reference_market_eur":g.potential_calibrated_eur.sum(),
                        "C_from_eur":c,"B_from_eur":b,"A_from_eur":a,"goodness_fit_pct":info["goodness_variance_fit"],
                        "method":"Cortes naturales óptimos sobre euros", "D_n":info["class_counts_D_C_B_A"][0],
                        "C_n":info["class_counts_D_C_B_A"][1],"B_n":info["class_counts_D_C_B_A"][2],"A_n":info["class_counts_D_C_B_A"][3]})
    cuts=pd.DataFrame(cutrows)
    csv("thresholds_reference.csv",cuts)
    madrid=s[s.ccaa.eq(MADRID)].copy()
    national_cuts=national["cuts_ascending_eur"]
    regional_cuts=all_cuts[MADRID]["cuts_ascending_eur"]
    madrid_log=natural_breaks(madrid.potential_calibrated_eur,logarithmic=True)
    s["category_national"]=assign(s.potential_calibrated_eur,national_cuts)
    s["category_ccaa"]=pd.Series(index=s.index,dtype=object)
    for ccaa,g in s.groupby("ccaa",sort=True):s.loc[g.index,"category_ccaa"]=assign(g.potential_calibrated_eur,all_cuts[ccaa]["cuts_ascending_eur"])
    madrid=s.loc[madrid.index].copy()
    print("References fitted. Madrid bootstrap 100; national bootstrap 40.",flush=True)
    regboot, regmatch=reviewed_bootstrap(madrid.reset_index(),regional_cuts,100,20261006,"madrid")
    natboot, natmatch_all=reviewed_bootstrap(s.reset_index(),national_cuts,40,20261007,"national")
    natmatch=pd.Series(natmatch_all,index=s.index).loc[madrid.index].to_numpy()
    m=madrid.reset_index().rename(columns={"potential_eur":"potential_v1_eur","category":"category_v1","category_base":"category_base_v1",
        "category_reason":"category_reason_v1","sb_high":"sb_high_v1","sensitive":"sensitive_v1","sensitivity_variation":"sensitivity_variation_v1",
        "own_eur":"own_eur_v1","neighbor_eur":"neighbor_eur_v1","sb_captured_eur":"sb_captured_eur_v1","share":"share_v1"})
    m["crm_ids"]=m.crm_ids.map(lambda ids: " | ".join(str(item) for item in ids))
    m["category_national_old_euros"]=assign(m.potential_v1_eur,national_cuts)
    m["category_ccaa_old_euros"]=assign(m.potential_v1_eur,regional_cuts)
    m["category_calibrated_cumulative_ccaa"]=rank_grades(m.set_index("id").potential_calibrated_eur).loc[m.id].to_numpy()
    m["category_national_log"]=assign(m.potential_calibrated_eur,national_log["cuts_ascending_eur"])
    m["category_ccaa_log"]=assign(m.potential_calibrated_eur,madrid_log["cuts_ascending_eur"])
    m["category_fixed_40_25_15_reference"]=assign(m.potential_calibrated_eur,[15000,25000,40000])
    m["category_national_quartiles"]=assign(m.potential_calibrated_eur,s.potential_calibrated_eur.quantile([.25,.5,.75]).to_numpy())
    m["category_ccaa_quartiles"]=assign(m.potential_calibrated_eur,m.potential_calibrated_eur.quantile([.25,.5,.75]).to_numpy())
    m["national_break_stability_pct"]=natmatch
    m["ccaa_break_stability_pct"]=regmatch
    m["delta_potential_eur"]=m.potential_calibrated_eur-m.potential_v1_eur
    m["delta_potential_pct"]=m.delta_potential_eur/m.potential_v1_eur
    m["sb_share_calibrated_pct"]=m.potential_calibrated_eur/m.sb_original_eur
    weights=m.groupby("sb_code").weighted_population.transform("sum")
    m["sb_share_demographic_pct"]=m.weighted_population/weights
    assert weights.gt(0).all()
    for prefix,c in [("national",national_cuts),("ccaa",regional_cuts)]:
        distances=np.abs(m.potential_calibrated_eur.to_numpy()[:,None]/np.asarray(c)[None,:]-1)
        m[f"near_{prefix}_cut_5pct"]=distances.min(axis=1)<=.05
    m["dual_category_disagreement"]=m.category_national.ne(m.category_ccaa)
    m["new_market_method"]="SB ubicación calibrado × peso demográfico atribuido; centimos mayores restos"

    crm=universe["crm"]
    crm=crm[crm.establishment_id.isin(m.id)][universe["original_columns"]+["establishment_id"]].copy()
    original_crm=crm[universe["original_columns"]].copy()
    crm=crm.merge(m.drop(columns=["postal_code"]),left_on="establishment_id",right_on="id",how="left",validate="many_to_one",sort=False)
    crm["contabilizar_modelo"]=~crm.establishment_id.duplicated()
    crm["mercado_sumable_crm_eur"]=np.where(crm.contabilizar_modelo,crm.potential_calibrated_eur,0)
    assert crm[universe["original_columns"]].equals(original_crm.reset_index(drop=True))
    assert len(m)==2947 and len(crm)==2953 and m.id.is_unique
    assert int(np.rint(crm.mercado_sumable_crm_eur*100).sum())==int(np.rint(m.potential_calibrated_eur*100).sum())
    assert (m.pop65>=m.pop75).all() and (m.pop75>=m.pop85).all()
    assert np.allclose(m[AGES].sum(axis=1),m.pop65)
    assert np.allclose(m[AGES[2:]].sum(axis=1),m.pop75)
    madrid_sb=m.sb_code.unique()
    assert s[s.sb_code.isin(madrid_sb)].ccaa.eq(MADRID).all()

    scenarios=["category_v1","category_base_v1","category_calibrated_cumulative_ccaa","category_national","category_ccaa",
               "category_national_old_euros","category_ccaa_old_euros","category_national_log","category_ccaa_log","category_fixed_40_25_15_reference",
               "category_national_quartiles","category_ccaa_quartiles"]
    rows=[]
    scopes=[("Comunidad de Madrid",m),("Madrid capital",m[m.municipality_code.eq("28079")]),
            ("Resto Comunidad de Madrid",m[~m.municipality_code.eq("28079")]),("CP 28002",m[m.postal_code.eq("28002")])]
    for scope,g in scopes:
        for column in scenarios:
            money="potential_v1_eur" if column in {"category_v1","category_base_v1","category_national_old_euros","category_ccaa_old_euros"} else "potential_calibrated_eur"
            rows.append(distribution(g,column,money,scope))
    distribution_frame=pd.concat(rows,ignore_index=True)
    matrices=[]
    comparisons=[("category_v1","category_national"),("category_v1","category_ccaa"),("category_national","category_ccaa"),
                 ("category_base_v1","category_calibrated_cumulative_ccaa"),("category_national_old_euros","category_national"),("category_ccaa_old_euros","category_ccaa")]
    for scope,g in scopes:
        for old,new in comparisons:
            t=matrix(g,old,new); t.insert(0,"comparison",old+" → "+new);t.insert(0,"scope",scope);matrices.append(t)
    matrices=pd.concat(matrices,ignore_index=True)
    sb=m.groupby("sb_code",sort=True).agg(sb_name=("sb_name","first"),establishments=("id","size"),
        original_eur=("sb_original_eur","first"),captured_v1_eur=("potential_v1_eur","sum"),captured_m1_eur=("potential_calibrated_eur","sum"),
        provisional_count=("location_provisional","sum"),weighted_population=("weighted_population","sum"),
        national_pattern=("category_national",lambda x:"".join(sorted(set(x)))),ccaa_pattern=("category_ccaa",lambda x:"".join(sorted(set(x))))).reset_index()
    sb["v1_vs_observed_pct"]=sb.captured_v1_eur/sb.original_eur-1
    sb["m1_vs_observed_eur"]=sb.captured_m1_eur-sb.original_eur
    assert np.rint(sb.m1_vs_observed_eur*100).eq(0).all()

    sensitivity_rows=[]
    for scope,f,base in [("ESPAÑA",s,national_cuts),(MADRID,m,regional_cuts)]:
        money=f.potential_calibrated_eur
        for label,values,log in [("Base euros",money,False),("Sensibilidad log-euros",money,True),("Referencia sin 1 % superior",money[money<=money.quantile(.99)],False),("Cuartiles iguales de puntos",money,False)]:
            info=natural_breaks(values,logarithmic=log) if label!="Cuartiles iguales de puntos" else {"cuts_ascending_eur":money.quantile([.25,.5,.75]).to_list(),"goodness_variance_fit":None}
            cuts2=info["cuts_ascending_eur"]
            cat_base=assign(m.potential_calibrated_eur,base)
            cat_new=assign(m.potential_calibrated_eur,cuts2)
            sensitivity_rows.append({"scope":scope,"scenario":label,"C_from_eur":cuts2[0],"B_from_eur":cuts2[1],"A_from_eur":cuts2[2],
                                     "madrid_changes":int((cat_base!=cat_new).sum()),"madrid_changes_pct":float((cat_base!=cat_new).mean()),"goodness_fit_pct":info["goodness_variance_fit"]})
        boot=natboot if scope=="ESPAÑA" else regboot
        for q in ["p05","p50","p95"]:
            c=boot[q+"_cuts_eur"]
            sensitivity_rows.append({"scope":scope,"scenario":"Remuestreo SB "+q,"C_from_eur":c[0],"B_from_eur":c[1],"A_from_eur":c[2],
                                     "madrid_changes":None,"madrid_changes_pct":None,"goodness_fit_pct":None})
    sensitivity_frame=pd.DataFrame(sensitivity_rows)
    audit=[]
    for dim in ["type","degurba","location_provisional","sensitive_v1","municipality"]:
        for value,g in m.groupby(dim,dropna=False):
            for cat in ["category_v1","category_national","category_ccaa"]:
                t=distribution(g,cat,"potential_v1_eur" if cat=="category_v1" else "potential_calibrated_eur",str(value))
                t.insert(0,"dimension",dim);audit.append(t)
    group_audit=pd.concat(audit,ignore_index=True)
    all_A_national=sb.national_pattern.eq("A")
    all_A_ccaa=sb.ccaa_pattern.eq("A")
    summary={"generated_at":datetime.now(timezone.utc).isoformat(),"scope":"Comunidad de Madrid","market_period":"Julio 2025–junio 2026",
        "n_physical":len(m),"n_crm":len(crm),"n_sb":len(sb),"n_capital":int(m.municipality_code.eq("28079").sum()),
        "madrid_v1_eur":round(float(m.potential_v1_eur.sum()),2),"madrid_calibrated_eur":round(float(m.potential_calibrated_eur.sum()),2),
        "national_cuts":national,"madrid_cuts":all_cuts[MADRID],
        "scenarios":{col:{"counts":m[col].value_counts().reindex(CATEGORIES,fill_value=0).to_dict(),
                         "changed_vs_v1":int(m[col].ne(m.category_v1).sum()),
                         "money_column":"potential_v1_eur" if col in {"category_v1","category_base_v1","category_national_old_euros","category_ccaa_old_euros"} else "potential_calibrated_eur",
                         "market_eur":m.groupby(col)["potential_v1_eur" if col in {"category_v1","category_base_v1","category_national_old_euros","category_ccaa_old_euros"} else "potential_calibrated_eur"].sum().reindex(CATEGORIES,fill_value=0).round(2).to_dict()} for col in scenarios},
        "cp28002":{col:m[m.postal_code.eq("28002")][col].value_counts().reindex(CATEGORIES,fill_value=0).to_dict() for col in scenarios},
        "old_sb_overrides":{"D_to_C":int((m.category_base_v1.eq("D")&m.category_v1.eq("C")).sum()),"A_to_B":int((m.category_base_v1.eq("A")&m.category_v1.eq("B")).sum())},
        "dual_disagreement_count":int(m.dual_category_disagreement.sum()),
        "dual_disagreement_market_eur":float(m.loc[m.dual_category_disagreement,"potential_calibrated_eur"].sum()),
        "allocation_only_changes_national_frozen_cuts":int(m.category_national_old_euros.ne(m.category_national).sum()),
        "allocation_only_changes_ccaa_frozen_cuts":int(m.category_ccaa_old_euros.ne(m.category_ccaa).sum()),
        "provisional":int(m.location_provisional.sum()),"sensitive_v1":int(m.sensitive_v1.sum()),
        "near_5pct":{"national":int(m.near_national_cut_5pct.sum()),"ccaa":int(m.near_ccaa_cut_5pct.sum())},
        "bootstrap_less_than_80pct":{"national":int((m.national_break_stability_pct<.8).sum()),"ccaa":int((m.ccaa_break_stability_pct<.8).sum())},
        "all_A_sb":{"national":int(all_A_national.sum()),"ccaa":int(all_A_ccaa.sum())},
        "assumptions":["Universo y edades constantes; el piloto no valida elegibilidad/canal ni ventas por punto.","4 grupos son una decisión de representación, no prueba de 4 segmentos comerciales.","Referencia EUR absoluta; log-EUR alternativa de interpretación, no mejora objetiva comparable por bondad de ajuste.","Congelar cortes para comparaciones temporales; recalcular cada año cambia la escala."],
    }
    controls={"physical_unique":len(m),"crm_rows":len(crm),"crm_original_fields_unchanged":True,"duplicate_extra_crm_rows":len(crm)-len(m),
        "madrid_market_cents":int(np.rint(m.potential_calibrated_eur*100).sum()),"national_reference_market_cents":int(observed.sum()),
        "all_273_location_sb_conserved_cents":True,"raw_RX_and_cache_budgets_match":True,"fresh_calibration_matches_previous_review":True,
        "ccaa_complete_sb_no_border_loss":True,"population_unchanged_by_store":bool(np.allclose(m[AGES],s.loc[m.id,AGES])),
        "ages_nested":True,"all_positive_potentials":bool(m.potential_calibrated_eur.gt(0).all()),"sb_protections_applied":False,
        "same_money_for_both_categories":True,"optimizer_independent_checks":optimization_controls,
        "cent_cut_partition_consistency":True,"production_excel_unchanged":production_before==sha256(PRODUCTION),
        "production_stores_unchanged":original_stores_hash==sha256(stores_path)}
    for info in all_cuts.values():
        assert all(lo<c<=hi for c,(lo,hi) in zip(info["cuts_ascending_eur"],info["boundary_gaps_eur"]))
    for col in ["category_national","category_ccaa"]:
        # Same EUR never carries different categories; higher EUR cannot have a lower class.
        codes=m.sort_values("potential_calibrated_eur")[col].map({"D":0,"C":1,"B":2,"A":3}).to_numpy()
        assert (np.diff(codes)>=0).all()
        assert m.groupby("potential_calibrated_eur")[col].nunique().max()==1
    for name,frame in [("establishments_madrid.csv",m),("crm_madrid.csv",crm),("distributions.csv",distribution_frame),
                       ("migrations.csv",matrices),("sanibricks_madrid.csv",sb),("sensitivity_cuts.csv",sensitivity_frame),
                       ("segments_audit.csv",group_audit),("cp_28002.csv",m[m.postal_code.eq("28002")]),
                       ("dual_category_disagreements.csv",m[m.dual_category_disagreement].sort_values("potential_calibrated_eur",ascending=False))]: csv(name,frame)
    save_json("summary.json",summary)
    save_json("bootstrap_cuts.json",{"national":natboot,"madrid":regboot})
    save_json("sources.json",{"input_hashes":{str(p.relative_to(ROOT)):sha256(p) for p in [stores_path,CACHE/"market.pickle",CACHE/"universe.pickle",REVIEW/"market_manufacturer_sb.csv",REVIEW/"establishments_review.csv",ROOT/"pharmacy/scripts/economic_breaks.py"]},
        "production_sha256":production_before,"method_source":"https://doc.esri.com/en/arcgis-pro/latest/help/mapping/layer-properties/data-classification-methods.html",
        "market_definition":"WELCOME PACK SANIBRICKS 2026.pdf, página física 4","local_only":True})
    controls.update(export_workbook(m,crm,cuts,summary,sb,distribution_frame,matrices,sensitivity_frame,controls))
    assert production_before==sha256(PRODUCTION) and original_stores_hash==sha256(stores_path)
    save_json("controls.json",controls)
    print(json.dumps({"workbook":str(BOOK.relative_to(ROOT)),"madrid":summary["scenarios"],"cp28002":summary["cp28002"],"controls":controls},ensure_ascii=False,default=str),flush=True)


if __name__=="__main__":main()
