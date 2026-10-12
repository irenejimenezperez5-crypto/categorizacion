"""Bind reviewed exports to the Data app and build the local, vendored frontend."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
APP = ROOT / "app"
DATA = ROOT / "data"


def write_snapshot(status="complete"):
    metadata = json.loads((DATA / "metadata.json").read_text(encoding="utf-8"))
    stores = json.loads((DATA / "stores.json").read_text(encoding="utf-8"))
    snapshot_path = APP / "src" / "data.json"
    snapshot = json.loads(snapshot_path.read_text(encoding="utf-8"))
    labels = {"geographic_scope":"Ámbito rural / no rural", "ccaa":"Comunidad autónoma", "province": "Provincia", "municipality": "Municipio", "postal_code": "Código postal",
              "sb_code": "Sanibrick", "category_national":"Categoría nacional", "category_ccaa":"Categoría CCAA",
              "sb_priority_national":"Prioridad SB nacional", "dual_category_disagreement":"Comparar letras", "location_provisional": "Ubicación"}
    definitions = [
        {"label": "Establecimientos", "definition": "Puntos físicos únicos del universo CRM; las filas duplicadas no suman dos veces."},
        {"label": "Potencial anual estimado", "definition": "Mercado RX ESSITY + RESTO sell-in, julio 2025–junio 2026, calibrado al SB de ubicación. No son ventas observadas por farmacia.", "formula": "sb_original_eur × allocation_weight / sb_allocation_weight; rural igual municipio–SB; conciliación a céntimos"},
        {"label":"Ruralidad y visita","definition":"DEGURBA original 1/2 No rural visitable, 3 Rural no visitable, 0 sin resolver requiere revisión. Rural: peso igual dentro de municipio–SB; edades originales siguen como contexto."},
        {"label": "Población atribuida", "definition": "Personas repartidas por el modelo, incluidas fracciones. 65+, 75+ y 85+ son acumulados solapados; no se suman entre sí."},
        {"label": "% del SB", "definition": "Potencial de la farmacia dividido por el presupuesto RX observado agregado del SB de ubicación. Las farmacias de cada SB suman el 100 % tras redondear a céntimos.", "formula": "potential_eur / sb_original_eur"},
        {"label": "Categorías nacional y CCAA", "definition": "Cortes aprobados 2027.2 congelados, dos referencias independientes de cortes naturales en euros; grupos contiguos con mínimo global de variación interna. Sin cuotas ni protecciones SB. Límites inferiores incluidos; cortes congelados en thresholds.json."},
        {"label":"Prioridad nacional SB","definition":"Cantidad y euros de A/B nacionales NO RURALES visitables ambos al menos P75 entre todos los SB, rurales con cero. Política de foco explícita; no rentabilidad observada ni rutas."},
        {"label": "Ubicación estimada", "definition": "Proxy espacial cuando no existe coordenada validada; representa una distribución de ubicación y no una dirección precisa."},
        {"label": "Sensibilidad histórica V1", "definition": "Diagnóstico conservado del reparto y categorías anteriores. No valida la estabilidad de las dos categorías actuales ni constituye un intervalo de confianza."},
    ]
    snapshot.update({"surface": "dashboard", "title": "Mapa de potencial de farmacias",
                     "status": "reviewed", "buildStatus": status,
                     "generatedAt": metadata.get("generated_at"), "authoredRevision":2,
                     "filters": [{"id": field, "label": label, "field": "municipality_code" if field == "municipality" else field, "defaultValue": "all",
                                   "queryIds": ["establishments"], "shareInUrl": False}
                                  for field, label in labels.items()],
                     "queries": {"establishments": {"rows": stores,
                         "source": {"provider": "IQVIA, INE y modelo local de atribución",
                             "files": [metadata['excel']['path'], "stores.json", "metadata.json", "export_controls.json", "sanibricks.geojson", "ccaa.geojson", "thresholds.json", "priority_policy.json"],
                             "metricDefinitions": definitions,
                             "evidenceFlow": [
                                 {"title": "Fuente", "detail": "Excel nacional categorizado; huella y controles guardados en /data/metadata.json."},
                                 {"title": "Exportación", "detail": "scripts/prepare_data.py reutiliza resultados conciliados y conserva identificadores y coordenadas validadas/proxy."},
                                 {"title": "Calibración V2", "detail": "Cada SB de ubicación conserva su presupuesto observado agregado, repartido por peso aplicado: individual no rural y promedio municipal–SB rural sin distancia. Edades V1 conservadas. Los flujos de euros V1 son históricos y no descomponen el resultado V2."},
                                 {"title": "Visualización", "detail": "Un marcador por establecimiento físico; todos los indicadores, fichas y CSV de selección usan la misma intersección de filtros."}],
                             },
                         "methods": [{"language": "Python", "code": "python visualizacion/scripts/prepare_data.py"}]}}})
    # Additional reviewed queries give SB priority and cut tables their own inspectable evidence.
    for query_id,filename in [('sanibricks','sanibricks.json'),('thresholds','thresholds.json'),('priority_policy','priority_policy.json')]:
        rows=json.loads((DATA/filename).read_text(encoding='utf-8'))
        snapshot['queries'][query_id]={'rows':rows if isinstance(rows,list) else [rows],
             'source':{'provider':'Modelo local V2 sobre IQVIA e INE','files':[filename,metadata['excel']['path']],
                       'metricDefinitions':definitions,'executedAt':metadata.get('generated_at')}}
    snapshot['title']='Potencial de farmacias · España y CCAA'
    snapshot['commercialReport']=True
    snapshot_path.write_text(json.dumps(snapshot, ensure_ascii=False, separators=(",", ":"), allow_nan=False), encoding="utf-8")
    return len(stores)


def find_node(explicit):
    if explicit:
        return Path(explicit)
    installed = shutil.which("node") or os.environ.get("DATA_APP_NODE")
    if installed:
        return Path(installed)
    candidates = list((Path(os.environ.get("LOCALAPPDATA", "")) / "OpenAI" / "Codex" / "runtimes").glob("**/bin/node.exe"))
    if candidates:
        return sorted(candidates)[-1]
    raise RuntimeError("Node no está disponible para recompilar. Usa --node RUTA o DATA_APP_NODE; arrancar la app compilada no necesita Node.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--refresh-data", action="store_true")
    parser.add_argument("--snapshot-only", action="store_true")
    parser.add_argument("--node")
    parser.add_argument('--source',action='store_true',help='Compile an explicitly authorized scoped runtime change with the existing local toolchain.')
    parser.add_argument("--status", choices=["creating", "updating", "complete", "paused"], default="complete")
    args = parser.parse_args()
    if args.refresh_data:
        subprocess.run([sys.executable, str(ROOT / "scripts" / "prepare_data.py")], check=True, cwd=ROOT.parent)
    count = write_snapshot(args.status)
    print(f"Snapshot enlazado: {count:,} establecimientos", flush=True)
    if not args.snapshot_only:
        plugin = Path.home() / ".codex" / "plugins" / "cache" / "openai-curated-remote" / "data-analytics"
        scripts = sorted(plugin.glob("*/scripts/data-app.mjs"))
        if not scripts:
            raise RuntimeError("Falta el compilador del plugin Data Analytics. La aplicación ya compilada funciona sin este plugin.")
        source_build=args.source or json.loads((APP/'src/data.json').read_text(encoding='utf-8')).get('commercialReport',False)
        # The installed compiler keeps source and prebuilt profiles distinct:
        # split reviewed data is available only in the prebuilt profile.
        profile=['--source'] if source_build else ['--separate-data']
        subprocess.run([str(find_node(args.node)), str(scripts[-1]), "build", "--project-dir", str(APP), *profile], check=True)


if __name__ == "__main__":
    main()
