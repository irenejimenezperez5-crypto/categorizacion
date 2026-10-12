"""Local OSM pedestrian-network contrast. No CRM coordinates leave the project."""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
import time

import networkx as nx
import numpy as np
import pandas as pd
from scipy import sparse
from scipy.spatial import cKDTree
import osmium

from categorization_core import age_demand, haversine_matrix


def audit_walking(stores, sections, pilot, cache: Path, output: Path):
    pbf = cache.parents[0] / "osm/spain-latest.osm.pbf"
    if not pbf.exists():
        pbf = cache.parent / "osm/spain-latest.osm.pbf"
    matrix = sparse.load_npz(cache / "shares_base.npz").tocsc()
    known = stores.latitude_validada.notna() & stores.degurba.isin([1,2])
    pilot_ids = set(pilot.loc[pilot.estrato.eq("urbano"),"establishment_id"])
    selected = list(stores.index[known & stores.establishment_id.isin(pilot_ids)])
    if len(selected)<3:
        selected += list(stores.index[known & ~stores.index.isin(selected)][:3-len(selected)])
    selected = selected[:3]
    routes, boxes = [], []
    known_idx = np.flatnonzero(known.values)
    known_ll = stores.loc[known,["longitude_validada","latitude_validada"]].to_numpy(float)
    for case,j in enumerate(selected,1):
        indices = matrix.indices[matrix.indptr[j]:matrix.indptr[j+1]]
        values = matrix.data[matrix.indptr[j]:matrix.indptr[j+1]]
        urban = sections.loc[indices,"degurba"].ne(3).values
        indices,values = indices[urban],values[urban]
        if not len(indices): continue
        ll = sections.loc[indices,["centroide_lon","centroide_lat"]].values
        dist = haversine_matrix(np.array([[stores.at[j,"longitude_validada"],stores.at[j,"latitude_validada"]]]),ll).ravel()
        feasible = (dist>=100.)&(dist<=1500.)
        if feasible.any():
            score = values*age_demand(sections.loc[indices, ["pop_65_69","pop_70_74","pop_75_79","pop_80_84","pop_85_mas"]].values)
            score[~feasible] = -1.
            sec_idx = indices[np.argmax(score)]
        else:
            sec_idx = indices[np.argmax(values)]
        origin = sections.loc[sec_idx,["centroide_lon","centroide_lat"]].to_numpy(float)
        distances = haversine_matrix(origin[None,:],known_ll).ravel()
        closest = known_idx[np.argsort(distances,kind="stable")[:2]]
        box = [origin[0]-.04,origin[1]-.035,origin[0]+.04,origin[1]+.035]
        boxes.append(box)
        for k in closest:
            dest = np.array([stores.at[k,"longitude_validada"],stores.at[k,"latitude_validada"]])
            routes.append({"caso":case,"sscc_origen":sections.at[sec_idx,"sscc"],"establecimiento":stores.at[k,"establishment_id"],"origen_lon":origin[0],"origen_lat":origin[1],"destino_lon":dest[0],"destino_lat":dest[1],"distancia_geografica_m":float(haversine_matrix(origin[None,:],dest[None,:])[0,0])})
    target = output/"recorridos_pie.csv"
    signature = hashlib.sha256(json.dumps({"routes":routes,"pbf_bytes":pbf.stat().st_size,"pbf_mtime":pbf.stat().st_mtime_ns,"code":Path(__file__).read_text(encoding="utf-8")},sort_keys=True).encode()).hexdigest()
    meta_file = output/"recorridos_pie_manifest.json"
    if target.exists() and meta_file.exists() and json.loads(meta_file.read_text(encoding="utf-8")).get("signature")==signature:
        return pd.read_csv(target)
    graph = nx.Graph()
    started,last = time.monotonic(),time.monotonic()
    scanned = 0
    processor = (osmium.FileProcessor(str(pbf)).with_locations("flex_mem")
                 .with_filter(osmium.filter.EntityFilter(osmium.osm.WAY))
                 .with_filter(osmium.filter.KeyFilter("highway")))
    excluded = {"motorway","motorway_link","trunk","trunk_link","construction","proposed","raceway"}
    for way in processor:
        scanned += 1
        hw = way.tags.get("highway")
        if hw in excluded or way.tags.get("foot") in {"no","private"}:
            continue
        if way.tags.get("access") in {"no","private"} and way.tags.get("foot") not in {"yes","designated","permissive"}:
            continue
        nodes = []
        # Most ways are rejected by their first point; bounding boxes include a generous margin.
        if not way.nodes: continue
        first = way.nodes[0]
        if not first.location.valid(): continue
        lon,lat = first.lon,first.lat
        if not any(a<=lon<=c and b<=lat<=d for a,b,c,d in boxes): continue
        for n in way.nodes:
            if n.location.valid(): nodes.append((n.ref,n.lon,n.lat))
        if len(nodes)>=2:
            ll = np.array([[n[1],n[2]] for n in nodes])
            for a,b in zip(nodes[:-1],nodes[1:]):
                length = float(haversine_matrix(np.array([[a[1],a[2]]]),np.array([[b[1],b[2]]]))[0,0])
                graph.add_node(a[0],lon=a[1],lat=a[2])
                graph.add_node(b[0],lon=b[1],lat=b[2])
                if not graph.has_edge(a[0],b[0]) or length<graph[a[0]][b[0]]["length"]:
                    graph.add_edge(a[0],b[0],length=length)
        if time.monotonic()-last>25:
            print(f"F6 walking: {scanned:,} highway ways scanned; {graph.number_of_edges():,} local edges",flush=True)
            last=time.monotonic()
    node_ids = list(graph.nodes)
    ll = np.array([[graph.nodes[n]["lon"],graph.nodes[n]["lat"]] for n in node_ids])
    for route in routes:
        if not len(ll):
            route["estado"]="SIN_RED_OSM"
            continue
        o = np.array([[route["origen_lon"],route["origen_lat"]]])
        d = np.array([[route["destino_lon"],route["destino_lat"]]])
        osnap = haversine_matrix(o,ll).ravel()
        dsnap = haversine_matrix(d,ll).ravel()
        a,b = int(np.argmin(osnap)),int(np.argmin(dsnap))
        route["acceso_origen_m"],route["acceso_destino_m"] = float(osnap[a]),float(dsnap[b])
        try:
            walked = nx.shortest_path_length(graph,node_ids[a],node_ids[b],weight="length")+osnap[a]+dsnap[b]
            route["recorrido_pie_m"] = float(walked)
            route["ratio_pie_geografica"] = float(walked/max(1.,route["distancia_geografica_m"]))
            route["estado"] = "OK" if max(osnap[a],dsnap[b])<=150. else "ACCESO_RED_MAYOR_150M"
        except nx.NetworkXNoPath:
            route["estado"] = "SIN_CAMINO: red local o barrera; revisar"
        route["metodo"] = "Red OSM local no dirigida; vias a pie permitidas; acceso al nodo mas cercano; contraste, no cambio del modelo"
    result = pd.DataFrame(routes)
    result.to_csv(target,index=False,encoding="utf-8-sig")
    meta_file.write_text(json.dumps({"signature":signature,"source":str(pbf),"network_nodes":graph.number_of_nodes(),"network_edges":graph.number_of_edges(),"scanned_highways":scanned,"elapsed_seconds":time.monotonic()-started},indent=2),encoding="utf-8")
    return result
