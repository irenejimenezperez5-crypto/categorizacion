"""Recover address tags omitted by the original OSM POI cache, using local PBF."""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
import sys
import time
import osmium
import pandas as pd

ROOT=Path(__file__).resolve().parents[2]
CACHE=ROOT/"pharmacy/cache/categorization"
PBF=ROOT/"pharmacy/cache/osm/spain-latest.osm.pbf"


def digest(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(2**20),b''):h.update(b)
    return h.hexdigest()


def prepare():
    target=CACHE/'osm_pharmacy_full_addresses.csv.gz'
    meta=CACHE/'osm_pharmacy_full_addresses_manifest.json'
    source_hash=digest(PBF)
    if target.exists() and meta.exists():
        record=json.loads(meta.read_text(encoding='utf-8'))
        if record.get('source_sha256')==source_hash and record.get('extract_sha256')==digest(target):
            return target,record
    rows=[]
    processor=(osmium.FileProcessor(str(PBF),osmium.osm.NODE)
               .with_filter(osmium.filter.TagFilter(('amenity','pharmacy'))))
    start=time.monotonic()
    print('OSM: reading original pharmacy address tags from local PBF',flush=True)
    for n in processor:
        if not n.location.valid():continue
        tags={tag.k:tag.v for tag in n.tags}
        rows.append({'osm_type':'node','osm_id':n.id,'name':tags.get('name',''),'lat':n.lat,'lon':n.lon,'is_point':1,'tags_json':json.dumps(tags,ensure_ascii=False)})
    data=pd.DataFrame(rows)
    CACHE.mkdir(parents=True,exist_ok=True)
    data.to_csv(target,index=False,compression='gzip',encoding='utf-8')
    record={'source':str(PBF.relative_to(ROOT)),'source_sha256':source_hash,'extract_sha256':digest(target),'node_pharmacies':len(data),'with_street_and_number':sum('addr:street' in json.loads(r) and 'addr:housenumber' in json.loads(r) for r in data.tags_json),'elapsed_seconds':time.monotonic()-start}
    meta.write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(record,ensure_ascii=True),flush=True)
    return target,record


if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    prepare()
