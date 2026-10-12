"""Independent, stdlib-only inspection of the actual current Excel, CSV and JSON.

Checks OOXML formula caches and values without relying on the writer/model helpers.
Can run in the production environment or the local QA environment.
"""
from pathlib import Path
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP
import csv
import hashlib
import json
import math
import zipfile
import xml.etree.ElementTree as ET

ROOT=Path(__file__).resolve().parents[2]
OUT=ROOT/'pharmacy/output/current'
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def digest(path):
    with path.open('rb') as file:return hashlib.file_digest(file,'sha256').hexdigest()

def read_csv(path):
    with path.open(encoding='utf-8-sig',newline='') as file:return list(csv.DictReader(file,delimiter=';'))

def value(cell):
    t=cell.get('t');v=cell.find('s:v',NS)
    if t=='inlineStr':return ''.join(cell.itertext())
    if v is None or v.text is None:return None
    if t=='str':return v.text
    if t=='b':return v.text=='1'
    assert t!='e',(cell.get('r'),v.text)
    return float(v.text)

def rows(archive,name):
    with archive.open(name) as file:
        for _,element in ET.iterparse(file,events=['end']):
            if element.tag!=f'{{{NS["s"]}}}row':continue
            record={}
            for c in element.findall('s:c',NS):
                label=''.join(x for x in c.get('r') if x.isalpha())
                col=0
                for x in label:col=col*26+ord(x)-64
                record[col-1]=(value(c),c.find('s:f',NS))
            yield int(element.get('r')),record
            element.clear()

def equal(actual,expected):
    if actual is None:return expected==''
    if isinstance(actual,bool):return expected==str(actual)
    if isinstance(actual,(int,float)):
        try:return math.isclose(actual,float(expected.replace(',','.')),rel_tol=1e-12,abs_tol=1e-8)
        except ValueError:return False
    return actual==expected

def main():
    manifest=json.loads((OUT/'manifest.json').read_text(encoding='utf-8'))
    for name,hashvalue in manifest['files_sha256'].items():assert digest(OUT/name)==hashvalue,name
    for name,hashvalue in manifest['inputs_sha256'].items():assert digest(ROOT/name)==hashvalue,name
    stores=json.loads((OUT/'stores.json').read_text(encoding='utf-8'))
    points={p['id']:p for p in stores}
    archived=json.loads((ROOT/'pharmacy/output/archive/2027.1/stores.json').read_text(encoding='utf-8'))
    assert len(stores)==22316 and len(points)==22316
    ages=['pop_65_69','pop_70_74','pop_75_79','pop_80_84','pop_85_mas']
    for old in archived:
        new=points[old['id']]
        for field in [*ages,'lat','lon','pop65','pop75','pop85','weighted_population']:
            assert new[field]==old[field],(old['id'],field)
    rural_groups=defaultdict(list)
    for p in stores:
        assert p['visit_eligible']==(p['degurba'] in [1,2])
        assert p['geographic_scope']==('No rural' if p['degurba'] in [1,2] else 'Rural' if p['degurba']==3 else 'Sin resolver')
        assert math.isclose(p['allocation_share'],p['allocation_weight']/p['sb_allocation_weight'],rel_tol=1e-12)
        assert math.isclose(p['unrounded_potential_eur'],p['sb_original_eur']*p['allocation_share'],rel_tol=1e-12)
        if p['degurba']==3:rural_groups[p['rural_group_key']].append(p)
        else:assert p['allocation_weight']==p['weighted_population']
    for key,g in rural_groups.items():
        assert max(p['allocation_weight'] for p in g)==min(p['allocation_weight'] for p in g)
        amounts=[round(p['potential_eur']*100) for p in g]
        assert max(amounts)-min(amounts)<=1,key
        assert all(p['rural_group_count']==len(g) for p in g)
    cuts={r['scope']:r for r in json.loads((OUT/'thresholds.json').read_text(encoding='utf-8'))}
    def grade(v,c):return 'A' if v>=c['A_from_eur'] else 'B' if v>=c['B_from_eur'] else 'C' if v>=c['C_from_eur'] else 'D'
    for p in stores:
        assert p['category_national']==grade(p['potential_eur'],cuts['ESPAÑA'])
        assert p['category_ccaa']==grade(p['potential_eur'],cuts[p['ccaa']])
        rounded=float(Decimal(str(p['unrounded_potential_eur'])).quantize(Decimal('.01'),rounding=ROUND_HALF_UP))
        assert round((rounded+p['rounding_adjustment_eur'])*100)==round(p['potential_eur']*100)
    crm=read_csv(OUT/'universo_crm.csv');oldcrm=read_csv(ROOT/'pharmacy/output/archive/2027.1/universo_crm.csv')
    original_fields=list(oldcrm[0])[:14]
    assert len(crm)==len(oldcrm)==22356
    for new,old in zip(crm,oldcrm):
        assert all(new[f]==old[f] for f in original_fields),(new['establishment_id'],'CRM original cambiado')
    book=OUT/'UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx'
    caches=0;checkedcells=0;sheetchecks={};original_cells=0
    with zipfile.ZipFile(book) as archive:
        workbook=ET.fromstring(archive.read('xl/workbook.xml'))
        names=[e.get('name') for e in workbook.findall('s:sheets/s:sheet',NS)]
        for i,sheet in enumerate(names,1):
            path=f'xl/worksheets/sheet{i}.xml'
            records=read_csv(OUT/('establecimientos.csv' if sheet=='Establecimientos' else 'universo_crm.csv')) if sheet in {'Establecimientos','CRM completo'} else None
            count=0
            for n,cells in rows(archive,path):
                if n<5:continue
                count+=1
                expected_values=list(records[n-5].values()) if records is not None else None
                for col,(actual,formula) in cells.items():
                    if formula is not None:
                        caches+=1;assert actual is not None,(sheet,n,col,'sin cache')
                    if records is not None:
                        expected=expected_values[col]
                        assert equal(actual,expected),(sheet,n,col,actual,expected)
                        checkedcells+=1
                        if sheet=='CRM completo' and col<14:original_cells+=1
                if records is not None:
                    assert all(equal(cells.get(j,(None,None))[0],v) for j,v in enumerate(expected_values))
            if records is not None:assert count==len(records)
            sheetchecks[sheet]=count
    current_source=OUT/'stores.json'
    visible_source=ROOT/'visualizacion/data/stores.json'
    if visible_source.exists() and json.loads(visible_source.read_text(encoding='utf-8'))[0].get('model_version')==manifest['version']:
        assert digest(current_source)==digest(visible_source)
        for name in ['universo_crm.csv','establecimientos.csv','sanibricks.csv','sb_prioritarios.csv','thresholds.json','priority_policy.json']:
            assert digest(OUT/name)==digest(ROOT/'visualizacion/data'/name),name
    control={'ready':True,'rural_groups_neutral':len(rural_groups),'rural_euro_range_max_one_cent':True,'rural_not_visitable':True,'workbook_sha256':digest(book),'formula_caches_checked':caches,'excel_csv_cells_checked':checkedcells,
             'crm_original_cells_checked':original_cells,'all_rows_and_original_fields_retained':True,
             'population_and_coordinates_identical_to_v1':True,'all_delivered_formula_caches_present':True,
             'xlsx_and_csv_values_match':True,'inclusive_grades_and_cent_rounding_verified':True,'sheets':sheetchecks,
             'verification_limit':'OOXML inspection and independent arithmetic; no Excel calculation engine invoked.'}
    (OUT/'final_validation.json').write_text(json.dumps(control,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(control,ensure_ascii=False))

if __name__=='__main__':main()
