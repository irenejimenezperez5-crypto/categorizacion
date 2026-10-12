"""Real-browser scope and rural exclusion regression, using local reviewed outputs."""
import csv
from decimal import Decimal
import io
import json
from pathlib import Path
from urllib.parse import urlsplit
from playwright.sync_api import sync_playwright
HERE=Path(__file__).resolve().parent
DATA=HERE.parent/'data'
STORES=json.loads((DATA/'stores.json').read_text(encoding='utf-8'))
results=[]
def check(name,ok,evidence=''):
    results.append({'check':name,'passed':bool(ok),'evidence':evidence})
    print(('PASS ' if ok else 'FAIL ')+name,flush=True)
    (HERE/'ui_rural_acceptance.partial.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    assert ok,(name,evidence)
def count(page,n):
    page.wait_for_function("n=>Number(document.querySelectorAll('.market-metric-body strong')[1]?.textContent.replaceAll('.',''))===n",arg=n,timeout=45000)
def exact_euros(page,rows):
    target=page.evaluate("v=>new Intl.NumberFormat('es-ES',{style:'currency',currency:'EUR',minimumFractionDigits:2,maximumFractionDigits:2}).format(v)",sum(round(r['potential_eur']*100) for r in rows)/100)
    return target in page.locator('.market-metric-body').first.inner_text()
def run():
    errors=[];external=[]
    with sync_playwright() as p:
        b=p.chromium.launch(executable_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',headless=True)
        page=b.new_page(viewport={'width':1600,'height':1040},accept_downloads=True)
        page.on('pageerror',lambda e:errors.append(str(e)))
        page.on('request',lambda r:external.append(r.url) if urlsplit(r.url).hostname not in {None,'localhost','127.0.0.1','::1'} else None)
        page.goto('http://127.0.0.1:8765/',wait_until='domcontentloaded',timeout=60000)
        page.locator('.market-dashboard').wait_for(timeout=120000);count(page,22316)
        scope=page.get_by_role('combobox',name='Ámbito rural / no rural',exact=True)
        panel=page.locator('.market-scope-filter').bounding_box();mapbox=page.locator('.market-map-frame').bounding_box()
        check('Scope is visible in the right panel',panel['x']>mapbox['x']+mapbox['width'])
        scope.select_option('No rural')
        nonrural=[r for r in STORES if r['degurba'] in [1,2]];count(page,len(nonrural))
        check('No rural selects cities and intermediate areas',len(nonrural)==17595 and exact_euros(page,nonrural))
        before=page.locator('.market-metrics').inner_text()
        page.get_by_role('group',name='Referencia de categoría',exact=True).get_by_role('button',name='CCAA',exact=True).click()
        check('Regional reference does not change scope or euros',page.locator('.market-metrics').inner_text()==before and scope.input_value()=='No rural')
        scope.select_option('Rural')
        rural=[r for r in STORES if r['degurba']==3];count(page,len(rural))
        check('Rural remains in the universe with its market',len(rural)==4677 and exact_euros(page,rural))
        with page.expect_download() as event:page.get_by_role('button',name='Descargar selección CSV',exact=True).click()
        path=HERE/'ui_rural.filtered.csv';event.value.save_as(path)
        records=list(csv.DictReader(io.StringIO(path.read_text(encoding='utf-8-sig')),delimiter=';'))
        check('Rural CSV uses the same identities and euros',
            {r['Establecimiento'] for r in records}=={r['id'] for r in rural}
            and sum(Decimal(r['Potencial anual estimado EUR'].replace(',','.')) for r in records)==sum(Decimal(str(r['potential_eur'])) for r in rural))
        check('CSV carries scope eligibility and rural method',all(r['Ámbito rural / no rural']=='Rural' and r['Elegible visita no rural'].lower()=='false' for r in records) and 'Método de reparto' in records[0] and 'Puntos del grupo rural' in records[0])
        target=next(r for r in rural if r['rural_group_count']>1)
        page.locator('#market-search').fill(target['id']);count(page,1)
        page.locator('.market-row-name').first.click()
        detail=page.locator('.market-store-detail');detail.wait_for()
        check('Rural detail shows no visit and equal group allocation','No visitar' in detail.inner_text() and 'reparto igual' in detail.inner_text() and f"Grupo rural: {target['rural_group_count']} puntos" in detail.inner_text())
        page.screenshot(path=str(HERE/'ui_rural.detail.png'),full_page=True)
        page.get_by_role('button',name='Limpiar',exact=True).click();count(page,22316)
        scope.select_option('Sin resolver');count(page,44)
        check('Unresolved scope is separately selectable',exact_euros(page,[r for r in STORES if r['degurba']==0]))
        page.get_by_role('button',name='Limpiar',exact=True).click()
        page.get_by_role('checkbox',name='Solo SB de visita prioritaria',exact=False).check()
        eligible=[r for r in STORES if r['sb_priority_national'] and r['visit_eligible']]
        count(page,len(eligible))
        check('Priority switch excludes rural points even in mixed SB',scope.input_value()=='No rural' and exact_euros(page,eligible))
        scope.select_option('Rural');count(page,0)
        check('Rural plus visit priority never returns rural pharmacies',page.locator('.market-table-empty').is_visible())
        page.get_by_role('button',name='Limpiar',exact=True).click()
        page.set_viewport_size({'width':390,'height':844});scope.select_option('No rural');count(page,17595)
        g=page.evaluate('({w:document.documentElement.scrollWidth,v:innerWidth})')
        check('Mobile scope operates without page overflow',g['w']<=g['v'] and scope.input_value()=='No rural',g)
        page.screenshot(path=str(HERE/'ui_rural.mobile.png'),full_page=True)
        check('No JS errors or third-party requests',not errors and not external,{'errors':errors,'external':external})
        (HERE/'ui_rural_acceptance.json').write_text(json.dumps({'checks':results,'errors':errors,'external':external,'browser':b.version},ensure_ascii=False,indent=2),encoding='utf-8')
        b.close()
if __name__=='__main__':run()
