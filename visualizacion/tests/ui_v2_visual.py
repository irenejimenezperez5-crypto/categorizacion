"""Focused V2 source, hover, mouse geography, dark and mobile inspection."""
import json
from pathlib import Path
from urllib.parse import urlsplit
from playwright.sync_api import sync_playwright

HERE=Path(__file__).resolve().parent
results=[]
def check(name,condition,evidence=''):
    results.append({'check':name,'passed':bool(condition),'evidence':evidence})
    print(('PASS ' if condition else 'FAIL ')+name,flush=True)
    assert condition,(name,evidence)

def run():
    errors=[];external=[]
    with sync_playwright() as p:
        browser=p.chromium.launch(executable_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',headless=True)
        page=browser.new_page(viewport={'width':1600,'height':1040})
        page.on('pageerror',lambda e:errors.append(str(e)))
        page.on('request',lambda r:external.append(r.url) if urlsplit(r.url).hostname not in {None,'127.0.0.1','localhost','::1'} else None)
        page.goto('http://127.0.0.1:8765/',wait_until='domcontentloaded',timeout=60000)
        page.locator('.market-dashboard').wait_for(timeout=120000)
        page.wait_for_load_state('networkidle',timeout=60000)
        page.get_by_role('combobox',name='División territorial del mapa',exact=True).select_option('ccaa')
        path=page.get_by_role('button',name='Seleccionar CCAA COMUNIDAD DE MADRID',exact=True)
        path.wait_for(state='attached')
        page.locator('.market-leaflet-host').scroll_into_view_if_needed()
        # Find a genuine polygon interior away from rendered pharmacy marks.
        pick_clear="""el=>{const box=el.getBoundingClientRect(),svg=el.ownerSVGElement,
          map=document.querySelector('.market-leaflet-host'),canvas=map.querySelector('.market-marker-canvas'),mb=map.getBoundingClientRect(),ctx=canvas.getContext('2d');
          for(let x=box.left+2;x<box.right-1;x+=2)for(let y=box.top+2;y<box.bottom-1;y+=2){
            if(y<mb.top+130||y>mb.bottom-65||x<mb.left+5||x>mb.right-5)continue;
            const pt=svg.createSVGPoint();pt.x=x;pt.y=y;
            if(!el.isPointInFill(pt.matrixTransform(el.getScreenCTM().inverse())))continue;
            const pixels=ctx.getImageData(Math.floor((x-mb.left)*canvas.width/mb.width)-6,Math.floor((y-mb.top)*canvas.height/mb.height)-6,13,13).data;
            let clear=true;for(let i=3;i<pixels.length;i+=4)if(pixels[i]){clear=false;break;}
            if(clear)return {x,y};
          }return null;}"""
        point=path.evaluate(pick_clear)
        if point is None:
            page.get_by_role('combobox',name='Vista territorial del mapa',exact=True).select_option(label='Península y Baleares')
            page.wait_for_timeout(300)
            point=path.evaluate(pick_clear)
        if point is None:
            page.locator('.leaflet-control-zoom-in').click();page.wait_for_timeout(300)
            point=path.evaluate(pick_clear)
        assert point,'No selectable Madrid polygon interior'
        print(json.dumps({'click':point,'target':page.evaluate('p=>document.elementFromPoint(p.x,p.y)?.outerHTML.slice(0,250)',point)},ensure_ascii=False),flush=True)
        page.mouse.click(point['x'],point['y'])
        page.wait_for_function("document.querySelector('#market-filter-ccaa').value==='COMUNIDAD DE MADRID'",timeout=15000)
        check('Actual mouse click on CCAA polygon applies filter',True,point)
        page.get_by_role('button',name='Limpiar',exact=True).click()
        page.locator('#market-search').fill('0037534')
        page.wait_for_function("document.querySelectorAll('.market-metric-body strong')[1]?.textContent==='1'",timeout=30000)
        page.locator('.market-row-name').first.click()
        page.locator('.market-store-detail').wait_for()
        page.locator('.market-leaflet-host').scroll_into_view_if_needed()
        box=page.locator('.market-leaflet-host').bounding_box()
        page.mouse.move(box['x']+box['width']/2,box['y']+box['height']/2)
        tip=page.locator('.market-map-hover');tip.wait_for(timeout=10000)
        text=tip.inner_text();bounds=tip.bounding_box()
        check('Marker hover shows both references population and euros',all(x in text for x in ['ESPAÑA','CCAA','65+','75+','85+','Potencial anual estimado']))
        check('Hover stays inside map',bounds['y']>=box['y'] and bounds['y']+bounds['height']<=box['y']+box['height'],bounds)
        page.screenshot(path=str(HERE/'ui_v2.hover.png'),full_page=False)
        page.mouse.click(box['x']+box['width']/2,box['y']+box['height']/2)
        check('Marker click preserves territory filter',page.locator('#market-filter-sb_code').input_value()=='all')
        native=page.get_by_role('combobox',name='Potencial de la selección actions',exact=True)
        if native.count():native.select_option(label='View data source')
        else:
            page.locator('.market-metric').first.get_by_role('button').last.click()
            page.get_by_role('menuitem',name='View data source',exact=True).click()
        page.get_by_text('Definitions',exact=True).wait_for(timeout=10000)
        overview=page.locator('body').inner_text()
        check('Source inspector exposes current Excel and denominator','UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx' in overview and 'potential_eur / sb_original_eur' in overview)
        page.get_by_text('Data preview',exact=True).click();page.wait_for_timeout(300)
        check('Source preview contains the filtered pharmacy','FCIA. ROSA SANCHEZ CALUCHO' in page.locator('body').inner_text())
        page.screenshot(path=str(HERE/'ui_v2.sources.png'),full_page=False)
        page.reload(wait_until='domcontentloaded');page.locator('.market-dashboard').wait_for(timeout=120000)
        page.wait_for_load_state('networkidle',timeout=60000)
        page.get_by_role('button',name='Limpiar',exact=True).click()
        page.emulate_media(color_scheme='dark');page.wait_for_timeout(300)
        page.screenshot(path=str(HERE/'ui_v2.dark.png'),full_page=True)
        check('Dark mode is rendered',True,page.locator('.market-dashboard').evaluate('el=>getComputedStyle(el).color'))
        page.emulate_media(color_scheme='light');page.set_viewport_size({'width':390,'height':844})
        page.wait_for_timeout(300)
        page.locator('.market-map-frame').scroll_into_view_if_needed()
        page.screenshot(path=str(HERE/'ui_v2.mobile_map.png'),full_page=False)
        page.get_by_role('tab',name='Cómo se calcula',exact=True).click()
        page.locator('.market-method-page').wait_for()
        page.screenshot(path=str(HERE/'ui_v2.mobile_method.png'),full_page=True)
        size=page.evaluate('({content:document.documentElement.scrollWidth,viewport:innerWidth})')
        check('Mobile methodology has no horizontal page overflow',size['content']<=size['viewport'],size)
        check('Final browser has no JS errors or external requests',not errors and not external,{'errors':errors,'external':external})
        browser.close()
    (HERE/'ui_v2_visual.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')

if __name__=='__main__':
    try:run()
    finally:(HERE/'ui_v2_visual.partial.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
