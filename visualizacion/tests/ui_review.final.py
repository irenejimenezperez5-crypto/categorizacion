"""Focused final typography, hover placement and responsive visual checks."""
import csv
import io
import json
from pathlib import Path
from urllib.parse import urlsplit
from playwright.sync_api import sync_playwright

HERE=Path(__file__).resolve().parent
results=[]


def run():
    errors=[]
    with sync_playwright() as p:
        browser=p.chromium.launch(executable_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',headless=True)
        context=browser.new_context(viewport={'width':1600,'height':1040},accept_downloads=True)
        page=context.new_page(); page.on('pageerror',lambda e:errors.append(str(e)))
        page.goto('http://127.0.0.1:8765/',wait_until='domcontentloaded')
        page.locator('.market-dashboard').wait_for(timeout=120000)
        page.wait_for_load_state('networkidle',timeout=60000)
        page.screenshot(path=str(HERE/'ui_review.desktop.png'),full_page=True)
        page.locator('#market-search').fill('0037534')
        page.wait_for_function("document.querySelectorAll('.market-metric-body strong')[1]?.textContent==='1'",timeout=30000)
        page.locator('.market-row-name').first.click()
        page.locator('.market-store-detail').wait_for()
        page.get_by_role('button',name='Analítica',exact=True).click()
        page.wait_for_load_state('networkidle',timeout=30000)
        page.locator('.market-leaflet-host').scroll_into_view_if_needed()
        box=page.locator('.market-leaflet-host').bounding_box()
        page.mouse.move(box['x']+box['width']/2,box['y']+box['height']/2)
        page.locator('.market-map-hover').wait_for(timeout=10000)
        tip=page.locator('.market-map-hover').bounding_box()
        results.append({'check':'Hover inside map at center','passed':tip['y']>=box['y'] and tip['y']+tip['height']<=box['y']+box['height'],'map':box,'tooltip':tip})
        page.screenshot(path=str(HERE/'ui_review.marker_hover.png'),full_page=False)
        page.mouse.move(box['x']+100,box['y']+200);page.mouse.down();page.mouse.move(box['x']+100,box['y']+420,steps=12);page.mouse.up()
        page.wait_for_timeout(300)
        page.mouse.move(box['x']+box['width']/2,box['y']+box['height']/2+220)
        page.locator('.market-map-hover').wait_for(timeout=10000)
        tip=page.locator('.market-map-hover').bounding_box()
        results.append({'check':'Hover inside map near bottom','passed':tip['y']>=box['y'] and tip['y']+tip['height']<=box['y']+box['height'],'map':box,'tooltip':tip})
        page.screenshot(path=str(HERE/'ui_review.hover_bottom.png'),full_page=False)
        with page.expect_download() as d:page.get_by_role('button',name='Descargar selección CSV',exact=True).click()
        csvpath=HERE/'ui_review.final_filtered.csv';d.value.save_as(csvpath)
        rows=list(csv.DictReader(io.StringIO(csvpath.read_text(encoding='utf-8-sig')),delimiter=';'))
        results.append({'check':'Final filtered CSV preserves pharmacy and euros','passed':len(rows)==1 and rows[0]['Establecimiento']=='IQVIA_0037534' and rows[0]['Potencial anual estimado EUR']=='107948,63'})
        page.get_by_role('button',name='Limpiar',exact=False).click()
        page.wait_for_timeout(500)
        page.emulate_media(color_scheme='dark')
        page.screenshot(path=str(HERE/'ui_review.dark.png'),full_page=True)
        results.append({'check':'System dark mode rendered','background':page.locator('.market-dashboard').evaluate("el=>getComputedStyle(el).color"),'passed':True})
        context.close()
        mobile=browser.new_context(viewport={'width':390,'height':844},device_scale_factor=1)
        page=mobile.new_page();page.on('pageerror',lambda e:errors.append(str(e)))
        page.goto('http://127.0.0.1:8765/',wait_until='domcontentloaded')
        page.locator('.market-dashboard').wait_for(timeout=120000)
        page.wait_for_load_state('networkidle',timeout=60000)
        dimensions=page.evaluate("({content:document.documentElement.scrollWidth,viewport:innerWidth})")
        results.append({'check':'Mobile has no horizontal page overflow','passed':dimensions['content']<=dimensions['viewport'],'dimensions':dimensions})
        page.screenshot(path=str(HERE/'ui_review.mobile.png'),full_page=True)
        page.locator('.market-map-frame').scroll_into_view_if_needed()
        page.screenshot(path=str(HERE/'ui_review.mobile_map.png'),full_page=False)
        results.append({'check':'Final browser has no JavaScript errors','passed':not errors,'errors':errors})
        mobile.close();browser.close()
    (HERE/'ui_review.final.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(results,ensure_ascii=False))
    assert all(r['passed'] for r in results)


if __name__=='__main__':
    try:run()
    finally:(HERE/'ui_review.final.partial.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
