"""Inspect the source action through the canonical local dashboard UI."""
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

HERE=Path(__file__).resolve().parent
with sync_playwright() as p:
    browser=p.chromium.launch(executable_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',headless=True)
    page=browser.new_page(viewport={'width':1600,'height':1040})
    page.goto('http://127.0.0.1:8765/',wait_until='domcontentloaded')
    page.locator('.market-dashboard').wait_for(timeout=120000)
    page.wait_for_load_state('networkidle',timeout=60000)
    page.locator('#market-search').fill('0037534')
    page.wait_for_function("document.querySelectorAll('.market-metric-body strong')[1]?.textContent==='1'",timeout=30000)
    native=page.get_by_role('combobox',name='Potencial anual de la selección actions',exact=True)
    if native.count():
        native.select_option(label='View data source')
    else:
        control=page.locator('.market-metric').first.get_by_role('button').last
        print('Opening source action:',control.get_attribute('aria-label'),flush=True)
        control.click()
        page.get_by_role('menuitem',name='View data source',exact=True).click()
    page.get_by_text('Definitions',exact=True).wait_for(timeout=10000)
    source_visible=page.get_by_text('UNIVERSO_FARMACIAS_CATEGORIZADO_2027.xlsx',exact=True).is_visible()
    overview=page.locator('body').inner_text()[-7000:]
    assert source_visible
    assert 'potential_eur / sb_captured_eur' in overview
    page.screenshot(path=str(HERE/'ui_review.source_inspector.png'),full_page=False)
    page.get_by_text('Data preview',exact=True).click()
    page.wait_for_timeout(500)
    preview=page.locator('body').inner_text()[-7000:]
    assert 'FCIA. ROSA SANCHEZ CALUCHO' in preview
    report={'source_action_opened':True,'reviewed_excel_visible':source_visible,'denominator_definition_visible':True,'filtered_pharmacy_in_preview':True,'overview_excerpt':overview,'preview_excerpt':preview}
    (HERE/'ui_review.sources.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in report.items() if k not in ['overview_excerpt','preview_excerpt']},ensure_ascii=False))
    browser.close()
