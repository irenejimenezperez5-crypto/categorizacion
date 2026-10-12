"""Final targeted dark-category contrast and overlapping-tooltip check."""
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

HERE=Path(__file__).resolve().parent


def luminance(rgb):
    channels=[v/255 for v in rgb]
    linear=[v/12.92 if v<=0.04045 else ((v+.055)/1.055)**2.4 for v in channels]
    return sum(a*b for a,b in zip(linear,[.2126,.7152,.0722]))


def run():
    with sync_playwright() as p:
        browser=p.chromium.launch(executable_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',headless=True)
        page=browser.new_page(viewport={'width':1600,'height':1040},color_scheme='dark')
        page.goto('http://127.0.0.1:8765/',wait_until='domcontentloaded')
        page.locator('.market-dashboard').wait_for(timeout=120000)
        page.wait_for_load_state('networkidle',timeout=60000)
        badges=page.locator('.market-category-summary .market-category-badge').evaluate_all("els=>els.map(el=>({category:el.textContent,color:getComputedStyle(el).color,background:getComputedStyle(el).backgroundColor,parent:getComputedStyle(el.closest('.market-summary-content')).backgroundColor}))")
        contrast=[]
        import re
        for badge in badges:
            nums=lambda s:[float(x) for x in re.findall(r'[\d.]+',s)]
            fg,bg,parent=nums(badge['color']),nums(badge['background']),nums(badge['parent'])
            alpha=bg[3] if len(bg)==4 else 1
            actual=[bg[i]*alpha+parent[i]*(1-alpha) for i in range(3)]
            l1,l2=sorted([luminance(fg[:3]),luminance(actual)])
            ratio=(l2+.05)/(l1+.05)
            contrast.append({'category':badge['category'],'ratio':round(ratio,3),'passed':ratio>=4.5,'colors':badge})
        page.screenshot(path=str(HERE/'ui_review.dark.png'),full_page=True)
        page.locator('#market-search').fill('0037534')
        page.wait_for_function("document.querySelectorAll('.market-metric-body strong')[1]?.textContent==='1'",timeout=30000)
        page.locator('.market-row-name').first.click()
        page.locator('.market-store-detail').wait_for()
        page.locator('.market-leaflet-host').scroll_into_view_if_needed()
        box=page.locator('.market-leaflet-host').bounding_box()
        page.mouse.move(box['x']+box['width']/2,box['y']+box['height']/2)
        page.locator('.market-map-hover').wait_for(timeout=10000)
        territory=page.locator('.market-territory-tooltip').evaluate_all("els=>els.map(el=>getComputedStyle(el).visibility)")
        hover=page.locator('.market-map-hover').inner_text()
        results={'dark_badges':contrast,'territory_tooltip_hidden_during_pharmacy_hover':all(v=='hidden' for v in territory),'hover':hover,'loading_bar_cleared':not page.get_by_text('Creating dashboard...',exact=True).count()}
        page.screenshot(path=str(HERE/'ui_review.dark_hover.png'),full_page=False)
        (HERE/'ui_review.contrast.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
        print(json.dumps(results,ensure_ascii=False),flush=True)
        browser.close()
        assert all(x['passed'] for x in contrast)
        assert results['territory_tooltip_hidden_during_pharmacy_hover']


if __name__=='__main__':run()
