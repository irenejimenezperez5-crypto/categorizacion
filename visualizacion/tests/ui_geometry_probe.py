from pathlib import Path
import json
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b=p.chromium.launch(executable_path='C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', headless=True)
    page=b.new_page(viewport={'width':1600,'height':1040})
    page.goto('http://127.0.0.1:8765/',wait_until='networkidle',timeout=60000)
    page.locator('.leaflet-overlay-pane svg path').first.wait_for(state='attached',timeout=120000)
    result=page.locator('.leaflet-container').evaluate('''el => [...el.querySelectorAll('.leaflet-pane,svg,g,canvas,path')].filter(e=>e.tagName !== 'path').concat([...el.querySelectorAll('path')].slice(0,1)).map(e=>{let s=getComputedStyle(e),r=e.getBoundingClientRect();return {tag:e.tagName,classes:e.getAttribute('class'),rect:{x:r.x,y:r.y,w:r.width,h:r.height},css:{attrs:e.outerHTML.slice(0,700),fillOpacity:s.fillOpacity,strokeOpacity:s.strokeOpacity,opacity:s.opacity,fill:s.fill,stroke:s.stroke,strokeWidth:s.strokeWidth,overflow:s.overflow,contain:s.contain,clipPath:s.clipPath,transform:s.transform,background:s.background,backgroundColor:s.backgroundColor,zIndex:s.zIndex,display:s.display,visibility:s.visibility,width:s.width,height:s.height},d:e.getAttribute('d')?.slice(0,100)}})''')
    Path(__file__).with_suffix('.json').write_text(json.dumps(result,indent=2))
    print(json.dumps(result))
    page.locator('.leaflet-overlay-pane svg').screenshot(path=str(Path(__file__).with_name('ui_geometry_svg.png')))
    page.screenshot(path=str(Path(__file__).with_suffix('.png')),full_page=True)
    b.close()
