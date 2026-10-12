"""Render the private local map in installed Edge for independent visual QA.

Run from the QA environment containing Playwright; production needs no browser
automation dependency. No external browser service or data upload is used.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import time
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="http://127.0.0.1:8765/")
    parser.add_argument("--width", type=int, default=1600)
    parser.add_argument("--height", type=int, default=1040)
    parser.add_argument("--name", default="desktop")
    args = parser.parse_args()
    if urlsplit(args.url).hostname not in {"127.0.0.1", "localhost", "::1"}:
        raise ValueError("QA supports only the local visualization")
    errors, failed_requests, external_requests = [], [], []
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=EDGE, headless=True)
        context = browser.new_context(viewport={"width": args.width, "height": args.height}, device_scale_factor=1)
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.on("requestfailed", lambda request: failed_requests.append({"url": request.url, "error": request.failure}))
        page.on("request", lambda request: external_requests.append(request.url) if urlsplit(request.url).hostname not in {"127.0.0.1", "localhost", "::1", None} else None)
        started = time.monotonic()
        page.goto(args.url, wait_until="networkidle", timeout=60000)
        page.locator(".market-dashboard").wait_for(state="visible", timeout=120000)
        page.locator(".leaflet-container canvas.market-marker-canvas").wait_for(state="attached", timeout=60000)
        page.get_by_role("combobox", name="Provincia", exact=True).wait_for(state="visible", timeout=60000)
        page.wait_for_load_state("networkidle", timeout=60000)
        page.wait_for_timeout(1000)
        screenshot = ROOT / f"ui_review.{args.name}.png"
        page.screenshot(path=str(screenshot), full_page=True)
        report = {
            "url": args.url, "browser": browser.version,
            "viewport": {"width": args.width, "height": args.height},
            "load_seconds": round(time.monotonic() - started, 3),
            "title": page.title(), "javascript_errors": errors,
            "failed_requests": failed_requests, "external_requests": external_requests,
            "body_text": page.locator(".market-dashboard").evaluate("el => { const walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT); let text='',node; while(node=walker.nextNode()){if(node.parentElement?.closest('option')) continue; if(node.textContent.trim())text+=node.textContent.trim()+'\\n';}return text.slice(0,20000); }"),
            "selects": page.locator("select").evaluate_all("nodes => nodes.map(el => ({aria:el.getAttribute('aria-label'),name:el.name,optionCount:el.options.length,firstOptions:[...el.options].slice(0,8).map(o=>({label:o.label,value:o.value})),label:el.closest('label')?.innerText}))"),
            "inputs": page.locator("input").evaluate_all("nodes => nodes.map(el => ({aria:el.getAttribute('aria-label'),type:el.type,placeholder:el.placeholder,label:el.closest('label')?.innerText}))"),
            "buttons": page.get_by_role("button").all_text_contents(),
            "document_geometry": page.evaluate("({width:document.documentElement.scrollWidth,height:document.documentElement.scrollHeight,viewportWidth:innerWidth})"),
            "screenshot": screenshot.name,
        }
        (ROOT / f"ui_review.{args.name}.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({k:v for k,v in report.items() if k not in {"body_text", "selects", "inputs", "buttons"}}, ensure_ascii=False))
        context.close()
        browser.close()


if __name__ == "__main__":
    main()
