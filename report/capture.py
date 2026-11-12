"""Capture the screenshots used in the project report (headless Chrome via Playwright).

Run from the tideline folder while the site is served on http://localhost:8777:
    python -m http.server 8777 --directory web
    python report/capture.py
"""
import os
import sys
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "screens")
SITE = "http://localhost:8777/"
NB = "http://localhost:8777/_nb/notebook.html"
os.makedirs(OUT, exist_ok=True)
ARGS = ["--use-angle=swiftshader", "--enable-unsafe-swiftshader", "--ignore-gpu-blocklist"]


def shot(page, name, el=None):
    path = os.path.join(OUT, name + ".png")
    (page.locator(el).first if el else page).screenshot(path=path)
    print("saved", name)


def show(page, sel, wait=1200):
    page.locator(sel).first.scroll_into_view_if_needed()
    page.evaluate("document.querySelectorAll('.reveal.pre').forEach(e => e.classList.remove('pre'))")
    page.wait_for_timeout(wait)


def site(b):
    page = b.new_page(viewport={"width": 1440, "height": 900}, device_scale_factor=1.25, color_scheme="light")
    page.set_default_timeout(240000)
    page.goto(SITE, wait_until="networkidle")
    page.wait_for_timeout(6000)
    vh = page.evaluate("innerHeight")
    for k, name in [(0, "web01_hero_2024"), (1, "web02_1880"), (2, "web03_2024_rate"), (3, "web04_2100_forecast"),
                    (4, "web05_2100_storm"), (5, "web06_2050_defended")]:
        page.evaluate(f"scrollTo({{top: {k * vh}, behavior: 'instant'}})")
        page.wait_for_timeout(5500)
        shot(page, name)
    # record and forecast cards
    for sel, name in [("#record .sec-head", None), ("#rec-stats", "web07_record_stats"), ("#c-gmsl", None)]:
        show(page, sel)
        if name:
            shot(page, name, sel)
    show(page, "#c-gmsl"); shot(page, "web08_gmsl_chart", "#record .card:has(#c-gmsl)")
    show(page, "#c-mum"); shot(page, "web09_mumbai_quality", "#record .grid.g2")
    show(page, "#metrics"); shot(page, "web10_law_metrics", "#forecast .grid.g2")
    show(page, "#c-pred"); shot(page, "web11_pred_vs_obs", "#forecast .card:has(#c-pred)")
    show(page, "#c-proj"); shot(page, "web12_projections", "#forecast .grid.g2:has(#c-proj)")
    # simulator: default hybrid, then raised houses and an overtopped embankment
    show(page, "#simcanvas", 6000); shot(page, "web13_sim_hybrid", ".sim")
    page.click('[data-d="Embankment"]')
    page.evaluate("const y=document.querySelector('#sl-year'); y.value=2092; y.dispatchEvent(new Event('input',{bubbles:true}))")
    page.wait_for_timeout(6000); shot(page, "web14_sim_embankment_2092", ".sim")
    page.click('[data-d="Mangrove belt"]')
    page.click('[data-s="SSP5-8.5"]')
    page.evaluate("const y=document.querySelector('#sl-year'); y.value=2100; y.dispatchEvent(new Event('input',{bubbles:true}))")
    page.wait_for_timeout(6000); shot(page, "web15_sim_mangrove_ssp585", ".sim")
    show(page, "#c-ead"); shot(page, "web16_ead_chart", "#defend .card:has(#c-ead)")
    show(page, "#dcards"); shot(page, "web17_defence_cards", "#dcards")
    show(page, "#c-cliff"); shot(page, "web18_cliff_chart", "#compare .card:has(#c-cliff)")
    show(page, "#assume"); shot(page, "web19_method", "#method .grid.g2")
    page.close()


def notebook(b):
    page = b.new_page(viewport={"width": 1280, "height": 900}, device_scale_factor=2, color_scheme="light")
    page.goto(NB, wait_until="load")
    page.wait_for_timeout(2500)
    cells = page.locator(".jp-Cell")
    n = cells.count()
    wanted = {0: "nb00_intro", 2: "nb01_setup", 4: "nb02_fetch", 6: "nb03_data_quality", 8: "nb04_splice", 10: "nb05_features",
              11: "nb06_features_plot", 13: "nb07_models", 15: "nb08_metrics", 17: "nb09_robustness", 19: "nb10_explainability",
              22: "nb11_mumbai", 24: "nb12_projections", 25: "nb13_projection_plot", 28: "nb14_simulator", 29: "nb15_simulator_plot"}
    done = set()
    for i, name in wanted.items():
        if i >= n:
            continue
        c = cells.nth(i)
        c.scroll_into_view_if_needed(); page.wait_for_timeout(200)
        box = c.bounding_box()
        if box and box["height"] > 2200:   # very long cells: keep the top part only
            page.screenshot(path=os.path.join(OUT, name + ".png"), full_page=True,
                            clip={"x": box["x"], "y": box["y"] + page.evaluate("scrollY"), "width": box["width"], "height": 2200})
        else:
            c.screenshot(path=os.path.join(OUT, name + ".png"))
        print("saved", name, "cell", i, int(box["height"]) if box else "?")
        done.add(name)
    print("missing:", sorted(set(wanted.values()) - done))
    page.close()


with sync_playwright() as p:
    b = p.chromium.launch(channel="chrome", headless=True, args=ARGS)
    which = sys.argv[1:] or ["site", "notebook"]
    if "site" in which:
        site(b)
    if "notebook" in which:
        notebook(b)
    b.close()
