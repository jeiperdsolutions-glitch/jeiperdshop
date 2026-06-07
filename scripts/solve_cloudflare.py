# -*- coding: utf-8 -*-
"""
Resuelve el reto 'Just a moment' de Cloudflare con un Chrome REAL (headed) bajo
pantalla virtual (xvfb) + tecnicas anti-deteccion, y guarda cookies + UA en
cf_cookies.json para que extraer_portatilshop.py descargue desde la nube.

Uso (en GitHub Actions):  xvfb-run -a python solve_cloudflare.py
"""
import os
import json
import time
from playwright.sync_api import sync_playwright

CARPETA = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(CARPETA, "cf_cookies.json")
URL = "https://portatilshoprd.com/"
API_TEST = "https://portatilshoprd.com/wp-json/wc/store/v1/products?per_page=1"
UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")

STEALTH = """
Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
Object.defineProperty(navigator, 'languages', {get: () => ['es-DO','es','en']});
Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
window.chrome = { runtime: {} };
const _q = window.navigator.permissions && window.navigator.permissions.query;
if (_q) { window.navigator.permissions.query = (p) =>
    p && p.name === 'notifications'
      ? Promise.resolve({state: Notification.permission})
      : _q(p); }
"""


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--start-maximized",
            ],
        )
        ctx = browser.new_context(
            locale="es-DO",
            user_agent=UA,
            viewport={"width": 1366, "height": 768},
        )
        ctx.add_init_script(STEALTH)
        page = ctx.new_page()
        print("Abriendo la web (headed + stealth) para resolver Cloudflare...")
        page.goto(URL, wait_until="domcontentloaded", timeout=90000)

        ok = False
        for intento in range(30):  # ~ hasta 75s
            time.sleep(2.5)
            titulo = (page.title() or "").lower()
            reto = ("just a moment" in titulo or "un momento" in titulo
                    or "moment" in titulo or "attention" in titulo or titulo == "")
            if not reto:
                # confirmar que la API ya responde 200 dentro del navegador
                try:
                    st = page.evaluate("async (u) => (await fetch(u)).status", API_TEST)
                except Exception:
                    st = None
                print(f"  intento {intento}: titulo={page.title()!r} api={st}")
                if st == 200:
                    ok = True
                    break
            else:
                print(f"  intento {intento}: aun en reto ({page.title()!r})")
                if intento == 12:
                    try:
                        page.reload(wait_until="domcontentloaded", timeout=60000)
                    except Exception:
                        pass

        time.sleep(3)
        try:
            status = page.evaluate("async (u) => (await fetch(u)).status", API_TEST)
        except Exception:
            status = None
        print(f"Resultado final: reto_resuelto={ok} api_status={status} titulo={page.title()!r}")

        cookies = ctx.cookies()
        ua = page.evaluate("() => navigator.userAgent")
        json.dump({"ua": ua, "cookies": cookies}, open(SALIDA, "w", encoding="utf-8"))
        print(f"Guardadas {len(cookies)} cookies en {os.path.basename(SALIDA)}")
        for c in cookies:
            if "cf" in c["name"].lower() or "clearance" in c["name"].lower():
                print(f"  cookie relevante: {c['name']}")
        browser.close()


if __name__ == "__main__":
    main()
