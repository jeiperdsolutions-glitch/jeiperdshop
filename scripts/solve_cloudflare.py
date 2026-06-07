# -*- coding: utf-8 -*-
"""
Abre un navegador real (Playwright/Chromium) para pasar el reto 'Just a moment'
de Cloudflare de portatilshoprd.com, y guarda las cookies + User-Agent en
cf_cookies.json para que extraer_portatilshop.py pueda descargar desde la nube.

Uso (en GitHub Actions): python solve_cloudflare.py
"""
import os
import json
import time
from playwright.sync_api import sync_playwright

CARPETA = os.path.dirname(os.path.abspath(__file__))
SALIDA = os.path.join(CARPETA, "cf_cookies.json")
URL = "https://portatilshoprd.com/"
API_TEST = "https://portatilshoprd.com/wp-json/wc/store/v1/products?per_page=1"


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        ctx = browser.new_context(
            locale="es-DO",
            viewport={"width": 1366, "height": 768},
        )
        page = ctx.new_page()
        print("Abriendo la web para resolver Cloudflare...")
        page.goto(URL, wait_until="domcontentloaded", timeout=90000)

        # esperar a que se resuelva el reto (el titulo deja de ser 'Just a moment')
        ok = False
        for intento in range(25):
            titulo = (page.title() or "").lower()
            if "just a moment" not in titulo and "moment" not in titulo and "attention" not in titulo:
                ok = True
                break
            time.sleep(2)
        time.sleep(4)  # margen extra para que asiente la cookie cf_clearance
        print(f"Titulo final: {page.title()!r} (reto resuelto: {ok})")

        # comprobar que la API ya responde JSON dentro del navegador
        try:
            status = page.evaluate(
                "async (u) => { const r = await fetch(u); return r.status; }", API_TEST
            )
            print(f"API status dentro del navegador: {status}")
        except Exception as e:
            print(f"No pude probar la API en el navegador: {e}")

        cookies = ctx.cookies()
        ua = page.evaluate("() => navigator.userAgent")
        json.dump({"ua": ua, "cookies": cookies}, open(SALIDA, "w", encoding="utf-8"))
        print(f"Guardadas {len(cookies)} cookies en {os.path.basename(SALIDA)}")
        browser.close()


if __name__ == "__main__":
    main()
