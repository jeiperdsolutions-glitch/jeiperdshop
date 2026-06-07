# -*- coding: utf-8 -*-
"""
Pasa Cloudflare con Chrome real (headed + xvfb + stealth) y baja el catalogo
NAVEGANDO directamente a la API (page.goto), que suele pasar donde fetch() no.
Guarda:
  - catalogo_navegador.json  (lista de productos crudos de la Store API)
  - cf_cookies.json          (cookies + UA, para descargar imagenes luego)

Uso (GitHub Actions):  xvfb-run -a python solve_cloudflare.py
"""
import os
import json
import time
from playwright.sync_api import sync_playwright

CARPETA = os.path.dirname(os.path.abspath(__file__))
COOKIES_OUT = os.path.join(CARPETA, "cf_cookies.json")
CAT_OUT = os.path.join(CARPETA, "catalogo_navegador.json")
HOME = "https://portatilshoprd.com/"
API = "https://portatilshoprd.com/wp-json/wc/store/v1/products?per_page=100&page="
UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")
STEALTH = """
Object.defineProperty(navigator,'webdriver',{get:()=>undefined});
Object.defineProperty(navigator,'languages',{get:()=>['es-DO','es','en']});
Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4,5]});
window.chrome={runtime:{}};
"""


def esperar_reto(page):
    for _ in range(30):
        t = (page.title() or "").lower()
        if not ("just a moment" in t or "un momento" in t or "moment" in t or t == ""):
            return True
        time.sleep(2.5)
    return False


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--no-sandbox", "--disable-dev-shm-usage",
                  "--disable-blink-features=AutomationControlled", "--start-maximized"],
        )
        ctx = browser.new_context(locale="es-DO", user_agent=UA,
                                  viewport={"width": 1366, "height": 768})
        ctx.add_init_script(STEALTH)
        page = ctx.new_page()

        print("Abriendo la web para pasar Cloudflare...")
        page.goto(HOME, wait_until="domcontentloaded", timeout=90000)
        paso = esperar_reto(page)
        time.sleep(3)
        print(f"Reto de la pagina principal pasado: {paso} (titulo={page.title()!r})")

        # bajar catalogo navegando a cada pagina de la API
        productos = []
        pag = 1
        while pag <= 40:
            resp = page.goto(API + str(pag), wait_until="domcontentloaded", timeout=60000)
            st = resp.status if resp else None
            if st != 200:
                print(f"  pagina {pag}: HTTP {st} -> me detengo")
                if pag == 1:
                    # reintento tras pequena espera por si el reto tardo
                    time.sleep(5)
                    resp = page.goto(API + str(pag), wait_until="domcontentloaded", timeout=60000)
                    st = resp.status if resp else None
                    print(f"  pagina {pag} (reintento): HTTP {st}")
                    if st != 200:
                        break
                else:
                    break
            try:
                txt = page.evaluate("() => document.body.innerText")
                lote = json.loads(txt)
            except Exception as e:
                print(f"  pagina {pag}: no pude leer JSON ({e})")
                break
            if not lote:
                break
            productos.extend(lote)
            print(f"  pagina {pag}: {len(lote)} productos (acumulado {len(productos)})")
            if len(lote) < 100:
                break
            pag += 1
            time.sleep(0.4)

        cookies = ctx.cookies()
        ua = page.evaluate("() => navigator.userAgent")
        json.dump({"ua": ua, "cookies": cookies}, open(COOKIES_OUT, "w", encoding="utf-8"))
        if productos:
            json.dump(productos, open(CAT_OUT, "w", encoding="utf-8"))
            print(f"OK: {len(productos)} productos guardados + {len(cookies)} cookies")
        else:
            print(f"FALLO: 0 productos (Cloudflare bloqueo tambien la navegacion). cookies={len(cookies)}")
        browser.close()


if __name__ == "__main__":
    main()
