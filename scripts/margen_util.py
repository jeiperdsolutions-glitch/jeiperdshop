# -*- coding: utf-8 -*-
"""
Logica UNICA de margenes de JeiperdShop (la usan generar_tienda.py y
agregar_margen_excel.py). Si cambias un margen aqui, cambia en la tienda Y en el Excel.

El precio base (costo) = precio actual de portatilshop en DOP.
"""
import unicodedata

# Margen de ganancia POR TIPO de producto (no todo lleva el mismo %).
# Se elige por palabras clave en la categoria/nombre (primera que coincide gana).
# Accesorios baratos -> margen alto | equipos caros -> margen bajo (competitivo).
REGLAS_MARGEN = [
    (("cable", "cargadores y cables"),                              0.55),
    (("cover", "protector", "mica", "vidrio templado", "funda"),   0.55),
    (("adaptador",),                                               0.50),
    (("memoria", "micro sd", "pen drive", "pendrive"),             0.50),
    (("cargador",),                                               0.45),
    (("audifono", "manos libres", "auricular", "airpod"),          0.40),
    (("bocina", "parlante", "speaker", "altavoz"),                 0.40),
    (("teclado", "mouse", "raton"),                                0.40),
    (("microfono",),                                              0.38),
    (("juguete",),                                                0.40),
    (("cocina",),                                                 0.38),
    (("hogar", "domotica"),                                       0.38),
    (("deporte", "fitness", "salud", "belleza"),                  0.38),
    (("control de juego", "mando", "gamepad"),                    0.35),
    (("creador de contenido", "ring light", "tripode", "estabilizador"), 0.32),
    (("reloj", "smartwatch"),                                     0.30),
    (("gps",),                                                    0.30),
    (("repetidor", "router", "wifi"),                             0.30),
    (("convertidor", "smart tv", "tv box"),                       0.30),
    (("equipos medicos", "medico", "tensiometro", "oximetro"),    0.30),
    (("camara",),                                                0.28),
    (("disco duro", "almacenamiento", "ssd"),                     0.25),
    (("proyector",),                                             0.25),
    (("dron", "drone"),                                          0.22),
    (("tableta", "tablet", "ipad"),                               0.18),
    (("consola", "playstation", "xbox", "nintendo", "videojuego"), 0.15),
    (("laptop", "portatil", "macbook", "notebook"),               0.13),
    (("celular", "telefono", "iphone", "smartphone"),             0.12),
    (("accesorio",),                                             0.45),  # accesorio generico
]
MARGEN_DEFAULT = 0.30


def _norm(s):
    return unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()


def margen(costo, cats, nombre):
    """Devuelve el margen (ej. 0.55) para un producto segun su tipo y precio."""
    txt = _norm(" | ".join(cats) + " | " + nombre)
    base = MARGEN_DEFAULT
    for kws, m in REGLAS_MARGEN:
        if any(k in txt for k in kws):
            base = m
            break
    # Techo por precio: en lo caro hay que ser competitivo (margen % chico)
    if costo >= 40000:
        base = min(base, 0.12)
    elif costo >= 20000:
        base = min(base, 0.16)
    elif costo >= 10000:
        base = min(base, 0.20)
    elif costo >= 5000:
        base = min(base, 0.28)
    # Piso por monto bajo: en lo muy barato, ganancia minima decente
    if costo < 400:
        base = max(base, 0.60)
    elif costo < 1000:
        base = max(base, 0.45)
    return base


def precio_nice(x):
    """Redondea a un precio comercial limpio."""
    if x is None:
        return None
    x = round(x)
    if x < 10000:
        return int(round(x / 50.0) * 50)
    return int(round(x / 100.0) * 100)
