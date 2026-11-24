"""Geometría e hidráulica de las obras del punto 2 (alcantarilla del problema A y tanque del problema B),
y contexto real del cierre de la microcuenca con la ortofoto oficial y OpenStreetMap.

Entradas (rutas relativas a la raíz):
  datos/punto2/dop20_33451-5743_3840px.jpg     ortofoto DOP20 RGB, cuadro UTM33 E 451–452 km, N 5743–5744 km
                                               (GeoBasis-DE/LGB, dl-de/by-2-0; copia de Wikimedia Commons a 3 840 px)
  datos/punto2/osm_overpass_cierre_r1500m.json zanjas, caminos y alcantarillas a 1,5 km del cierre (© OpenStreetMap, ODbL)
Salidas:
  img/p2_orto_cierre.webp                      ortofoto 520 × 520 m con las capas de OSM
  productos/punto2/alcantarillas_osm.csv       longitud de cada alcantarilla real cartografiada
  productos/punto2/hidraulica_obras.csv        tirantes, velocidades, Froude y dimensiones del tanque
Valores de diseño tomados de la diapositiva del problema A y B (scripts 07_problema_A.py y 08_problema_B.py de Marlin).
Ejecutar desde la raíz:  python codigo/punto2/10_geometria_obras.py   (requiere Pillow y pyproj)
"""
from pathlib import Path
import csv
import json
import math
import statistics

from PIL import Image, ImageDraw
from pyproj import Transformer

RAIZ = Path(__file__).resolve().parents[2]
DAT, OUT, IMG = RAIZ / "datos" / "punto2", RAIZ / "productos" / "punto2", RAIZ / "img"
OUT.mkdir(parents=True, exist_ok=True)

CIERRE = (14.2950, 51.8391)            # lon, lat del cierre de la microcuenca (diapositiva p24-microcuenca)
E0, N1, LADO = 451000.0, 5744000.0, 1000.0   # cuadro de la ortofoto (esquina NO), m
G = 9.81

# Problema A (valores de diseño)
QP, D, N_MAN, S0 = 1.05, 1.00, 0.013, 0.005
# Problema B
V_TANQUE, H_MAX, CD, AO, Q_PREVIO, Q_POST = 7900.0, 1.5, 0.6, 0.332, 1.08, 3.79


def ortofoto(osm, a_utm):
    im = Image.open(DAT / "dop20_33451-5743_3840px.jpg").convert("RGB")
    res = LADO / im.size[0]
    px = lambda lon, lat: ((a_utm.transform(lon, lat)[0] - E0) / res, (N1 - a_utm.transform(lon, lat)[1]) / res)
    dr = ImageDraw.Draw(im)
    for e in osm["elements"]:
        if e["type"] != "way":
            continue
        tg, pts = e.get("tags", {}), [px(p["lon"], p["lat"]) for p in e["geometry"]]
        if tg.get("tunnel") == "culvert":
            dr.line(pts, fill=(235, 40, 40), width=16)
        elif "waterway" in tg:
            dr.line(pts, fill=(42, 140, 255), width=6)
        elif "highway" in tg:
            dr.line(pts, fill=(255, 205, 50), width=5)
    x, y = px(*CIERRE)
    dr.ellipse((x - 18, y - 18, x + 18, y + 18), outline=(255, 255, 255), width=7)
    m = 260 / res
    im.crop((int(x - m), int(y - m), int(x + m), int(y + m))).resize((1100, 1100)).save(IMG / "p2_orto_cierre.webp", quality=84)


def alcantarillas(osm, a_utm):
    e0, n0 = a_utm.transform(*CIERRE)
    filas = []
    for e in osm["elements"]:
        if e.get("tags", {}).get("tunnel") != "culvert":
            continue
        xy = [a_utm.transform(p["lon"], p["lat"]) for p in e["geometry"]]
        L = sum(math.dist(a, b) for a, b in zip(xy, xy[1:]))
        cx, cy = sum(a for a, _ in xy) / len(xy), sum(b for _, b in xy) / len(xy)
        filas.append((e["id"], round(L, 1), round(math.hypot(cx - e0, cy - n0))))
    with open(OUT / "alcantarillas_osm.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["osm_way_id", "longitud_m", "distancia_al_cierre_m"]); w.writerows(sorted(filas, key=lambda r: r[2]))
    Ls = [r[1] for r in filas]
    q = statistics.quantiles(Ls, n=4)
    print(f"Alcantarillas OSM a ≤ 1,5 km: {len(Ls)}; longitud mediana {statistics.median(Ls):.1f} m (Q1 {q[0]:.1f}, Q3 {q[2]:.1f}); "
          f"la más cercana a {min(r[2] for r in filas)} m")


def seccion(y):
    th = 2 * math.acos(1 - 2 * y / D)
    return D * D / 8 * (th - math.sin(th)), th * D / 2, D * math.sin(th / 2)   # A, P, T


def biseccion(f, lo, hi):
    for _ in range(100):
        m = (lo + hi) / 2
        lo, hi = (m, hi) if f(m) < 0 else (lo, m)
    return (lo + hi) / 2


def hidraulica():
    filas = []
    a_ll = math.pi * D * D / 4
    q_ll = a_ll * (D / 4) ** (2 / 3) * S0 ** 0.5 / N_MAN
    filas.append(("A", "capacidad a tubo lleno Q_ll", round(q_ll, 3), "m3/s"))
    filas.append(("A", "velocidad a tubo lleno", round(q_ll / a_ll, 2), "m/s"))
    for q, nom in ((QP, "caudal de diseño"), (q_ll * 0.999, "cerca de la capacidad")):
        yn = biseccion(lambda y: seccion(y)[0] * (seccion(y)[0] / seccion(y)[1]) ** (2 / 3) * S0 ** 0.5 / N_MAN - q, 1e-4, 0.938 * D)
        yc = biseccion(lambda y: 1 - q * q * seccion(y)[2] / (G * seccion(y)[0] ** 3), 1e-4, 0.99 * D)
        A, P, T = seccion(yn)
        V = q / A
        fr = V / math.sqrt(G * A / T)
        for k, v, u in (("tirante normal y_n", yn, "m"), ("y_n/D", yn / D, "-"), ("tirante crítico y_c", yc, "m"),
                        ("velocidad", V, "m/s"), ("número de Froude", fr, "-"), ("energía específica", yn + V * V / (2 * G), "m")):
            filas.append(("A", f"{k} ({nom}, Q = {q:.2f})", round(v, 3), u))
    # pendiente crítica para el caudal de diseño: la que hace y_n = y_c
    yc = biseccion(lambda y: 1 - QP * QP * seccion(y)[2] / (G * seccion(y)[0] ** 3), 1e-4, 0.99 * D)
    A, P, _ = seccion(yc)
    sc = (QP * N_MAN / (A * (A / P) ** (2 / 3))) ** 2
    filas.append(("A", "pendiente crítica S_c (Q de diseño)", round(100 * sc, 3), "%"))
    # tanque
    a_s = V_TANQUE / H_MAX
    filas += [("B", "área en planta A_s = V/h_máx", round(a_s), "m2"), ("B", "lado de un tanque cuadrado", round(math.sqrt(a_s), 1), "m"),
              ("B", "diámetro del orificio", round(math.sqrt(4 * AO / math.pi), 3), "m"),
              ("B", "salida máxima Cd·Ao·√(2 g h_máx)", round(CD * AO * math.sqrt(2 * G * H_MAX), 3), "m3/s"),
              ("B", "tiempo de vaciado desde h_máx sin entrada", round(2 * a_s * math.sqrt(H_MAX) / (CD * AO * math.sqrt(2 * G)) / 3600, 1), "h")]
    for h in (0.3, 0.4):   # vertedero de excedencias (propuesta, no calculado por Marlin): Q = Cw L h^1.5, Cw ≈ 1.7 (cresta ancha, SI)
        filas.append(("B", f"vertedero de excedencias para Q_post con h = {h} m", round(Q_POST / (1.7 * h ** 1.5), 1), "m de cresta"))
    with open(OUT / "hidraulica_obras.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["problema", "magnitud", "valor", "unidad"]); w.writerows(filas)
    for r in filas:
        print(*r)


def main():
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    osm = json.load(open(DAT / "osm_overpass_cierre_r1500m.json", encoding="utf-8"))
    a_utm = Transformer.from_crs(4326, 25833, always_xy=True)
    ortofoto(osm, a_utm)
    alcantarillas(osm, a_utm)
    hidraulica()


if __name__ == "__main__":
    main()
