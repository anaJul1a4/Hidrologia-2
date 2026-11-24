"""Diseño de la «misma obra» en la parte alta de La Presidenta con las reglas de EPM, y datos de la diapositiva
«p2-presidenta» (productos/punto2/presidenta_datos.js).

Diseño (método racional, como el problema A de Cottbus):
  · C = 0,15 + 0,85·I (EPM GD-AS002, ec. 6-6) con I de la tabla 6-3: ladera con vegetación 0,30; residencial con
    zonas duras 0,75.
  · t_c con la ecuación de la FAA que exige EPM (ec. 6-8): T_e = 0,707 (1,1 − C) L^0,5 / S^(1/3), L y S del cauce
    principal de 13_cuencas_presidenta.py.
  · i de la IDF oficial de EPM (anexo 6.1, 2025): estaciones Medellín (1 491 m) y Palmas (2 550 m).
  · T = 10 años (tabla 6-1 de EPM, áreas de 10–1 000 ha) y T = 25 (el mismo de Cottbus).
  · Manning a tubo lleno, concreto n = 0,013, S₀ = 2 %, y velocidad máxima de EPM de 5 m/s.
Entradas: productos/punto2/presidenta_cuencas.{csv,geojson}, presidenta_crecidas.csv, presidenta_maximos_lluvia.csv,
          presidenta_resumen.json, datos/punto2/siata/presidenta_eventos_1min.csv
Salidas:  productos/punto2/presidenta_diseno.csv, productos/punto2/presidenta_datos.js
Uso: python codigo/punto2/15_presidenta_web.py
"""
from pathlib import Path
import csv
import json
import math
import sys

import pandas as pd

RAIZ = Path(__file__).resolve().parents[2]
OUT = RAIZ / "productos" / "punto2"
IDF = {"Medellín": {2.33: (46.48585, 0.96443), 5: (55.16226, 0.97671), 10: (62.21859, 0.98428), 25: (71.12588, 0.99176), 50: (77.72966, 0.99624), 100: (84.282, 1.00001)},
       "Palmas": {2.33: (54.61676, 0.94512), 5: (67.70875, 0.96868), 10: (78.3823, 0.98279), 25: (91.87895, 0.99648), 50: (101.89675, 1.00455), 100: (111.8448, 1.01128)}}
ESCENARIOS = {"ladera con vegetación": 0.30, "urbanizada (residencial, zonas duras)": 0.75}
N_MAN, S0, VMAX = 0.013, 0.02, 5.0
DN = [0.6, 0.8, 0.9, 1.0, 1.2, 1.4, 1.5, 1.6, 1.8, 2.0, 2.2, 2.4]


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    cu = pd.read_csv(OUT / "presidenta_cuencas.csv").set_index("cierre")
    pa = cu.loc["parte_alta"]
    A, L, S = float(pa["area_km2"]), float(pa["L_cauce_m"]), float(pa["S_cauce_m_m"])
    filas = []
    for esc, I in ESCENARIOS.items():
        C = 0.15 + 0.85 * I
        tc = 0.707 * (1.1 - C) * L ** 0.5 / S ** (1 / 3)
        for est, tab in IDF.items():
            for T in (10, 25):
                k, n = tab[T]
                i = k / (0.4 + tc / 60) ** n
                Q = C * i * A / 3.6
                D = (N_MAN * Q / (0.312 * math.sqrt(S0))) ** 0.375
                dn = next(v for v in DN if v >= D - 1e-9)
                V = Q / (math.pi * dn * dn / 4)
                # pendiente que deja la velocidad a tubo lleno en 5 m/s con ese DN (Manning)
                s5 = (VMAX * N_MAN / (dn / 4) ** (2 / 3)) ** 2
                filas.append({"escenario": esc, "I": I, "C": round(C, 3), "tc_FAA_min": round(tc, 1), "estacion_IDF": est, "T": T,
                              "i_mm_h": round(i, 1), "Q_m3_s": round(Q, 2), "q_m3_s_km2": round(Q / A, 1), "D_calc_m": round(D, 2),
                              "DN_m": dn, "V_tubo_lleno_m_s": round(V, 2), "S0_para_V5_pct": round(100 * s5, 2)})
    pd.DataFrame(filas).to_csv(OUT / "presidenta_diseno.csv", index=False)
    print(pd.DataFrame(filas).to_string(index=False))

    res = json.loads((OUT / "presidenta_resumen.json").read_text(encoding="utf-8"))
    ev = pd.read_csv(OUT / "presidenta_crecidas.csv")
    mx = pd.read_csv(OUT / "presidenta_maximos_lluvia.csv")
    ve = pd.read_csv(RAIZ / "datos" / "punto2" / "siata" / "presidenta_eventos_1min.csv")
    geo = json.loads((OUT / "presidenta_cuencas.geojson").read_text(encoding="utf-8"))
    datos = {
        "cuencas": cu.reset_index().to_dict(orient="records"), "geo": geo,
        "estaciones": [{"n": "SIATA 101 · nivel", "lon": -75.5694, "lat": 6.2079}, {"n": "SIATA 41 · lluvia", "lon": -75.5662, "lat": 6.2067},
                       {"n": "EPM Medellín (IDF)", "lon": -75.568633, "lat": 6.25296809}, {"n": "EPM Palmas (IDF)", "lon": -75.5355004, "lat": 6.15225492}],
        "crecidas": {"subida_cm": ev["subida_cm"].tolist(), "t_subida": ev["t_subida_min"].tolist(), "retardo": ev["retardo_centroide_min"].tolist()},
        "maximos": mx[mx["valido"]].to_dict(orient="records"),
        "eventos": [{"pico": p, "min": g["min_desde_pico"].tolist(), "nivel": g["nivel_cm"].where(g["nivel_cm"].notna(), None).tolist(),
                     "lluvia": g["lluvia_mm"].where(g["lluvia_mm"].notna(), None).tolist()} for p, g in ve.groupby("pico", sort=False)],
        "resumen": {k: v for k, v in res.items() if k != "gringorten"}, "idf": {e: {str(T): v for T, v in t.items()} for e, t in IDF.items()},
        "diseno": filas,
    }
    txt = json.dumps(datos, ensure_ascii=False, separators=(",", ":"), default=lambda o: None).replace("NaN", "null")
    (OUT / "presidenta_datos.js").write_text("/* generado por codigo/punto2/15_presidenta_web.py */\nwindow.PRES = " + txt + ";\n", encoding="utf-8")
    print("presidenta_datos.js:", round(len(txt) / 1e3), "kB")


if __name__ == "__main__":
    main()
