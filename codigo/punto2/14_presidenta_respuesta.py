"""La Presidenta (Medellín) frente a Cottbus: lluvia extrema e intervalo lluvia → crecida medidos con SIATA.

1. Lluvia (pluviómetro 41, Comisaría El Poblado, 1 min): máximos anuales de 10–120 min con ventanas móviles;
   año válido si ≥ 80 % de los minutos tienen dato con calidad 1. Se comparan con la IDF oficial de EPM
   (estación Medellín, anexo 6.1 de GD-AS002, 2025) y con la IDF de Cottbus de 2.3.
2. Respuesta (nivel 101, Parque Lineal, 1 min): crecidas independientes (≥ 6 h entre picos) con subida ≥ 40 cm.
   Para cada una: inicio de la subida, pico, y la lluvia del pluviómetro 41 en las 4 h previas. Se miden
     · tiempo de subida = pico − inicio de la subida,
     · tiempo de retardo = pico − centroide de la lluvia,
   y se comparan con los t_c de 13_cuencas_presidenta.py (retardo ≈ 0,6 t_c en la convención del SCS).
Se usan p1 (balancín 1) y la calidad 1; p2 se usa solo para medir la diferencia entre sensores.
Entradas: datos/punto2/siata_crudo/ (12_descarga_siata.py), productos/punto2/presidenta_cuencas.csv
Salidas: productos/punto2/presidenta_maximos_lluvia.csv, presidenta_crecidas.csv, presidenta_resumen.json,
         datos/punto2/siata/presidenta_eventos_1min.csv (ventanas de las crecidas, para la gráfica y el ZIP)
Uso: python codigo/punto2/14_presidenta_respuesta.py
"""
from pathlib import Path
import json
import math
import sys

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[2]
CRUDO = RAIZ / "datos" / "punto2" / "siata_crudo"
OUT = RAIZ / "productos" / "punto2"
DOUT = RAIZ / "datos" / "punto2" / "siata"
DOUT.mkdir(parents=True, exist_ok=True)
DUR = [10, 20, 30, 60, 120]
EPM_MED = {2.33: (46.48585, 0.96443), 5: (55.16226, 0.97671), 10: (62.21859, 0.98428),
           25: (71.12588, 0.99176), 50: (77.72966, 0.99624), 100: (84.282, 1.00001)}


def leer(carpeta, col):
    partes = []
    for f in sorted((CRUDO / carpeta).glob("*.csv")):
        try:
            d = pd.read_csv(f, usecols=["fecha_hora", col, "calidad"])
        except Exception as e:  # noqa: BLE001
            print("No se pudo leer", f.name, e)
            continue
        partes.append(d)
    d = pd.concat(partes, ignore_index=True)
    d["fecha_hora"] = pd.to_datetime(d["fecha_hora"], errors="coerce").dt.floor("1min")   # algunos archivos traen segundos
    d = d.dropna(subset=["fecha_hora"]).drop_duplicates("fecha_hora").set_index("fecha_hora").sort_index()
    ok = (d["calidad"] == 1) & (d[col] >= 0)
    s = d[col].where(ok)
    return s.asfreq("1min"), d


def maximos_lluvia(p):
    filas = []
    for y, g in p.groupby(p.index.year):
        n_esp = (366 if pd.Timestamp(y, 12, 31).dayofyear == 366 else 365) * 1440
        frac = g.notna().sum() / n_esp
        fila = {"anio": y, "fraccion_minutos_validos": round(frac, 3), "valido": frac >= 0.80}
        for dmin in DUR:
            r = g.rolling(dmin, min_periods=dmin).sum()
            fila[f"i_{dmin}min_mm_h"] = round(float(r.max()) * 60 / dmin, 1) if r.notna().any() else None
        filas.append(fila)
    return pd.DataFrame(filas)


def gringorten_vs_idf(mx):
    v = mx[mx["valido"]]
    out = {}
    for dmin in DUR:
        x = np.sort(v[f"i_{dmin}min_mm_h"].dropna().values)[::-1]
        n = len(x)
        T = [(n + 0.12) / (m + 1 - 0.44) for m in range(n)]
        out[dmin] = {"i": x.round(1).tolist(), "T": [round(t, 2) for t in T]}
    return out


def crecidas(h, p):
    h10 = h.rolling(10, min_periods=6, center=True).mean()
    base = h10.rolling(6 * 60, min_periods=60).min()          # mínimo de las 6 h previas
    sube = h10 - base
    cand = sube[sube >= 40].dropna()
    eventos, ultimo = [], None
    for t in cand.index:
        if ultimo is not None and t - ultimo < pd.Timedelta(hours=6):
            continue
        ventana = h10[t: t + pd.Timedelta(hours=3)]
        if ventana.dropna().empty:
            continue
        tp = ventana.idxmax()
        pre = h10[tp - pd.Timedelta(hours=4): tp]
        b = float(base.get(tp, np.nan))
        if np.isnan(b):
            continue
        umbral = b + 0.1 * (float(h10[tp]) - b)               # inicio: cuando supera el 10 % de la subida
        bajo = pre[pre <= umbral]
        if bajo.empty:
            continue
        t0 = bajo.index[-1]
        lluv = p[tp - pd.Timedelta(hours=4): tp]
        tot = float(lluv.sum())
        if lluv.notna().mean() < 0.9 or tot < 5:
            ultimo = tp
            continue
        cum = lluv.fillna(0).cumsum()
        tc_ = cum.index[np.searchsorted(cum.values, tot / 2)]   # mediana temporal de la lluvia
        cent = lluv.fillna(0)
        t_cent = cent.index[0] + pd.Timedelta(minutes=float((np.arange(len(cent)) * cent.values).sum() / cent.values.sum()))
        i10 = lluv.rolling(10).sum()
        eventos.append({"pico": tp, "nivel_pico_cm": round(float(h10[tp]), 1), "subida_cm": round(float(h10[tp]) - b, 1),
                        "inicio_subida": t0, "t_subida_min": (tp - t0).total_seconds() / 60,
                        "lluvia_4h_mm": round(tot, 1), "i10_max_mm_h": round(float(i10.max()) * 6, 1),
                        "retardo_centroide_min": round((tp - t_cent).total_seconds() / 60, 1),
                        "retardo_mitad_lluvia_min": round((tp - tc_).total_seconds() / 60, 1),
                        "retardo_i10max_min": round((tp - i10.idxmax()).total_seconds() / 60, 1)})
        ultimo = tp
    return pd.DataFrame(eventos)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    p, praw = leer("41_lluvia", "p1")
    p2 = praw["p2"] if "p2" in praw else None
    h, _ = leer("101_nivel", "nivel")
    print(f"Lluvia 41: {p.index[0]} a {p.index[-1]}, minutos válidos {p.notna().mean():.1%}")
    print(f"Nivel 101: {h.index[0]} a {h.index[-1]}, minutos válidos {h.notna().mean():.1%}")
    mx = maximos_lluvia(p)
    mx.to_csv(OUT / "presidenta_maximos_lluvia.csv", index=False)
    print(mx.to_string(index=False))
    ev = crecidas(h, p)
    ev.to_csv(OUT / "presidenta_crecidas.csv", index=False)
    print(f"Crecidas con subida ≥ 40 cm y lluvia registrada: {len(ev)}")
    if ev.empty:
        raise SystemExit("No se encontraron crecidas con lluvia registrada: revisar los criterios o los datos")
    tcs = pd.read_csv(OUT / "presidenta_cuencas.csv").set_index("cierre").loc["siata_101"]
    res = {"periodo_lluvia": [str(p.index[0].date()), str(p.index[-1].date())],
           "periodo_nivel": [str(h.index[0].date()), str(h.index[-1].date())],
           "anios_validos_lluvia": int(mx["valido"].sum()), "n_crecidas": int(len(ev)),
           "tc_101_min": {k.replace("tc_", "").replace("_min", ""): float(v) for k, v in tcs.items() if str(k).startswith("tc_")},
           "area_101_km2": float(tcs["area_km2"])}
    for c in ("t_subida_min", "retardo_centroide_min", "retardo_i10max_min"):
        q = ev[c].quantile([0.25, 0.5, 0.75])
        res[c] = [round(float(x), 1) for x in q]
        print(c, "Q1/mediana/Q3:", res[c])
    if p2 is not None:
        a = praw.loc[praw["calidad"] == 1, ["p1", "p2"]].clip(lower=0)
        res["p2_sobre_p1"] = round(float(a["p2"].sum() / a["p1"].sum()), 3)
    res["gringorten"] = gringorten_vs_idf(mx)
    res["epm_medellin"] = {str(k): v for k, v in EPM_MED.items()}
    # crecidas mayores para la gráfica: ventana de −3 h a +3 h, 1 min
    top = ev.nlargest(6, "subida_cm")
    filas = []
    for k, e in enumerate(top.itertuples()):
        w0, w1 = e.pico - pd.Timedelta(hours=3), e.pico + pd.Timedelta(hours=3)
        for t in pd.date_range(w0, w1, freq="1min"):
            filas.append({"evento": k, "pico": e.pico, "min_desde_pico": int((t - e.pico).total_seconds() / 60),
                          "nivel_cm": None if pd.isna(h.get(t)) else round(float(h.get(t)), 1),
                          "lluvia_mm": None if pd.isna(p.get(t)) else float(p.get(t))})
    pd.DataFrame(filas).to_csv(DOUT / "presidenta_eventos_1min.csv", index=False)
    res["eventos_grafica"] = [{"pico": str(e.pico), "subida_cm": e.subida_cm, "lluvia_4h_mm": e.lluvia_4h_mm,
                               "retardo_centroide_min": e.retardo_centroide_min} for e in top.itertuples()]
    (OUT / "presidenta_resumen.json").write_text(json.dumps(res, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
