"""Deforestación y urbanización: cómo cambian el balance anual y los caudales pico.

1. Balance anual del Chicama (punto 1) con la curva de Fu y el modelo de cobertura del apartado 1.5 (diapositiva
   p15-refo): ω = ω₀ + Δω·(f − f₀)/100, ET = P·[1 + φ − (1 + φ^ω)^(1/ω)], Q = P − ET, sin reciclaje ni niebla.
   Parámetros leídos de index.html (objeto D.refo de la versión de Marlin Schwebel).
2. Balance anual de La Presidenta: lluvia anual del pluviómetro SIATA 41 (años con ≥ 80 % de minutos válidos) y
   caudal medio simulado por Corantioquia (2023, tabla 9: 332,9 L/s en condición normal, 13,38 km²).
3. Caudal pico de la parte alta de La Presidenta (0,55 km²) con la norma EPM GD-AS002 para la secuencia de
   coberturas de su tabla 6-3: ladera con vegetación (I = 0,30), ladera sin vegetación (0,60), residencial con
   zonas duras (0,75) y comercial o pavimentos (0,90). C = 0,15 + 0,85·I, t_c de la FAA, IDF EPM Medellín.
Salidas: productos/punto2/cobertura_chicama.csv, cobertura_presidenta_pico.csv, cobertura_resumen.json,
         productos/punto2/fig_cobertura.png
Uso: python codigo/punto2/19_cobertura_caudales.py   (después de 14_presidenta_respuesta.py)
"""
from pathlib import Path
import json
import math
import re
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

RAIZ = Path(__file__).resolve().parents[2]
OUT = RAIZ / "productos" / "punto2"
IDF_MED = {10: (62.21859, 0.98428), 25: (71.12588, 0.99176)}
COB = [("ladera con vegetación", 0.30), ("ladera deforestada", 0.60), ("residencial, zonas duras", 0.75), ("comercial / pavimentos", 0.90)]


def fu(phi, w):
    return 1 + phi - (1 + phi ** w) ** (1 / w)


def refo():
    s = (RAIZ / "index.html").read_text(encoding="utf-8")
    bloque = next(x for x in re.findall(r"<script[^>]*>(.*?)</script>", s, re.S) if x.strip().startswith("const D = "))
    return json.loads(bloque[bloque.find("{"): bloque.rfind("}") + 1])["refo"]


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    r = refo()
    filas = []
    for f in (0, 5, 10, r["f0"], 25, 50, 100):
        for tipo, nom in (("nat", "bosque nativo"), ("pl", "plantación")):
            w = r["w0"] + r["dw"][tipo] * (f - r["f0"]) / 100
            et = r["P0"] * fu(r["ETo0"] / r["P0"], w)
            filas.append({"bosque_pct": round(f, 1), "tipo": nom, "omega": round(w, 3), "ET_mm": round(et), "Q_mm": round(r["P0"] - et)})
    ch = pd.DataFrame(filas)
    ch.to_csv(OUT / "cobertura_chicama.csv", index=False)
    hoy = ch[(ch.tipo == "bosque nativo") & (ch.bosque_pct == round(r["f0"], 1))].iloc[0]
    sin = ch[(ch.tipo == "bosque nativo") & (ch.bosque_pct == 0)].iloc[0]

    mx = pd.read_csv(OUT / "presidenta_maximos_lluvia.csv")
    from importlib import util
    spec = util.spec_from_file_location("r14", RAIZ / "codigo" / "punto2" / "14_presidenta_respuesta.py")
    r14 = util.module_from_spec(spec); spec.loader.exec_module(r14)
    p, _ = r14.leer("41_lluvia", "p1")
    anual = []
    for y, g in p.groupby(p.index.year):
        n_esp = (366 if pd.Timestamp(y, 12, 31).dayofyear == 366 else 365) * 1440
        frac = g.notna().sum() / n_esp
        if frac >= 0.8:
            anual.append({"anio": y, "fraccion": round(frac, 3), "P_mm": round(float(g.sum()) / frac)})   # escalado a año completo
    an = pd.DataFrame(anual)
    P_pres = float(an["P_mm"].mean())
    Q_pres = 0.332933 * 86400 * 365.25 / (13.382196e6) * 1000

    cu = pd.read_csv(OUT / "presidenta_cuencas.csv").set_index("cierre").loc["parte_alta"]
    A, L, S = float(cu["area_km2"]), float(cu["L_cauce_m"]), float(cu["S_cauce_m_m"])
    picos = []
    for nom, I in COB:
        C = 0.15 + 0.85 * I
        tc = 0.707 * (1.1 - C) * L ** 0.5 / S ** (1 / 3)
        fila = {"cobertura": nom, "I": I, "C": round(C, 3), "tc_min": round(tc, 1)}
        for T, (k, n) in IDF_MED.items():
            i = k / (0.4 + tc / 60) ** n
            fila[f"i_T{T}_mm_h"] = round(i, 1)
            fila[f"Q_T{T}_m3_s"] = round(C * i * A / 3.6, 2)
        picos.append(fila)
    pk = pd.DataFrame(picos)
    pk.to_csv(OUT / "cobertura_presidenta_pico.csv", index=False)
    print(ch[ch.tipo == "bosque nativo"].to_string(index=False)); print(an.to_string(index=False)); print(pk.to_string(index=False))
    res = {"chicama_hoy": hoy.to_dict(), "chicama_sin_bosque": sin.to_dict(), "P_presidenta_mm": round(P_pres), "P_presidenta_anios": an["anio"].tolist(),
           "Q_presidenta_mm": round(Q_pres), "ET_presidenta_mm": round(P_pres - Q_pres), "Q_sobre_P_presidenta": round(Q_pres / P_pres, 2)}
    (OUT / "cobertura_resumen.json").write_text(json.dumps(res, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
    print(res)

    fig, (a, b) = plt.subplots(1, 2, figsize=(9.4, 3.6))
    for tipo, col in (("bosque nativo", "#1baf7a"), ("plantación", "#b07a0c")):
        d = ch[ch.tipo == tipo]
        a.plot(d.bosque_pct, d.Q_mm, "-o", color=col, ms=3, label=f"Q, {tipo}")
        a.plot(d.bosque_pct, d.ET_mm, "--", color=col, lw=1.2, label=f"ET, {tipo}")
    a.axvline(r["f0"], color="#898781", lw=0.8); a.text(r["f0"] + 1, a.get_ylim()[0] + 10 if a.get_ylim()[0] > 0 else 300, "hoy (15 %)", fontsize=7.5, color="#898781")
    a.set_xlabel("cobertura de bosque (%)"); a.set_ylabel("mm/año"); a.legend(fontsize=7.5, ncol=2)
    a.set_title("(a) Chicama: balance anual (Fu, ω según cobertura)", fontsize=9, loc="left")
    x = range(len(pk))
    b.bar([i - 0.2 for i in x], pk["Q_T10_m3_s"], width=0.4, color="#2a78d6", alpha=0.55, label="T = 10 (EPM)")
    b.bar([i + 0.2 for i in x], pk["Q_T25_m3_s"], width=0.4, color="#2a78d6", label="T = 25")
    for i, row in pk.iterrows():
        b.text(i + 0.2, row["Q_T25_m3_s"] + 0.3, f"t$_c$ {row['tc_min']:.0f} min", ha="center", fontsize=7)
    b.axhline(1.05, color="#b07a0c", ls=":", lw=1.2); b.text(2.45, 1.4, "Cottbus prados (1,05)", fontsize=7, color="#b07a0c")
    b.axhline(3.79, color="#b07a0c", ls="--", lw=1.2); b.text(2.45, 4.1, "Cottbus urbanizada (3,79)", fontsize=7, color="#b07a0c")
    corto = ["ladera\nvegetada", "ladera\ndeforestada", "residencial", "comercial,\npavimentos"]
    b.set_xticks(list(x), [f"{c}\nI = {I:.2f}".replace(".", ",") for c, (_, I) in zip(corto, COB)], fontsize=7.5)
    b.set_ylabel("caudal pico (m³/s)")
    b.set_title("(b) La Presidenta alta (0,55 km²): pico según cobertura", fontsize=9, loc="left"); b.legend(fontsize=7.5, loc="upper left", bbox_to_anchor=(0, 0.85))
    for ax in (a, b):
        ax.spines[["top", "right"]].set_visible(False); ax.grid(color="#e1e0d9", lw=0.6)
    fig.tight_layout(); fig.savefig(OUT / "fig_cobertura.png", dpi=200); plt.close(fig)


if __name__ == "__main__":
    main()
