"""Figuras estáticas (PNG) del informe para los análisis que en la página son interactivos.
Lee solo productos ya generados por los scripts 00, 11, 13, 14 y 15; no recalcula nada nuevo, salvo las curvas
de los escenarios de cambio climático, con las mismas ecuaciones de 11_escenarios_clima.py.
Salidas: productos/punto1/fig_balance_cubeta.png, productos/punto2/fig_presidenta_respuesta.png,
         productos/punto2/fig_idf_comparacion.png, productos/punto2/fig_clima_escenarios.png
Uso: python codigo/informe/16_figuras_informe.py
"""
from pathlib import Path
import json
import math

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[2]
P1, P2 = RAIZ / "productos" / "punto1", RAIZ / "productos" / "punto2"
# paleta de la Tarea 1 (estilo_figuras.py): P_L azul, P_I naranja, Q verde
C_PL, C_PI, C_Q, C_E, C_ET, GR, TINTA = "#2a78d6", "#eb6834", "#1baf7a", "#b07a0c", "#0f8a8a", "#898781", "#0b0b0b"
plt.rcParams.update({"font.size": 9, "axes.spines.top": False, "axes.spines.right": False, "axes.grid": True,
                     "grid.color": "#e1e0d9", "grid.linewidth": 0.6, "axes.edgecolor": "#c3c2b7", "savefig.dpi": 200,
                     "legend.frameon": False})
MES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def balance():
    m = pd.read_csv(P1 / "balance_cubeta_mensual.csv")
    m["mo"] = m["mes"].str[5:7].astype(int)
    c = m.groupby("mo").mean(numeric_only=True)
    # ΔS de la climatología = P − ET − Q (suelo + depósito)
    ds = c["P_L_mm"] - c["ET_sim_mm"] - c["Q_sim_mm"]
    fig, (a, b) = plt.subplots(2, 1, figsize=(7.2, 5.2), sharex=True, gridspec_kw={"height_ratios": [2.3, 1]})
    x = np.arange(12)
    a.bar(x, c["P_L_mm"], color=C_PL, alpha=0.45, label="P (P$_L$, PISCOp)")
    a.fill_between(x, c["ET_sim_mm"], c["ETp_PISCOeo_mm"], color=C_E, alpha=0.15, label="demanda no satisfecha")
    a.plot(x, c["ETp_PISCOeo_mm"], color=C_E, ls="--", lw=2, label="ETp (PISCOeo PM)")
    a.plot(x, c["ET_sim_mm"], color=C_ET, lw=2.5, marker="o", ms=4, label="ET real (simulada)")
    a.plot(x, c["Q_sim_mm"], color=C_Q, lw=2, label="Q simulado")
    a.plot(x, c["R_obs_mm"], ls="none", marker="o", mfc="none", mec=TINTA, ms=6, label="Q observado (R)")
    a.set_ylabel("mm/mes"); a.legend(ncol=3, fontsize=8, loc="upper right")
    b.bar(x, ds, color=[C_PL if v >= 0 else "#e34948" for v in ds], alpha=0.75)
    b.axhline(0, color=GR, lw=0.8); b.set_ylabel("ΔS (mm/mes)")
    b.set_xticks(x, MES)
    fig.tight_layout(); fig.savefig(P1 / "fig_balance_cubeta.png"); plt.close(fig)


def presidenta():
    ev = pd.read_csv(P2 / "presidenta_crecidas.csv")
    w = pd.read_csv(RAIZ / "datos" / "punto2" / "siata" / "presidenta_eventos_1min.csv")
    res = json.loads((P2 / "presidenta_resumen.json").read_text(encoding="utf-8"))
    e0 = w[w["pico"] == w["pico"].iloc[0]]
    fig, (a, b) = plt.subplots(1, 2, figsize=(9.2, 3.6))
    h = e0["min_desde_pico"] / 60
    a.plot(h, e0["nivel_cm"], color=C_Q, lw=2, label="nivel, estación 101")
    a.set_xlim(-2.5, 2.5); a.set_xlabel("horas respecto al pico"); a.set_ylabel("nivel (cm)")
    a2 = a.twinx(); a2.bar(h, e0["lluvia_mm"] * 60, width=1 / 60, color=C_PL, alpha=0.6, label="lluvia, pluviómetro 41")
    a2.invert_yaxis(); a2.set_ylabel("lluvia (mm/h)"); a2.grid(False); a2.spines["right"].set_visible(True)
    a.set_title(f"(a) crecida del {str(e0['pico'].iloc[0])[:10]}", fontsize=9, loc="left")
    a.legend(loc="upper left", fontsize=8); a2.legend(loc="center left", fontsize=8)
    bins = np.arange(0, 135, 5)
    b.hist(ev["t_subida_min"].clip(upper=130), bins=bins, color=C_Q, alpha=0.45, density=True, label=f"todas ({len(ev)})")
    b.hist(ev.loc[ev["subida_cm"] >= 100, "t_subida_min"].clip(upper=130), bins=bins, color=C_PL, alpha=0.6, density=True,
           label=f"≥ 1 m ({int((ev['subida_cm'] >= 100).sum())})")
    nom = {"Kirpich": "Kirpich", "Temez": "Témez", "Bransby_Williams": "Bransby-W.", "Giandotti": "Giandotti"}
    ymax = b.get_ylim()[1]
    for j, (k, v) in enumerate(res["tc_101_min"].items()):
        b.axvline(v, color="#e34948" if k == "Kirpich" else GR, ls="--", lw=1.8 if k == "Kirpich" else 1.2)
        b.text(v + 1, ymax * (0.95 - 0.08 * j), f"{nom.get(k, k)} {v:.0f}", fontsize=7.5, color="#e34948" if k == "Kirpich" else GR)
    b.set_xlabel("tiempo de subida (min)"); b.set_ylabel("densidad")
    b.set_title("(b) tiempo de subida medido frente a t$_c$", fontsize=9, loc="left"); b.legend(fontsize=8, loc="center right")
    fig.tight_layout(); fig.savefig(P2 / "fig_presidenta_respuesta.png"); plt.close(fig)


def idf():
    res = json.loads((P2 / "presidenta_resumen.json").read_text(encoding="utf-8"))
    mx = pd.read_csv(P2 / "presidenta_maximos_lluvia.csv"); mx = mx[mx["valido"]]
    idf_epm = res["epm_medellin"]
    pal = {2.33: (54.61676, 0.94512), 5: (67.70875, 0.96868), 10: (78.3823, 0.98279), 25: (91.87895, 0.99648), 50: (101.89675, 1.00455), 100: (111.8448, 1.01128)}
    MU, SG, XI = 19.42, 7.52, 0.044
    Ts = np.logspace(np.log10(2), 2, 60)
    fig, ax = plt.subplots(1, 3, figsize=(9.6, 3.3), sharey=False)
    for a, dmin in zip(ax, (10, 30, 60)):
        d = dmin / 60
        cott = [(MU + SG * ((-math.log(1 - 1 / T)) ** (-XI) - 1) / XI) / (d + 0.137) ** 0.872 for T in Ts]
        a.plot(Ts, cott, color=C_E, lw=2, label="Cottbus (IDF 2.3)")
        TT = sorted(float(k) for k in idf_epm)
        a.plot(TT, [idf_epm[str(k) if str(k) in idf_epm else str(int(k))][0] / (0.4 + d) ** idf_epm[str(k) if str(k) in idf_epm else str(int(k))][1] for k in TT],
               color=C_PL, lw=2, marker="o", ms=3, label="EPM Medellín")
        a.plot(list(pal), [k / (0.4 + d) ** n for k, n in pal.values()], color=C_PL, lw=1.2, ls="--", label="EPM Palmas")
        x = np.sort(mx[f"i_{dmin}min_mm_h"].values)[::-1]; n = len(x)
        a.plot([(n + 0.12) / (m + 1 - 0.44) for m in range(n)], x, "o", mfc="none", mec=TINTA, ms=5, label=f"SIATA 41 ({n} años)")
        a.set_xscale("log"); a.set_xticks([2, 5, 10, 25, 50, 100], ["2", "5", "10", "25", "50", "100"])
        a.set_title(f"{dmin} min", fontsize=9, loc="left"); a.set_xlabel("T (años)")
    ax[0].set_ylabel("intensidad (mm/h)"); ax[0].legend(fontsize=7.5)
    fig.tight_layout(); fig.savefig(P2 / "fig_idf_comparacion.png"); plt.close(fig)


def clima():
    MU, SG, XI, TC = 19.42, 7.52, 0.044, 102 / 60
    den = (TC + 0.137) ** 0.872
    yT = lambda T: ((-np.log(1 - 1 / T)) ** (-XI) - 1) / XI
    Ts = np.logspace(np.log10(2), np.log10(500), 80)
    cap = 1.695 * 3.6 / (0.335 * 0.437 * 0.976)
    fig, a = plt.subplots(figsize=(6.4, 3.6))
    for (dt, r, ds), lab, col, ls in (((0, .07, 0), "hoy (IDF 2.3)", C_PL, "-"), ((1.5, .07, 0), "+1,5 °C a 7 %/°C", C_E, "--"),
                                       ((2, .14, 0), "+2 °C a 14 %/°C", "#e34948", "--"), ((0, .07, .25), "σ +25 %, misma media", C_ET, ":")):
        f = (1 + r) ** dt
        a.plot(Ts, (MU * f + SG * f * (1 + ds) * yT(Ts)) / den, color=col, ls=ls, lw=2, label=lab)
    a.axhline(cap, color=TINTA, ls=":", lw=1.2); a.text(2.2, cap + 0.6, "capacidad del DN 1 000", fontsize=8)
    a.axvline(25, color=GR, lw=0.8); a.axvspan(60, 500, color=GR, alpha=0.08); a.text(70, 22, "T > 2N: extrapolación", fontsize=7.5, color=GR)
    a.set_xscale("log"); a.set_xticks([2, 5, 10, 25, 50, 100, 200, 500], ["2", "5", "10", "25", "50", "100", "200", "500"])
    a.set_xlabel("periodo de retorno T (años)"); a.set_ylabel("i en t$_c$ = 102 min (mm/h)"); a.legend(fontsize=8)
    fig.tight_layout(); fig.savefig(P2 / "fig_clima_escenarios.png"); plt.close(fig)


if __name__ == "__main__":
    balance(); presidenta(); idf(); clima()
    print("figuras del informe generadas")
