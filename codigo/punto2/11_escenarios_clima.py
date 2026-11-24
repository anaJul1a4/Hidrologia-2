"""Escenarios de cambio climático para la alcantarilla del problema A (bono 2.5.3, análisis de sensibilidad).

Escala la IDF de 2.3, i(d,T) = a(T)/(d + θ)^η con a(T) una GEV (μ, σ, ξ), con f = (1 + r)^ΔT
(Clausius–Clapeyron, r ≈ 7 %/°C; Westra et al., 2014) y un cambio extra en la escala σ:
    μ' = μ·f,   σ' = σ·f·(1 + Δσ).
Para cada escenario calcula la intensidad de diseño (t_c = 102 min, T = 25), el caudal racional, el diámetro,
el periodo de retorno efectivo del DN 1 000 instalado y su riesgo en 30 años. Es el mismo cálculo que hace la
diapositiva «p25-clima» en el navegador.
Salida: productos/punto2/escenarios_clima.csv.   Ejecutar desde la raíz:  python codigo/punto2/11_escenarios_clima.py
"""
from pathlib import Path
import csv
import math

RAIZ = Path(__file__).resolve().parents[2]
OUT = RAIZ / "productos" / "punto2"
OUT.mkdir(parents=True, exist_ok=True)

MU, SG, XI, TH, ETA = 19.42, 7.52, 0.044, 0.137, 0.872      # IDF de 2.3 (d en h, i en mm/h)
TC, C, A, FRA, N_MAN, S0, QCAP, VIDA, TD = 102 / 60, 0.335, 0.437, 0.976, 0.013, 0.005, 1.695, 30, 25
DN = [0.6, 0.8, 1.0, 1.2, 1.4, 1.5, 1.6, 1.8, 2.0]


def y_t(T):
    return ((-math.log(1 - 1 / T)) ** (-XI) - 1) / XI


def t_de(a, mu, sg):
    z = 1 + XI * (a - mu) / sg
    return 1 / (1 - math.exp(-z ** (-1 / XI)))


def escenario(dt, r, ds):
    f = (1 + r) ** dt
    mu, sg = MU * f, SG * f * (1 + ds)
    den = (TC + TH) ** ETA
    i25 = (mu + sg * y_t(TD)) / den
    q = C * i25 * A * FRA / 3.6
    d = (N_MAN * q / (0.312 * math.sqrt(S0))) ** 0.375
    a_cap = QCAP * 3.6 / (C * A * FRA) * den
    t_ef = t_de(a_cap, mu, sg)
    return {"dT_C": dt, "tasa_por_C": r, "cambio_sigma": ds, "factor_f": round(f, 4),
            "i_tc_T25_mm_h": round(i25, 2), "Q_diseno_m3_s": round(q, 3), "D_calculado_m": round(d, 3),
            "DN_comercial_m": next(v for v in DN if v >= d - 1e-9),
            "T_efectivo_DN1000_anios": round(t_ef), "riesgo_30_anios": round(1 - (1 - 1 / t_ef) ** VIDA, 3),
            "T_nuevo_de_la_lluvia_de_25_hoy": round(t_de(MU + SG * y_t(TD), mu, sg), 1)}


def main():
    filas = [escenario(*e) for e in [(0, 0.07, 0), (1, 0.07, 0), (1.5, 0.07, 0), (2, 0.07, 0), (3, 0.07, 0),
                                     (2, 0.14, 0), (0, 0.07, 0.25), (1.5, 0.07, 0.25)]]
    with open(OUT / "escenarios_clima.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(filas[0])); w.writeheader(); w.writerows(filas)
    for r in filas:
        print(r)
    s25 = SG * y_t(25) / (MU + SG * y_t(25)); s100 = SG * y_t(100) / (MU + SG * y_t(100))
    print(f"Peso del término de escala en a(T): T = 25 → {s25:.0%}, T = 100 → {s100:.0%}")


if __name__ == "__main__":
    main()
