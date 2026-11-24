"""Balance de humedad del suelo diario (cubeta tipo Thornthwaite–Mather con escorrentía tipo HBV)
para la cuenca del río Chicama hasta El Tambo, 1998-01 a 2016-12.

Es la versión con datos de la cuenca de la figura #s14 de las notas del curso
(Hoyos, 2026, «El balance hídrico: P = Q + ET + ΔS»). Sirve para la diapositiva
«p1-balance» del sitio y responde de forma gráfica la pregunta del punto 1:
cuánta demanda hay (ETp), cuánta se satisface (ET), qué controla la diferencia
mes a mes y entre años, y qué tan sensible es a la fuente de lluvia y de demanda.

Entradas (rutas relativas a la raíz del repositorio):
  datos/punto1/camels_pe_PE_47E0A2A8_diario_1998_2016.csv  P_L (PISCOp v2.1), PET (PISCOeo_pm), caudal obs.
  datos/punto1/tarea1/imerg_cuenca_mensual.csv              P_I mensual (IMERG V07 Final), Tarea 1
  datos/punto1/tarea1/series_mensuales_camels_pe.csv        R mensual rellenado (Tarea 1)
  datos/punto1/eto_era5land_diaria_cuenca.csv  (opcional)   ETo FAO-56 diaria de ERA5-Land (columnas date, eto_mm)
Salidas:
  productos/punto1/balance_cubeta_datos.js         datos que lee index.html (window.BAL)
  productos/punto1/balance_cubeta_mensual.csv      resultado mensual con los parámetros por defecto
  productos/punto1/balance_cubeta_anual.csv        resultado por año hidrológico (oct–sep)

Modelo, paso diario (se agrega a mes al final, como pide la guía):
  1. Escorrentía rápida: Pr = P · (S/Smax)^β   (Bergström, 1992, HBV): escurre más cuando el suelo está húmedo.
  2. Infiltra Pi = P − Pr. Si Pi ≥ ETp: ET = ETp, el suelo se llena y el exceso sobre Smax recarga.
     Si Pi < ETp: ET = Pi + S·[1 − exp(−(ETp − Pi)/Smax)]   (Thornthwaite y Mather, 1955).
  3. Pr y la recarga entran a un depósito lineal G; el caudal es Q = k·G.
Ejecutar desde la raíz:  python codigo/punto1/00_balance_cubeta.py
"""
from pathlib import Path
import csv
import json
import math
from collections import defaultdict

RAIZ = Path(__file__).resolve().parents[2]
DAT = RAIZ / "datos" / "punto1"
OUT = RAIZ / "productos" / "punto1"
OUT.mkdir(parents=True, exist_ok=True)

# Parámetros por defecto (se justifican en la diapositiva y en el informe)
SMAX, BETA, K = 100.0, 1.0, 0.05      # mm, adimensional, 1/día
ERA5_SOBRE_PISCOEO = 1210 / 1335      # cociente de totales anuales 1981–2020 (apartado 1.1.b del sitio)


def leer_diario():
    f = DAT / "camels_pe_PE_47E0A2A8_diario_1998_2016.csv"
    filas = list(csv.DictReader(open(f, encoding="utf-8")))
    fechas = [r["date"] for r in filas]
    p = [float(r["prec_mm"]) for r in filas]
    pet = [float(r["pet_mm"]) for r in filas]
    if any(math.isnan(v) for v in pet):
        raise ValueError("Hay PET faltante en el periodo: revisar la fuente")
    return fechas, p, pet


def leer_imerg():
    out = {}
    for r in csv.DictReader(open(DAT / "tarea1" / "imerg_cuenca_mensual.csv", encoding="utf-8")):
        out[r["fecha"][:7]] = float(r["P_I_mm_mes"])
    return out


def leer_r_obs():
    out = {}
    for r in csv.DictReader(open(DAT / "tarea1" / "series_mensuales_camels_pe.csv", encoding="utf-8")):
        if r["R_mm_mes"] not in ("", "NA"):
            out[r["fecha"][:7]] = float(r["R_mm_mes"])
    return out


def leer_eto_era5():
    f = DAT / "eto_era5land_diaria_cuenca.csv"
    if not f.exists():
        return None
    return {r["date"]: float(r["eto_mm"]) for r in csv.DictReader(open(f, encoding="utf-8"))}


def desagregar_imerg(fechas, p_l, imerg):
    """P_I diaria: reparte el total mensual de IMERG con el patrón diario de P_L
    (uniforme si P_L del mes es casi cero)."""
    tot_l = defaultdict(float)
    ndias = defaultdict(int)
    for d, v in zip(fechas, p_l):
        tot_l[d[:7]] += v
        ndias[d[:7]] += 1
    out = []
    for d, v in zip(fechas, p_l):
        m = d[:7]
        if m not in imerg:
            raise KeyError(f"Falta IMERG en {m}")
        out.append(imerg[m] * v / tot_l[m] if tot_l[m] > 1.0 else imerg[m] / ndias[m])
    return out


def cubeta(p, etp, smax=SMAX, beta=BETA, k=K, vueltas=2):
    """Devuelve listas diarias de ET, Q, S y recarga. La primera vuelta sirve de calentamiento."""
    s, g = 0.5 * smax, 50.0
    for _ in range(vueltas):
        et_l, q_l, s_l = [], [], []
        for pp, ee in zip(p, etp):
            pr = pp * (s / smax) ** beta
            pi = pp - pr
            if pi >= ee:
                et = ee
                s += pi - ee
                rec = max(0.0, s - smax)
                s = min(s, smax)
            else:
                sn = s * math.exp(-(ee - pi) / smax)
                et = pi + s - sn
                s = sn
                rec = 0.0
            g += pr + rec
            q = k * g
            g -= q
            et_l.append(et); q_l.append(q); s_l.append(s)
    return et_l, q_l, s_l


def mensual(fechas, **series):
    acc = defaultdict(lambda: defaultdict(float))
    for i, d in enumerate(fechas):
        for nom, v in series.items():
            if nom == "S":
                acc[d[:7]]["S"] = v[i]          # almacenamiento al final del mes
            else:
                acc[d[:7]][nom] += v[i]
    return acc


def main():
    try:
        import sys
        sys.stdout.reconfigure(encoding="utf-8")   # consola de Windows (cp1252)
    except AttributeError:
        pass
    fechas, p_l, pet = leer_diario()
    imerg = leer_imerg()
    r_obs = leer_r_obs()
    p_i = desagregar_imerg(fechas, p_l, imerg)
    eto = leer_eto_era5()
    if eto is not None:
        dem2 = [eto[d] for d in fechas]
        dem2_nombre = "ETo FAO-56 ERA5-Land (serie diaria)"
    else:
        dem2 = [v * ERA5_SOBRE_PISCOEO for v in pet]
        dem2_nombre = f"PISCOeo_pm × {ERA5_SOBRE_PISCOEO:.3f} (total de ERA5-Land FAO-56)"

    meses = sorted({d[:7] for d in fechas})
    faltan = [m for m in meses if m not in r_obs]
    print(f"Días: {len(fechas)} ({fechas[0]} a {fechas[-1]}); meses sin R observado: {faltan or 'ninguno'}")

    datos = {
        "fechas0": fechas[0], "n": len(fechas),
        "PL": [round(v, 2) for v in p_l], "PI": [round(v, 2) for v in p_i],
        "E1": [round(v, 2) for v in pet], "E2": [round(v, 2) for v in dem2],
        "E2nombre": dem2_nombre, "meses": meses, "Robs": [r_obs.get(m) for m in meses],
        "def": {"smax": SMAX, "beta": BETA, "k": K},
    }
    (OUT / "balance_cubeta_datos.js").write_text(
        "/* generado por codigo/punto1/00_balance_cubeta.py — no editar a mano */\nwindow.BAL = "
        + json.dumps(datos, separators=(",", ":")) + ";\n", encoding="utf-8")

    et, q, s = cubeta(p_l, pet)
    m = mensual(fechas, P=p_l, ETp=pet, ET=et, Q=q, S=s)
    with open(OUT / "balance_cubeta_mensual.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["mes", "P_L_mm", "ETp_PISCOeo_mm", "ET_sim_mm", "Q_sim_mm", "R_obs_mm", "S_fin_mm"])
        for mm in meses:
            v = m[mm]
            w.writerow([mm, f"{v['P']:.1f}", f"{v['ETp']:.1f}", f"{v['ET']:.1f}", f"{v['Q']:.1f}",
                        "" if r_obs.get(mm) is None else f"{r_obs[mm]:.1f}", f"{v['S']:.1f}"])

    # año hidrológico octubre–septiembre (empieza con el almacenamiento mínimo, como en 1.2)
    anual = defaultdict(lambda: defaultdict(float))
    for mm in meses:
        y, mo = int(mm[:4]), int(mm[5:])
        ah = y if mo >= 10 else y - 1
        for c in ("P", "ETp", "ET", "Q"):
            anual[ah][c] += m[mm][c]
        anual[ah]["R"] += r_obs.get(mm, float("nan"))
        anual[ah]["n"] += 1
    with open(OUT / "balance_cubeta_anual.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["anio_hidrologico_oct_sep", "P_L_mm", "ETp_mm", "ET_sim_mm", "Q_sim_mm", "R_obs_mm", "P_menos_R_mm"])
        for ah in sorted(anual):
            v = anual[ah]
            if v["n"] == 12:
                w.writerow([f"{ah}-{ah + 1}", *(f"{v[c]:.0f}" for c in ("P", "ETp", "ET", "Q", "R")), f"{v['P'] - v['R']:.0f}"])

    anios = len(fechas) / 365.25
    tot = {c: sum(m[mm][c] for mm in meses) / anios for c in ("P", "ETp", "ET", "Q")}
    robs = sum(v for v in (r_obs.get(mm) for mm in meses) if v is not None) / anios
    clim_sim = [sum(m[mm]["Q"] for mm in meses if int(mm[5:]) == i) / anios for i in range(1, 13)]
    clim_obs = [sum(r_obs[mm] for mm in meses if int(mm[5:]) == i) / anios for i in range(1, 13)]
    mae = sum(abs(a - b) for a, b in zip(clim_sim, clim_obs)) / 12
    print(f"Medias anuales 1998–2016 (mm/año): P {tot['P']:.0f}, ETp {tot['ETp']:.0f}, ET {tot['ET']:.0f}, "
          f"Q sim {tot['Q']:.0f}, R obs {robs:.0f}, P − R {tot['P'] - robs:.0f}")
    print(f"ET/ETp {tot['ET'] / tot['ETp']:.2f}, ET/P {tot['ET'] / tot['P']:.2f}; MAE del ciclo de Q {mae:.1f} mm/mes")
    print(f"Demanda alternativa: {dem2_nombre}")


if __name__ == "__main__":
    main()
