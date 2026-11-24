"""Delimitación de la quebrada La Presidenta (Medellín) con Copernicus GLO-30, como en la Tarea 1, y tiempos de
concentración de tres cierres:
  · «parte alta»: unidad hidrológica 2701-01-130-004 de Corantioquia (0,593 km², cierre −75.536, 6.207);
  · estación de nivel SIATA 101 (Parque Lineal, −75.5694, 6.2079): la que se usa para medir la respuesta;
  · salida al río Medellín (2701-01-130, 13,38 km², cierre −75.578, 6.213).
El cierre se ajusta a la celda cuya área acumulada se acerca más al área oficial (o a la de mayor área, para 101)
dentro de un radio de búsqueda. Limitación: GLO-30 es un modelo de superficie (incluye edificios y dosel).

Entradas: datos/punto2/dem_crudo/Copernicus_DSM_COG_10_N06_00_W076_00_DEM.tif (AWS Open Data, fuera de git)
Salidas:  productos/punto2/presidenta_cuencas.csv, productos/punto2/presidenta_cuencas.geojson
Uso: python codigo/punto2/13_cuencas_presidenta.py   (requiere rasterio, pysheds, geopandas, shapely)
"""
from pathlib import Path
import csv
import json
import math
import sys

import numpy as np
if not hasattr(np, "in1d"):          # compatibilidad pysheds 0.5 con NumPy 2 (igual que en la Tarea 1)
    np.in1d = np.isin
import rasterio
from rasterio.warp import calculate_default_transform, reproject, Resampling
from rasterio.windows import from_bounds
from rasterio.features import shapes
from shapely.geometry import shape, mapping
from shapely.ops import unary_union, transform as shp_transform
from pyproj import Transformer
from pysheds.grid import Grid

RAIZ = Path(__file__).resolve().parents[2]
DEM = RAIZ / "datos" / "punto2" / "dem_crudo" / "Copernicus_DSM_COG_10_N06_00_W076_00_DEM.tif"
TMP = RAIZ / "datos" / "punto2" / "dem_crudo" / "presidenta_utm18n.tif"
OUT = RAIZ / "productos" / "punto2"
OUT.mkdir(parents=True, exist_ok=True)
BBOX = (-75.62, 6.15, -75.49, 6.25)
CRS = "EPSG:32618"
CIERRES = [  # nombre, lon, lat, área oficial km² (None = tomar la de mayor acumulación), radio de búsqueda m
    ("parte_alta", -75.536, 6.207, 0.592728, 250),
    ("siata_101", -75.5694, 6.2079, None, 120),
    ("salida", -75.578, 6.213, 13.382196, 400),
]


def recortar_y_proyectar():
    with rasterio.open(DEM) as src:
        win = from_bounds(*BBOX, transform=src.transform)
        data = src.read(1, window=win)
        tr = src.window_transform(win)
        t2, w, h = calculate_default_transform(src.crs, CRS, data.shape[1], data.shape[0], *BBOX, resolution=30)
        dst = np.full((h, w), -9999, dtype="float32")
        reproject(data, dst, src_transform=tr, src_crs=src.crs, dst_transform=t2, dst_crs=CRS,
                  resampling=Resampling.bilinear, dst_nodata=-9999)
        prof = dict(driver="GTiff", height=h, width=w, count=1, dtype="float32", crs=CRS, transform=t2, nodata=-9999)
    with rasterio.open(TMP, "w", **prof) as d:
        d.write(dst, 1)


def tc_formulas(L_m, S, A_km2, H_m):
    """Tiempos de concentración en minutos. L longitud del cauce principal (m), S pendiente (m/m),
    A área (km²), H desnivel medio de la cuenca sobre el cierre (m, para Giandotti)."""
    out = {"Kirpich": 0.0195 * L_m ** 0.77 * S ** -0.385,                         # Kirpich (1940)
           "Temez": 60 * 0.3 * ((L_m / 1000) / S ** 0.25) ** 0.76,                 # Témez (1978)
           "Bransby_Williams": 14.6 * L_m / 1000 * A_km2 ** -0.1 * S ** -0.2}       # Bransby-Williams (L km, A km²)
    if H_m > 0:
        out["Giandotti"] = 60 * (4 * math.sqrt(A_km2) + 1.5 * L_m / 1000) / (0.8 * math.sqrt(H_m))   # Giandotti (1934)
    return out


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    if not DEM.exists():
        raise FileNotFoundError(f"Falta el DEM {DEM}: descárguelo de copernicus-dem-30m (AWS) como indica el README")
    recortar_y_proyectar()
    grid = Grid.from_raster(str(TMP))
    dem = grid.read_raster(str(TMP))
    inf = grid.resolve_flats(grid.fill_depressions(grid.fill_pits(dem)))
    fdir = grid.flowdir(inf)
    acc = grid.accumulation(fdir)
    cel = abs(grid.affine.a * grid.affine.e) / 1e6     # km² por celda
    a_utm = Transformer.from_crs(4326, CRS, always_xy=True)
    a_geo = Transformer.from_crs(CRS, 4326, always_xy=True)
    filas, feats = [], []
    for nombre, lon, lat, a_of, radio in CIERRES:
        x0, y0 = a_utm.transform(lon, lat)
        c0, r0 = ~grid.affine * (x0, y0)
        n = int(radio / 30) + 1
        mejor = None
        for r in range(int(r0) - n, int(r0) + n + 1):
            for c in range(int(c0) - n, int(c0) + n + 1):
                if math.hypot(r - r0, c - c0) * 30 > radio:
                    continue
                a = acc[r, c] * cel
                crit = -a if a_of is None else abs(a - a_of)
                if mejor is None or crit < mejor[0]:
                    mejor = (crit, r, c, a)
        _, r, c, a = mejor
        xs, ys = grid.affine * (c + 0.5, r + 0.5)
        catch = grid.catchment(x=c, y=r, fdir=fdir, xytype="index")
        dist = grid.distance_to_outlet(x=c, y=r, fdir=fdir, xytype="index")   # en pasos de celda D8
        print(nombre, "celdas de la cuenca:", int(np.asarray(catch).sum()), "→", round(int(np.asarray(catch).sum()) * cel, 3), "km² (acumulación:", round(a, 3), ")")
        dist_m = np.where(catch, dist * 30.0, np.nan)
        L = float(np.nanmax(dist_m))
        rr, cc = np.unravel_index(np.nanargmax(dist_m), dist_m.shape)
        z_out, z_far = float(inf[r, c]), float(inf[rr, cc])
        S = (z_far - z_out) / L
        zc = np.where(catch, inf, np.nan)
        z_mean = float(np.nanmean(zc))
        tcs = tc_formulas(L, S, a, z_mean - z_out)
        lon_o, lat_o = a_geo.transform(xs, ys)
        fila = {"cierre": nombre, "lon": round(lon_o, 5), "lat": round(lat_o, 5), "area_km2": round(a, 3),
                "area_oficial_km2": a_of, "L_cauce_m": round(L), "S_cauce_m_m": round(S, 4),
                "z_cierre_m": round(z_out), "z_max_m": round(float(np.nanmax(zc))), "z_media_m": round(z_mean),
                **{f"tc_{k}_min": round(v, 1) for k, v in tcs.items()}}
        filas.append(fila)
        print(fila)
        geom = unary_union([shape(g) for g, v in shapes(catch.astype("uint8"), mask=catch.astype(bool), transform=grid.affine) if v == 1])
        feats.append({"type": "Feature", "properties": {"cierre": nombre, "area_km2": round(a, 3)},
                      "geometry": mapping(shp_transform(lambda x, y, z=None: a_geo.transform(x, y), geom.simplify(15)))})
    with open(OUT / "presidenta_cuencas.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(filas[0]) + [k for k in filas[1] if k not in filas[0]], extrasaction="ignore")
        w.writeheader(); w.writerows(filas)
    (OUT / "presidenta_cuencas.geojson").write_text(json.dumps({"type": "FeatureCollection", "features": feats}), encoding="utf-8")


if __name__ == "__main__":
    main()
