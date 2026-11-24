"""Descarga los históricos de SIATA (Dataverse datos.siata.gov.co, sin registro) que usa la comparación con
la quebrada La Presidenta (Medellín):
  101 · Q. La Presidenta – Parque Lineal (nivel, 1 min)     doi:10.83041/JEO98W   (−75.5694, 6.2079)
  41  · Comisaría El Poblado (lluvia, 1 min, p1 y p2)        doi:10.83041/8E02OC   (−75.5662, 6.2067)
Los archivos mensuales crudos (≈ 480 MB) quedan en datos/punto2/siata_crudo/ (fuera de git; el ZIP de entrega
lleva solo los productos procesados que genera 13_presidenta.py). Si una descarga falla, el script lo informa y
sigue: no rellena nada.
Uso: python codigo/punto2/12_descarga_siata.py
"""
from pathlib import Path
import json
import time
import urllib.request

RAIZ = Path(__file__).resolve().parents[2]
OUT = RAIZ / "datos" / "punto2" / "siata_crudo"
OUT.mkdir(parents=True, exist_ok=True)
API = "https://datos.siata.gov.co/api"
UA = {"User-Agent": "hydrology-coursework-client"}
ESTACIONES = {"101_nivel": "doi:10.83041/JEO98W", "41_lluvia": "doi:10.83041/8E02OC"}


def get(url, timeout=180):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=timeout) as r:
        return r.read()


def main():
    fallas = []
    for nombre, doi in ESTACIONES.items():
        meta = json.loads(get(f"{API}/datasets/:persistentId/?persistentId={doi}"))
        archivos = meta["data"]["latestVersion"]["files"]
        carpeta = OUT / nombre
        carpeta.mkdir(exist_ok=True)
        (carpeta / "metadatos_dataverse.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
        for a in sorted(archivos, key=lambda x: x["dataFile"]["filename"]):
            fn = a["dataFile"]["filename"].replace(".tab", ".csv")
            if "_tm_" in fn:            # serie paralela de otra configuración del sensor (2014); no se usa
                continue
            dest = carpeta / fn
            if dest.exists() and dest.stat().st_size > 0:
                continue
            try:
                dest.write_bytes(get(f"{API}/access/datafile/{a['dataFile']['id']}?format=original"))
            except Exception as e:  # noqa: BLE001 – se reporta y se sigue
                fallas.append((fn, str(e)))
                print("FALLA", fn, e)
            time.sleep(0.2)
        print(nombre, "archivos:", len(list(carpeta.glob("*.csv"))))
    print("descargas fallidas:", len(fallas))
    for f in fallas:
        print("  ", *f)


if __name__ == "__main__":
    main()
