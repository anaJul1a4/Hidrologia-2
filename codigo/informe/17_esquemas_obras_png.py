"""Exporta a PNG los esquemas SVG de las obras (diapositivas p24-obra y p24b-obra de index.html) para el informe.
Usa Chrome o Edge sin interfaz. Salidas: productos/punto2/fig_obra_alcantarilla.png, fig_obra_tanque.png
Uso: python codigo/informe/17_esquemas_obras_png.py
"""
from pathlib import Path
import re
import shutil
import subprocess
import tempfile

RAIZ = Path(__file__).resolve().parents[2]
OUT = RAIZ / "productos" / "punto2"
NAV = next((p for p in [shutil.which("chrome"), r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", shutil.which("google-chrome"),
                        shutil.which("chromium")] if p and Path(p).exists()), None)
CSS = """:root{--ink:#0b0b0b;--ink2:#52514e;--muted:#898781;--line:#e1e0d9;--axis:#c3c2b7;--card:#fcfcfb;--well:#f0efec;
--acc:#2a78d6;--sw:#b07a0c;--vPL:#2a78d6}body{margin:0;background:#fff;font-family:system-ui,Segoe UI,sans-serif}
svg{display:block;width:100%;height:auto;overflow:visible}svg text{font-family:system-ui,Segoe UI,sans-serif}
svg text:not([fill]){fill:#0b0b0b}svg text:not([font-size]){font-size:12px}svg text.mut{fill:#898781}
.row{display:grid;grid-template-columns:2.6fr 1fr;gap:12px;align-items:center;padding:10px}"""


def svgs(slide_id):
    s = (RAIZ / "index.html").read_text(encoding="utf-8")
    sec = s[s.index(f"id='{slide_id}'"):]
    sec = sec[:sec.index("</section>")]
    return re.findall(r"<svg viewBox=.*?</svg>", sec, re.S)


def captura(html, salida, ancho, alto):
    with tempfile.TemporaryDirectory() as td:
        f = Path(td) / "e.html"
        f.write_text(f"<!doctype html><meta charset='utf-8'><style>{CSS}</style>{html}", encoding="utf-8")
        subprocess.run([NAV, "--headless=new", "--hide-scrollbars", "--force-device-scale-factor=2", f"--window-size={ancho},{alto}",
                        f"--screenshot={salida}", f.as_uri()], check=True, capture_output=True, timeout=120)


def main():
    if NAV is None:
        raise SystemExit("No se encontró Chrome ni Edge para exportar los SVG")
    a = svgs("p24-obra")
    captura(f"<div class='row'>{a[0]}{a[1]}</div>", OUT / "fig_obra_alcantarilla.png", 1400, 420)
    b = svgs("p24b-obra")
    captura(f"<div style='padding:10px 20px'>{b[0]}</div>", OUT / "fig_obra_tanque.png", 1100, 380)
    print("esquemas exportados")


if __name__ == "__main__":
    main()
