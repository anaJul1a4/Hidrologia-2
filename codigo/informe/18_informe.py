"""
Informe de la Tarea 2 en PDF y en Word editable, generados del mismo texto (misma maquetación que el informe de la
Tarea 1, script 40_informe_pdf.py de ese repositorio). Reúne el texto, las figuras de img/ y productos/ y las cifras de
los CSV. Requiere reportlab, python-docx, pandas, Pillow y matplotlib (fuente DejaVu Sans).

Salidas (raíz): Informe_Tarea2_Chicama_Cottbus.pdf y .docx. Para el ZIP, el PDF se renombra T2_202602_GNN_informe.pdf.
El .docx es la versión para editar en grupo; si se edita a mano, ese archivo pasa a ser la fuente del informe.
Uso: python codigo/informe/18_informe.py
"""
import html as _html
import re
import tempfile
from pathlib import Path

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor
import matplotlib
import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak,
                                KeepTogether, CondPageBreak)

PROY = Path(__file__).resolve().parents[2]
SALIDA = PROY / "Informe_Tarea2_Chicama_Cottbus.pdf"
SALIDA_DOCX = PROY / "Informe_Tarea2_Chicama_Cottbus.docx"
PIE = "Hidrología 2026-2 · Tarea 2 · Chicama (Perú) y Cottbus (Alemania)"

TTF = Path(matplotlib.__file__).parent / "mpl-data" / "fonts" / "ttf"
for nombre, archivo in (("DV", "DejaVuSans.ttf"), ("DV-B", "DejaVuSans-Bold.ttf"), ("DV-I", "DejaVuSans-Oblique.ttf"),
                        ("DV-BI", "DejaVuSans-BoldOblique.ttf")):
    pdfmetrics.registerFont(TTFont(nombre, str(TTF / archivo)))
pdfmetrics.registerFontFamily("DV", normal="DV", bold="DV-B", italic="DV-I", boldItalic="DV-BI")

AZUL = colors.HexColor("#184f95")
GRIS = colors.HexColor("#52514e")
st = {
    "titulo": ParagraphStyle("titulo", fontName="DV-B", fontSize=17, leading=22, alignment=TA_CENTER, spaceAfter=6),
    "subtitulo": ParagraphStyle("sub", fontName="DV", fontSize=11, leading=15, alignment=TA_CENTER, textColor=GRIS),
    "h1": ParagraphStyle("h1", fontName="DV-B", fontSize=13.5, leading=18, spaceBefore=14, spaceAfter=6, textColor=AZUL, keepWithNext=1),
    "h2": ParagraphStyle("h2", fontName="DV-B", fontSize=11, leading=15, spaceBefore=10, spaceAfter=4, keepWithNext=1),
    "p": ParagraphStyle("p", fontName="DV", fontSize=9.4, leading=13.6, alignment=TA_JUSTIFY, spaceAfter=6),
    "bul": ParagraphStyle("bul", fontName="DV", fontSize=9.4, leading=13.4, alignment=TA_JUSTIFY, leftIndent=14,
                          bulletIndent=4, spaceAfter=3),
    "cap": ParagraphStyle("cap", fontName="DV", fontSize=8.2, leading=11, textColor=GRIS, alignment=TA_JUSTIFY,
                          spaceBefore=3, spaceAfter=10),
    "celda": ParagraphStyle("celda", fontName="DV", fontSize=7.8, leading=10),
    "celdab": ParagraphStyle("celdab", fontName="DV-B", fontSize=7.8, leading=10),
    "ref": ParagraphStyle("ref", fontName="DV", fontSize=8.3, leading=11.2, leftIndent=14, firstLineIndent=-14,
                          spaceAfter=3),
    "aviso": ParagraphStyle("aviso", fontName="DV-I", fontSize=8.6, leading=12, textColor=GRIS, alignment=TA_JUSTIFY,
                            borderColor=GRIS, borderWidth=0.5, borderPadding=6, spaceBefore=6, spaceAfter=10),
}
ANCHO = letter[0] - 4 * cm
story = []

# ---------------------------------------------------------------------- Word (python-docx), en paralelo al PDF
W = docx.Document()
_sec = W.sections[0]
_sec.page_width, _sec.page_height = Cm(21.59), Cm(27.94)                 # carta, igual que el PDF
_sec.left_margin = _sec.right_margin = Cm(2)
_sec.top_margin = _sec.bottom_margin = Cm(1.8)
ANCHO_W = Cm(21.59 - 4)
_n = W.styles["Normal"]
_n.font.name = "Calibri"; _n.font.size = Pt(10.5)
_n.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
_n.paragraph_format.space_after = Pt(5); _n.paragraph_format.line_spacing = 1.15
for _nivel, _tam, _col in ((1, 14, RGBColor(0x18, 0x4F, 0x95)), (2, 11.5, RGBColor(0x0B, 0x0B, 0x0B))):
    _h = W.styles[f"Heading {_nivel}"]
    _h.font.name = "Calibri"; _h.font.size = Pt(_tam); _h.font.bold = True; _h.font.color.rgb = _col
    _h.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
GRIS_W = RGBColor(0x52, 0x51, 0x4E)
_TOKEN = re.compile(r"(<br/>|</?b>|</?i>|</?sub>|</?super>)")


def runs(par, texto, tam=None, color=None, italica=False):
    """Escribe en un párrafo de Word el marcado simple de reportlab (<b>, <i>, <sub>, <super>, <br/>)."""
    b = i = sub = sup = False
    for trozo in _TOKEN.split(texto):
        if trozo == "<br/>":
            par.add_run().add_break(); continue
        if trozo in ("<b>", "</b>"):
            b = trozo == "<b>"; continue
        if trozo in ("<i>", "</i>"):
            i = trozo == "<i>"; continue
        if trozo in ("<sub>", "</sub>"):
            sub = trozo == "<sub>"; continue
        if trozo in ("<super>", "</super>"):
            sup = trozo == "<super>"; continue
        if not trozo:
            continue
        r = par.add_run(_html.unescape(trozo))
        r.bold = b or None; r.italic = (i or italica) or None
        if sub: r.font.subscript = True                         # (asignar None a una borra la otra)
        if sup: r.font.superscript = True
        if tam: r.font.size = Pt(tam)
        if color is not None: r.font.color.rgb = color


def W_par(texto, estilo="p"):
    par = W.add_paragraph()
    if estilo == "ref":
        par.paragraph_format.left_indent = Cm(0.6); par.paragraph_format.first_line_indent = Cm(-0.6)
        runs(par, texto, tam=9)
    elif estilo == "aviso":
        runs(par, texto, tam=9.5, color=GRIS_W, italica=True)
        _borde(par)
    elif estilo == "cap":
        runs(par, texto, tam=9, color=GRIS_W)
    else:
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        runs(par, texto)
    return par


def _borde(par):
    pPr = par._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    for lado in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{lado}")
        for k, v in (("val", "single"), ("sz", "4"), ("space", "4"), ("color", "52514E")):
            e.set(qn(f"w:{k}"), v)
        bdr.append(e)
    pPr.append(bdr)


def _sombra(celda, color):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def _campo_pagina(par):
    r = par.add_run()
    for tipo, texto in (("begin", None), (None, "PAGE"), ("end", None)):
        if tipo:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tipo); r._r.append(e)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = texto; r._r.append(e)


def P(t, s="p"):
    story.append(Paragraph(t, st[s]))
    W_par(t, s)


def B(t):
    story.append(Paragraph(t, st["bul"], bulletText="•"))
    par = W.add_paragraph(style="List Bullet")
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    runs(par, t)


def H1(t):
    story.append(Paragraph(t, st["h1"]))
    W.add_heading(_html.unescape(t), level=1)


def H2(t):
    story.append(CondPageBreak(7 * cm))                      # no dejar un subtítulo solo al pie de página
    story.append(Paragraph(t, st["h2"]))
    W.add_heading(_html.unescape(t), level=2)


N_FIG = [0]


_TMP = Path(tempfile.mkdtemp())


def FIG(ruta, leyenda, ancho=1.0, alto_max=17 * cm):
    N_FIG[0] += 1
    ruta = Path(ruta)
    if ruta.suffix.lower() == ".webp":                         # python-docx no lee WebP
        png = _TMP / (ruta.stem + ".png")
        PILImage.open(ruta).convert("RGB").save(png)
        ruta = png
    w, h = PILImage.open(ruta).size
    W = ANCHO * ancho                                           # (variable local; el documento Word es globals()["W"])
    H = W * h / w
    if H > alto_max:
        H, W = alto_max, alto_max * w / h
    story.append(KeepTogether([Image(str(ruta), width=W, height=H),
                               Paragraph(f"<b>Figura {N_FIG[0]}.</b> {leyenda}", st["cap"])]))
    par = globals()["W"].add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.keep_with_next = True
    par.add_run().add_picture(str(ruta), width=int(ANCHO_W * W / ANCHO))
    W_par(f"<b>Figura {N_FIG[0]}.</b> {leyenda}", "cap")


N_TAB = [0]


def TAB(filas, leyenda, anchos, cabecera=True):
    N_TAB[0] += 1
    datos = [[Paragraph(str(c), st["celdab" if (i == 0 and cabecera) else "celda"]) for c in fila]
             for i, fila in enumerate(filas)]
    t = Table(datos, colWidths=[a * ANCHO for a in anchos], repeatRows=1 if cabecera else 0)
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEABOVE", (0, 0), (-1, 0), 0.8, colors.black),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.black),
        ("LINEBELOW", (0, -1), (-1, -1), 0.8, colors.black),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f3ef")]),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]))
    story.append(Paragraph(f"<b>Tabla {N_TAB[0]}.</b> {leyenda}", st["cap"]))
    story.append(t)
    story.append(Spacer(1, 8))
    W_par(f"<b>Tabla {N_TAB[0]}.</b> {leyenda}", "cap").paragraph_format.keep_with_next = True
    tw = W.add_table(rows=len(filas), cols=len(filas[0]))
    tw.style = "Table Grid"
    tw.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, fila in enumerate(filas):
        for j, c in enumerate(fila):
            celda = tw.cell(i, j)
            celda.width = int(ANCHO_W * anchos[j])
            runs(celda.paragraphs[0], str(c) if not (i == 0 and cabecera) else f"<b>{c}</b>", tam=8.5)
            celda.paragraphs[0].paragraph_format.space_after = Pt(0)
            if i == 0 and cabecera:
                _sombra(celda, "E1E0D9")
    W.add_paragraph()


def pie(canvas, doc):
    canvas.saveState()
    canvas.setFont("DV", 7.5)
    canvas.setFillColor(GRIS)
    canvas.drawString(2 * cm, 1.2 * cm, PIE)
    canvas.drawRightString(letter[0] - 2 * cm, 1.2 * cm, f"{doc.page}")
    canvas.restoreState()


IMG, PR1, PR2 = PROY / "img", PROY / "productos" / "punto1", PROY / "productos" / "punto2"


# ====================================================================== portada
TITULO = "Evapotranspiración en la cuenca del río Chicama y diseño hidrológico con lluvia de 10 minutos"
PREGUNTA = ("¿Cuánta agua devuelve a la atmósfera una cuenca andina que la atmósfera querría secar, y qué lluvia debe "
            "soportar una obra de drenaje cuando los extremos pueden intensificarse?")
INTEGR = "Integrantes: Marlin Julie Schwebel · Juan Diego Galvis · [confirmar integrantes y número de grupo]"
SITIOS = ("Punto 1: río Chicama hasta la estación El Tambo (CAMELS-PE PE_47E0A2A8), La Libertad, Perú · 2 170 km², 678–4 308 m "
          "s.n.m. · la misma cuenca de la Tarea 1. Punto 2: lluvia de 10 minutos de Cottbus (DWD 00880), Brandeburgo, Alemania, y "
          "microcuenca de 0,44 km² a 7,2 km del pluviómetro; comparación con la quebrada La Presidenta (Medellín, SIATA).")
WEB = ("Código, datos y página de resultados: https://github.com/anaJul1a4/Hidrologia-2 · "
       "https://anajul1a4.github.io/Hidrologia-2/")
story += [Spacer(1, 2.2 * cm), Paragraph(TITULO, st["titulo"]), Spacer(1, 4), Paragraph(PREGUNTA, st["subtitulo"]),
          Spacer(1, 1.2 * cm),
          Paragraph("Hidrología · Semestre 2026-2 · Universidad Nacional de Colombia, Sede Medellín<br/>Profesor: Carlos David Hoyos",
                    st["subtitulo"]),
          Spacer(1, 0.6 * cm), Paragraph(INTEGR, st["subtitulo"]), Spacer(1, 1.2 * cm), Paragraph(SITIOS, st["subtitulo"]),
          Spacer(1, 1.0 * cm), Paragraph(WEB, st["subtitulo"]), PageBreak()]
for _txt, _tam, _neg, _antes in ((TITULO, 20, True, 120), (PREGUNTA, 12, False, 8),
                                 ("Hidrología · Semestre 2026-2 · Universidad Nacional de Colombia, Sede Medellín<br/>Profesor: Carlos David Hoyos", 11, False, 40),
                                 (INTEGR, 11, False, 18), (SITIOS, 10.5, False, 40), (WEB, 10, False, 30)):
    _p = W.add_paragraph(); _p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _p.paragraph_format.space_before = Pt(_antes)
    runs(_p, f"<b>{_txt}</b>" if _neg else _txt, tam=_tam, color=None if _neg else GRIS_W)
W.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# ====================================================================== 1. Introducción
H1("1. Introducción")
P("La Tarea 1 clasificó la cuenca del río Chicama hasta El Tambo como un régimen pluvial tropical andino de vertiente pacífica: "
  "unimodal, con el 81 % de la lluvia entre diciembre y abril, isotermal (16,2 °C) y subhúmedo seco (P/PET = 0,62). La estación "
  "seca de mayo a septiembre no se debe a falta de humedad sino a la subsidencia del anticiclón del Pacífico sur y a la inversión "
  "térmica asociada, presente en el 91–95 % de los radiosondeos de Trujillo (Klein y Hartmann, 1993; Garreaud, 2009). Esa tarea "
  "dejó una pregunta abierta: de 833 mm de lluvia al año salen 391 por el río, y la escorrentía observada (R/P = 0,48) triplica la "
  "que esperaba la curva de Budyko (0,16). Unos 440 mm/año vuelven a la atmósfera sin que los hayamos medido (Figura 1).")
FIG(IMG / "t1_mapa_cuenca.webp", "Cuenca del río Chicama hasta El Tambo delimitada en la Tarea 1 con Copernicus GLO-30 (D8): relieve, "
    "red de drenaje, río principal, punto de cierre y estaciones de SENAMHI. Tarea 1, script 06.", ancho=0.78)
P("<b>Punto 1.</b> La guía pregunta cuál es la demanda evaporativa de la atmósfera sobre la cuenca, cuánta de esa demanda se "
  "satisface con el agua disponible, qué controla la diferencia a lo largo del año y entre años, y qué tan sensibles son estas "
  "respuestas al método y a la fuente de datos. La cuenca está a 7,6° S, con el Sol alto todo el año, y sobre la inversión: la "
  "demanda debería depender sobre todo de la radiación y de las nubes, y la evapotranspiración real del agua que deja la temporada "
  "de lluvias.")
P("<b>Hipótesis inicial (H0).</b> La cuenca evapora cerca de lo que pide la atmósfera, y el balance de la Tarea 1 no cierra por "
  "errores de la lluvia o del caudal. <b>Hipótesis alternativa (H1).</b> La cuenca está limitada por agua con alternancia estacional: "
  "evapora al ritmo de la energía solo en el pico de lluvias y el resto del año vive del agua almacenada en el suelo. La "
  "concentración de la lluvia cuando la demanda es menor haría que buena parte escurra antes de evaporarse (ω de Fu bajo), sin "
  "necesidad de un error en los datos (Budyko, 1974; Fu, 1981; Zhang et al., 2001).")
P("<b>Punto 2.</b> Diseñar una alcantarilla y un tanque de detención exige lluvia subdiaria. Cerca de El Tambo no hay series largas: "
  "el aeropuerto de Cajamarca solo reporta acumulados de 6 h y las estaciones automáticas de SENAMHI tienen menos de 12 años. Por "
  "eso, siguiendo el orden de criterios de la guía (longitud del registro, resolución y calidad antes que cercanía), el punto 2 se "
  "hace con la serie de 10 minutos de Cottbus (Deutscher Wetterdienst, estación 00880), en la llanura de Brandeburgo, donde las "
  "lluvias que dimensionan obras pequeñas son tormentas convectivas de verano. Para contrastar el resultado con un régimen andino "
  "tropical, se compara con la quebrada La Presidenta, en El Poblado (Medellín), muy urbanizada y monitoreada cada minuto por el "
  "SIATA.")
P("<b>Hipótesis de diseño.</b> Se esperaría que el diámetro dependa sobre todo de la lluvia de diseño (periodo de retorno y "
  "distribución). La alternativa es que, en una cuenca pequeña y plana, pese más el tiempo de concentración, porque define qué "
  "duración de lluvia llena la obra, y que el diámetro comercial cambie más el riesgo real que la propia incertidumbre de la lluvia.")
P("<b>Objetivos.</b> (1) Calcular la ETo FAO-56 de la cuenca y probar su sensibilidad al método, la fuente y la resolución; "
  "(2) estimar la evapotranspiración real y decidir qué la limita cada mes; (3) evaluar los cambios de la demanda y su atribución; "
  "(4) construir la IDF de Cottbus con su incertidumbre; (5) dimensionar las dos obras y su riesgo efectivo; (6) discutir la escala, "
  "el cambio de cobertura, el cambio climático y la diferencia con una quebrada del Valle de Aburrá.")

# ====================================================================== 2. Metodología
H1("2. Metodología")
H2("2.1 Datos")
TAB([["Variable", "Fuente", "Resolución y periodo", "Uso"],
     ["Forzantes de ETo", "ERA5-Land, series horarias por punto (reanalysis-era5-land-timeseries; Muñoz-Sabater et al., 2021)",
      "30 celdas de 0,1°, 1981–2025; 0 horas faltantes", "ETo FAO-56 (1.1)"],
     ["Medias mensuales", "ERA5-Land monthly-means", "0,1°, 1981–2025 (sin 2022-09 a 2024-02)", "Error por promediar, Rn, evaporación, humedad del suelo"],
     ["Estación", "NOAA GSOD, Cajamarca (2 676 m)", "Diaria, 8 694 días pareados", "Prueba de la referencia (1.1.b)"],
     ["Demanda y ET de productos", "PISCOeo_pm (SENAMHI, vía CAMELS-PE); TerraClimate (Abatzoglou et al., 2018)",
      "Diaria/mensual, 1981–2016 y 1981–2020", "Pruebas y ET real (1.1.b, 1.2)"],
     ["Lluvia, caudal, PET diarios", "CAMELS-PE: prec (PISCOp v2.1), flow_obs, pet", "Diaria, 1998–2016 (PET hasta 2016)", "Balance P − Q y balance diario del suelo"],
     ["Lluvia satelital", "IMERG V07 Final (Tarea 1)", "Mensual, 0,1°", "Sensibilidad a la fuente de lluvia"],
     ["Lluvia subdiaria", "DWD, Cottbus 00880", "10 min, 1991–2025; 30 años válidos", "Extremos e IDF (2.1–2.3)"],
     ["IDF oficial", "KOSTRA-DWD-2020, celda 121203", "Duración parcial 1951–2020", "Contraste de la IDF"],
     ["Terreno y cobertura", "DGM1 lidar de Brandeburgo (LGB); ESA WorldCover 2021; ortofoto DOP20; OpenStreetMap",
      "5 m, 10 m, 20 cm", "Microcuenca, C y contexto de la obra"],
     ["Medellín", "SIATA: nivel 101 y lluvia 41 (1 min); Corantioquia (2023); EPM (2025), GD-AS002; Copernicus GLO-30",
      "Nivel 2013–2026; lluvia 11 años válidos", "Comparación con La Presidenta"]],
    "Fuentes de datos. Los subconjuntos usados están en datos/punto1 y datos/punto2; las series crudas de SIATA y el DEM se "
    "regeneran con los scripts de descarga.", [0.16, 0.36, 0.24, 0.24])
H2("2.2 Punto 1: demanda y evapotranspiración real")
B("<b>ETo FAO-56 diaria.</b> Ecuación de Penman–Monteith (Allen et al., 1998, ec. 6) en cada celda y luego promediada con los pesos de "
  "área de la Tarea 1. Temperaturas pasadas a hora local (UTC−5) para no partir la Tmín; e<sub>s</sub> con Tmáx y Tmín; e<sub>a</sub> "
  "del punto de rocío; viento como magnitud horaria √(u²+v²) reducida a 2 m (ec. 47); R<sub>n</sub> del pasto de referencia "
  "(albedo 0,23); γ con la elevación media (2 557 m). Un día se calculó a mano y se verificó con pyet (3,494 mm/d).")
B("<b>Sensibilidad.</b> Tres formas mensuales (M1 con extremos, M2 sin ellos, M3 con el producto mensual), corrección de T y "
  "T<sub>d</sub> a la elevación del DEM, nueve métodos simples (siete de la guía, con Hamon en la forma de Lu et al., 2005, y Oudin et al., 2005, más Makkink, 1957, y Linacre, 1977), un experimento de "
  "perturbación de cada forzante (R<sub>s</sub> −10 %, T +1 °C, u<sub>2</sub> +10 %, e<sub>a</sub> +10 %) y los cocientes "
  "α* = ETo/E<sub>eq</sub>.")
B("<b>ET real.</b> Balance P<sub>L</sub> − Q en 1998–2024 con propagación de error (σ<sub>P</sub> = |P<sub>L</sub> − P<sub>I</sub>|/2, "
  "σ<sub>Q</sub> = 15 %; McMillan et al., 2012), año hidrológico octubre–septiembre, techo energético con R<sub>n</sub> de ERA5-Land, "
  "productos (TerraClimate, ERA5-Land), régimen mensual con ET/ETo (≥ 0,75 energía, ≤ 0,40 agua), Budyko–Fu y elasticidad del "
  "caudal (Sankarasubramanian et al., 2001).")
B("<b>Balance diario del suelo.</b> Cubeta tipo Thornthwaite y Mather (1955) con escorrentía P·(S/S<sub>máx</sub>)<sup>β</sup> "
  "(Bergström, 1992) y depósito lineal Q = k·G, a paso diario con P<sub>L</sub> y PET de CAMELS-PE (1998–2016) y agregado a mes al "
  "final. Valores por defecto S<sub>máx</sub> = 100 mm, β = 1, k = 0,05 d⁻¹; con IMERG, la lluvia mensual se reparte con el patrón "
  "diario de P<sub>L</sub>. Es la figura #s14 de las notas del curso (Hoyos, 2026) con datos de la cuenca.")
B("<b>Cambios.</b> Sen con Mann–Kendall de Hamed y Rao (1998) y control FDR, atribución dETo/dt ≈ Σ(∂ETo/∂x<sub>i</sub>)(dx<sub>i</sub>/dt) "
  "con las sensibilidades mensuales, ventanas móviles y descomposición de Shapley del cambio de caudal con la curva de Fu (bono).")
H2("2.3 Punto 2: extremos, IDF y diseño")
B("<b>Control de calidad.</b> Hora pasada a UTC (MEZ hasta 1999); 22 días con más lluvia en 10 min que en el pluviómetro diario "
  "eliminados; año válido con ≥ 80 % de intervalos y ≥ 90 % en mayo–septiembre (sensibilidad con 29 y 32 años); corrección de "
  "Weiss (1964) por intervalos fijos (1,14 en 10 min).")
B("<b>Extremos.</b> Máximos anuales de 9 duraciones (10 min a 24 h); Gumbel, GEV, log-Pearson III y GLO por momentos, L-momentos y "
  "máxima verosimilitud (Hosking y Wallis, 1997; Coles, 2001); posiciones de Gringorten (1963); KS, Anderson–Darling y PPCC; "
  "bootstrap de 2 000 remuestreos de años completos.")
B("<b>IDF.</b> i(d,T) = a(T)/(d + θ)<sup>η</sup> con a(T) GEV ajustada a las intensidades reescaladas de todas las duraciones "
  "(Koutsoyiannis et al., 1998); contraste con KOSTRA-DWD-2020.")
B("<b>Problema A.</b> Microcuenca delimitada con el DGM1 lidar (Copernicus confundía el dosel con relieve); T por riesgo con "
  "n = 30 años y R = 70 %; t<sub>c</sub> de Kirpich (1940), Témez (1978), velocidad en la zanja y retardo del SCS; C de Chow et "
  "al. (1988, tabla 15.1.1) creciente con T; método racional con factor de reducción areal; Manning a tubo lleno (concreto, "
  "n = 0,013, S₀ = 0,5 %). Monte Carlo con réplicas bootstrap de la IDF, t<sub>c</sub>, C, n y S₀. Hidráulica a sección parcial "
  "(tirantes normal y crítico) para verificar el régimen.")
B("<b>Problema B.</b> Urbanización al 50 % impermeable con t<sub>c</sub> = Kirpich × 0,4; tanque prismático con orificio "
  "(C<sub>d</sub> = 0,6) dimensionado para dejar salir el pico previo; tránsito de piscina nivelada con hidrogramas trapezoidales y "
  "bloques alternos de 6 h con el pico al inicio, al centro y al final (Chow et al., 1988, §8.2, §14.4, §15.4).")
B("<b>La Presidenta.</b> Cuencas delimitadas con GLO-30 y cierres ajustados a las áreas de Corantioquia (2023); 427 crecidas "
  "con subida ≥ 40 cm en la estación de nivel 101 (2013–2026), con la lluvia de la estación 41 en las 4 h previas; tiempo de "
  "subida y retardo respecto al centroide de la lluvia; diseño de la misma obra con la norma de EPM (2025): C = 0,15 + 0,85·I, "
  "t<sub>c</sub> de la FAA, IDF oficial y velocidad máxima de 5 m/s.")
B("<b>Cambio climático (bono 2.5.3).</b> Escalado de la IDF con f = (1 + r)<sup>ΔT</sup> (r ≈ 7 %/°C; Westra et al., 2014) y un "
  "cambio adicional de la escala σ, aplicado a la alcantarilla ya diseñada.")
P("Todas las figuras y tablas salen de los scripts de codigo/punto1, codigo/punto2 y codigo/informe (Python 3, dependencias en "
  "requirements.txt); cada leyenda indica su origen. Las figuras de los apartados 1.1–1.5 y 2.1–2.5 corresponden a las diapositivas "
  "de la página de resultados.")

# ====================================================================== 3. Resultados y discusión
H1("3. Resultados y discusión física")
H2("3.1 La demanda evaporativa: casi constante y dominada por la radiación")
P("La ETo media de la cuenca es 1 211 mm/año (1 127–1 285 mm entre años) y varía poco en el año (88–115 mm/mes; Figura 2). El 90 % "
  "es término radiativo. A 7,6° S el Sol siempre está alto: lo que modula la demanda son las nubes. En febrero–abril la radiación en "
  "el tope de la atmósfera es máxima, pero las nubes de la temporada de lluvias la recortan (R<sub>s</sub>/R<sub>so</sub> = 0,59 en "
  "un día típico de enero) y la ETo baja; el máximo llega en octubre–noviembre, con cielo despejado bajo la subsidencia. En junio–"
  "julio la radiación del tope es mínima, pero el aire más seco (el déficit de vapor se duplica, de 0,37 a 0,80 kPa) sube el término "
  "aerodinámico y los dos efectos se compensan en parte.")
FIG(IMG / "eto_ciclo.webp", "Ciclo anual de la ETo FAO-56 de la cuenca (ERA5-Land horario, 1981–2025), lluvia y radiación. "
    "Script codigo/punto1/03_eto_diaria_cuenca.py (versión de Marlin Schwebel).", ancho=0.9)
P("El mecanismo físico de la estación seca de la Tarea 1 también explica la demanda: de noche, con cielo despejado, el suelo se "
  "enfría y forma una capa estable con inversión de superficie; de día la capa límite convectiva crece y mezcla hacia abajo aire "
  "seco de la troposfera libre (entrañamiento), lo que explica a la vez la mayor amplitud térmica, el rocío más bajo y el déficit "
  "máximo de julio–agosto (Hoyos, 2026, #s36–#s37). Aun así, el viento de ERA5-Land es débil (máximos horarios de 2,6–4,4 m/s en "
  "45 años) y el aire seco casi no se renueva: el término aerodinámico pasa del 6 % de la ETo en febrero al 16 % en julio.")
P("<b>Resolución y forma de cálculo.</b> Calcular la ETo con las medias diarias y luego agregar cambia el resultado muy poco "
  "(±0,5 %), porque la ecuación es casi lineal en R<sub>n</sub>; omitir los extremos de temperatura la subestima 1–3 % (la curva de "
  "saturación es convexa), y usar el producto mensual de ERA5-Land, 2,5–6,4 %, porque al promediar las componentes del viento las "
  "brisas de valle y montaña se cancelan. Además, ese producto tiene un error documentado por ECMWF entre 2022-09 y 2024-02 "
  "(radiación al 31 % de la real), meses que se excluyeron. La elevación del modelo acierta en promedio (2 641 frente a 2 554 m) pero "
  "no tiene ni el fondo del valle ni las cumbres; corregir T y T<sub>d</sub> con gradientes de −5,1 a −5,4 °C/km cambia la ETo solo "
  "+1,1 %, porque T y T<sub>d</sub> bajan juntos.")
H2("3.2 ¿Qué tan buena es la referencia y qué tan sensibles son los métodos?")
P("En el mismo punto y los mismos días, la estación de Cajamarca da una ETo 23 % mayor que ERA5-Land (3,76 frente a 3,06 mm/d). "
  "La causa principal es la amplitud térmica diaria: 14,8 °C en la estación frente a 8,0 °C en el modelo. Una celda de 9 km con "
  "relieve suavizado no forma la inversión nocturna de fondo de valle ni el calentamiento diurno de un valle pequeño. Los productos "
  "dan el mismo orden de magnitud por caminos distintos: TerraClimate llega al mismo total (1 201 mm/año) con 23 % menos radiación y "
  "un viento tres a cuatro veces mayor, una compensación de errores, y PISCOeo_pm es 10 % mayor (hasta 25 % en junio–julio) porque "
  "usa temperaturas de estaciones más cálidas (Figura 3). La demanda de referencia es por tanto ERA5-Land FAO-56 con una banda "
  "de −10 a −25 %.")
FIG(IMG / "eto_productos.webp", "ETo de ERA5-Land (FAO-56), PISCOeo_pm y TerraClimate, 1981–2020. La banda gris (mínimo y máximo "
    "por mes) es la dispersión usada para juzgar los métodos simples.", ancho=0.9)
P("De los nueve métodos simples (Tabla 2, Figura 4), la familia radiativa sigue a FAO-56 porque la demanda es radiativa: Makkink "
  "acierta en el año (−2 %) y en la fase, y Priestley–Taylor con α = 1,26 sobrestima 21 % porque FAO-56 incluye la resistencia del "
  "pasto y el aire casi no se mueve (α* implícito ≈ 1,04). Hargreaves funciona (+6 %) porque usa la amplitud térmica como indicador "
  "de nubosidad. Los métodos de temperatura sola fallan por 40 %: se calibraron en latitudes medias, donde temperatura y radiación "
  "suben juntas en verano, y en los Andes tropicales hace frío con un Sol intenso. Los de transferencia de masa solo ven el aire "
  "seco y tienen el máximo en julio–agosto.")
TAB([["Método", "Familia", "mm/año", "Sesgo", "RMSE (mm/mes)", "r", "Máximo"],
     ["FAO-56 (referencia)", "combinado", "1 208", "—", "—", "—", "Nov"],
     ["Makkink", "radiación", "1 180", "−2 %", "3,9", "0,97", "Nov"],
     ["Hargreaves", "T + Ra", "1 277", "+6 %", "6,8", "0,96", "Oct"],
     ["Linacre", "T + rocío", "1 303", "+8 %", "13,0", "0,65", "Ago"],
     ["Oudin", "T + Ra", "1 007", "−17 %", "19,8", "0,50", "Ene"],
     ["Priestley–Taylor", "radiación", "1 466", "+21 %", "21,9", "0,96", "Nov"],
     ["Meyer", "masa", "873", "−28 %", "33,5", "0,61", "Ago"],
     ["Dalton (FAO-24)", "masa", "840", "−30 %", "36,2", "0,60", "Ago"],
     ["Hamon", "temperatura", "736", "−39 %", "40,8", "0,55", "Dic"],
     ["Thornthwaite", "temperatura", "691", "−43 %", "44,5", "0,48", "Dic"]],
    "Métodos simples frente a FAO-56 (serie mensual de la cuenca, 1981–2025), ordenados por el sesgo absoluto. Hargreaves, Oudin, "
    "Priestley–Taylor y Linacre se verificaron con pyet (±0,4 %).", [0.22, 0.15, 0.11, 0.1, 0.16, 0.1, 0.16])
FIG(IMG / "metodos.webp", "Ciclo anual de los nueve métodos simples frente a FAO-56 y la banda de dispersión de las pruebas de 1.1.b.", ancho=0.9)
P("El experimento de sensibilidad confirma el diagnóstico (Figura 5). En FAO-56, R<sub>s</sub> −10 % baja la ETo 7,3 %; +1 °C la sube "
  "3,0 % con e<sub>a</sub> fija y 2,5 % con humedad relativa fija; el viento casi no importa (+0,3 % por +10 %). Thornthwaite sube 10 % "
  "por grado, tres veces más, porque lee el calentamiento como si llegara más energía aunque la radiación no cambie (Milly y Dunne, "
  "2016). Si solo tuviéramos temperatura usaríamos Hargreaves (Hargreaves y Samani, 1985): sesgo de +6 %, fase correcta y una "
  "respuesta al calentamiento como la de FAO-56.")
FIG(IMG / "sensibilidad.webp", "Respuesta de cada método a perturbaciones de la radiación, la temperatura, la humedad y el viento.", ancho=0.85)
H2("3.3 La evapotranspiración real: cuánta demanda se satisface")
P("De 833 mm de lluvia al año (1998–2024), 391 salen por el río y ≈ 442 ± 158 mm vuelven a la atmósfera: ET/P = 0,53 y ET/ETo = 0,36 "
  "(Figura 6). La incertidumbre la domina la lluvia, no el caudal: con IMERG la ET sería de 148 mm/año, una fracción evaporativa de "
  "0,08 en una cuenca con pasto, arbustal y bosque que no es creíble y confirma que IMERG subestima la lluvia. Año a año, P − Q va de "
  "≈ 400 a ≈ 690 mm y apenas sigue a la lluvia (r = 0,29). Las excepciones son 2000 y 2017, años de mar cálido en que el agua guardada "
  "sale después por el río (ΔS ≠ 0). Todas las estimaciones quedan muy por debajo del techo energético (R<sub>n</sub> = 140 W/m² ≈ 1 806 "
  "mm/año): no hace falta advección para explicar la ET.")
FIG(IMG / "et_balance.webp", "Balance de largo plazo P<sub>L</sub> − Q y por año hidrológico (octubre–septiembre), con la propagación "
    "del error de la lluvia y del caudal.", ancho=0.85)
P("El balance diario del suelo muestra el mecanismo mes a mes (Figura 7). Con la lluvia y la PET de CAMELS-PE, una cubeta de "
  "100 mm reproduce el ciclo del caudal observado (error medio de 5,6 mm/mes) y da una ET de ≈ 500 mm/año, el 37 % de la demanda. "
  "En febrero–marzo la ET llega al 74–82 % de la ETp; de junio a septiembre el suelo se vacía y cae al 3–16 %. El almacenamiento se "
  "llena de diciembre a abril y se vacía de mayo a noviembre, que es el R &gt; P del estiaje que encontró la Tarea 1. La ET simulada "
  "depende más del suelo (S<sub>máx</sub> de 50 a 200 mm: 477 a 538 mm/año) que de la demanda (−9 % en ETp: −4 % en ET), y con la "
  "lluvia de IMERG el modelo solo produce ≈ 130 mm/año de caudal frente a 403 observados.")
FIG(PR1 / "fig_balance_cubeta.png", "Balance diario del suelo agregado a meses (1998–2016): lluvia, demanda (PISCOeo PM), ET real "
    "simulada, caudal simulado y observado, y cambio de almacenamiento ΔS = P − ET − Q. Adaptado de la figura #s14 de Hoyos (2026). "
    "Script codigo/punto1/00_balance_cubeta.py.", ancho=0.82)
P("La ET de los productos hereda su propia lluvia. ERA5-Land da 947 mm/año, más que toda la lluvia observada, porque su reanálisis "
  "hace llover 2,3 veces más y su suelo nunca se seca; TerraClimate (518 mm/año) queda a +17 % del balance por compensación, con "
  "menos lluvia y casi sin caudal. Para el total multianual el balance es la mejor estimación; para el ciclo anual dependemos de "
  "productos o de modelos, y el de TerraClimate (máximo en marzo, 12 mm en agosto) es el más coherente con una cuenca limitada por "
  "agua.")
H2("3.4 ¿Manda la energía o el agua?")
P("Con ET/ETo de TerraClimate, febrero y marzo están limitados por energía (≥ 0,75), enero, abril, octubre y diciembre son de "
  "transición, y de mayo a septiembre y noviembre manda el agua (≤ 0,40; Figura 8). La clasificación coincide con las temporadas de "
  "la Tarea 1. De mayo a septiembre la ET es una a cuatro veces la lluvia: esa agua sale del suelo llenado en febrero–abril, y quizá "
  "de las areniscas fisuradas que sostienen el estiaje. En ERA5-Land la fracción evaporativa se estanca por encima de θ* ≈ 0,33 m³/m³ "
  "y cae casi en línea recta por debajo; julio–septiembre están bajo el umbral en el 96–100 % de los años. La energía que no se "
  "evapora calienta el aire (la razón de Bowen del balance es ≈ 3), y la capa límite crece hasta chocar con el aire estable de la "
  "inversión.")
FIG(IMG / "regimen.webp", "Régimen evaporativo mensual (ET/ETo) y fracción evaporativa frente a la humedad del suelo de ERA5-Land.", ancho=0.85)
P("En el promedio multianual manda el agua: φ = ETo/P = 1,47. Pero la cuenca evapora mucho menos de lo que predice Budyko (ET/P = "
  "0,53 frente a 0,82): ω de Fu = 1,57 con todos los años y 1,77 sin 2000 ni 2017, frente a ≈ 2,6 en promedio mundial (Figura 9). Un ω "
  "bajo significa que la lluvia se va por el río antes de evaporarse: el 81 % cae en diciembre–abril, cuando la demanda es menor por "
  "las nubes; los suelos se saturan, las laderas de 27° escurren rápido, y cuando la demanda es máxima ya no hay agua. La explicación "
  "alternativa, que P<sub>L</sub> subestime la lluvia de las cabeceras, haría ET/P mayor y no se puede descartar. La elasticidad del "
  "caudal a la lluvia es 1,46 (1,57 en CAMELS-PE): 10 % menos lluvia, ≈ 15 % menos caudal.")
FIG(IMG / "budyko.webp", "Espacio de Budyko con la curva de Fu y la relación complementaria entre ETo y ET real.", ancho=0.85)
P("La relación complementaria se observa: en los años secos la demanda es mayor (1 240 frente a 1 216 mm/a; r(ETo, P) = −0,51) y la "
  "ET menor (438 frente a 512 mm/a), y en la estación seca las anomalías mensuales de ETo y ET van en sentidos opuestos (r = −0,58). "
  "El mecanismo propuesto es atmosférico (Bouchet, 1963): con el suelo seco la superficie entrega más calor sensible, la capa límite "
  "se calienta, crece y mezcla aire seco de arriba. Una alternativa que estos datos no separan es que un año seco también tiene menos "
  "nubes y más radiación.")
H2("3.5 Cambios en la demanda y paisaje de la cuenca")
P("La ETo sube 9,2 mm/año por década en 1981–2025 (Sen; IC 95 % 3,2 a 17,3; p = 0,01), +0,8 % por década; solo cuatro meses de la "
  "estación seca pasan el control FDR. Thornthwaite sube casi tres veces más en términos relativos, como predijo la sensibilidad. La "
  "atribución reproduce la tendencia (residuo de 0,2 mm por década): la temperatura (+0,20 °C/década) explica ≈ ¾ y la radiación el "
  "resto (Figura 10). Desde 2000 la temperatura sigue subiendo pero la radiación baja y la ETo cae (−4,2 mm por década): una versión "
  "local de la paradoja del tanque evaporímetro (Roderick y Farquhar, 2002), aquí por radiación y no por el viento. La caída de la "
  "radiación puede deberse a más nubes, a aerosoles o a cambios del reanálisis; no se pueden separar. En el bono, el caudal casi no "
  "cambió entre 1999–2011 y 2012–2024 (−2 mm/año), y la mayor demanda aportó menos de 1 mm/año.")
FIG(IMG / "atribucion.webp", "Atribución de la tendencia de ETo a cada forzante en 1981–2025 y 2000–2025.", ancho=0.85)
P("Con la altura la demanda baja ≈ 110 mm por km (aire más frío) pero la lluvia sube ≈ 210 mm por km, así que la ET real aumenta de "
  "241 mm/año en el fondo del valle a 785 en las cumbres y la limitación por agua se relaja (ET/PET de 0,18 a 0,74; Figura 11). Las dos "
  "rectas se cruzan cerca de la divisoria: el máximo intermedio de ET que se esperaría en una montaña más alta no aparece dentro de la "
  "cuenca. La niebla costera no alcanza a cerrar la brecha entre la ET de los productos y el balance (solo el 4,3 % de la cuenca queda "
  "en la franja de la inversión); las nubes de ladera de la temporada húmeda podrían aportar algo, pero verificarlo exigiría "
  "colectores.")
FIG(IMG / "gradiente.webp", "Gradiente altitudinal de la demanda, la lluvia y la ET real por bandas de 500 m (TerraClimate y ERA5-Land "
    "corregido).", ancho=0.85)

H2("3.6 Lluvia de 10 minutos en Cottbus: datos, extremos e IDF")
P("La serie de Cottbus tiene tres épocas de instrumento (contador de gotas hasta 2007, pluviómetros de pesada después). El control de "
  "calidad eliminó 22 días de 1991–1994 con más lluvia en 10 min que en el pluviómetro diario, marcó 2021 por un aviso de firmware y "
  "conservó como incertidumbre el 3 de agosto de 2014 (48,6 mm en 10 min frente a 102,6 mm diarios). Quedan 30 años válidos. Ninguna "
  "duración tiene tendencia significativa (Mann–Kendall–Hamed–Rao, p ≥ 0,15) ni diferencias entre épocas de instrumento "
  "(Kruskal–Wallis, p ≥ 0,50), aunque todas las pendientes son positivas (+2 a +8 % por década).")
P("Hasta T ≈ 10 años las nueve combinaciones de distribución y método coinciden; después se abren (13 % en el cuantil de 100 años de "
  "60 min y hasta 40 % en 180 min; Figura 12). Se eligió la GEV por L-momentos por ser la distribución límite de los máximos de bloque, "
  "tener el mejor ajuste promedio (KS 0,100, AD 0,29, PPCC 0,984) y una estimación estable con 30 datos; la máxima verosimilitud dio una "
  "forma inestable. La forma de la cola tiene sentido físico: ξ = −0,13 en 10 min (una celda convectiva tiene un techo de lluvia), ≈ 0 "
  "en 20 min–1 h y +0,2 en 2–3 h (tormentas organizadas). El IC 95 % del bootstrap va del 31 % del valor en T = 2 al 85 % en T = 100: "
  "más allá de T ≈ 2N = 60 años, el resultado depende casi por completo de la cola supuesta.")
FIG(IMG / "ajustes.webp", "Máximos anuales de Cottbus con posiciones de Gringorten y ajustes de Gumbel, GEV y log-Pearson III por tres "
    "métodos. Script codigo/punto2 (versión de Marlin Schwebel).", ancho=0.9)
P("La IDF es i(d,T) = a(T)/(d + 0,137)<sup>0,872</sup> [mm/h, d en h], con a(T) GEV (μ = 19,42, σ = 7,52, ξ = 0,044), con 5,7 % de "
  "error medio frente a los cuantiles de cada duración (Figura 13). Queda entre 0,80 y 1,11 veces la IDF oficial KOSTRA-DWD-2020 y "
  "dentro de su incertidumbre (±18–27 %). La tormenta del 9 de junio de 2018 (51,6 mm en 1 h y casi nada a 50 km) tiene T ≈ 110–190 "
  "años en 10–60 min pero solo ≈ 7 años en 24 h: es la firma de una celda convectiva, que descarga en minutos sobre pocos km².")
FIG(IMG / "idf.webp", "Curvas IDF de Cottbus (GEV por L-momentos) con las bandas del bootstrap para T = 10 y 100 años y la IDF "
    "oficial KOSTRA-DWD-2020.", ancho=0.85)

H2("3.7 Problema A: la alcantarilla y su riesgo efectivo")
P("La microcuenca es una zanja de prados de 0,44 km² a 7,2 km del pluviómetro, en la misma llanura y sin relieve entre ellos (83 % "
  "pastizal, 17 % bosque, 2,1 m de desnivel). Con una vida útil de 30 años y un riesgo aceptado del 70 % (consecuencias bajas: un camino "
  "rural y prados) el periodo de retorno es 25 años, dentro de la práctica de Chow et al. (1988) para vías secundarias. Con Kirpich "
  "(t<sub>c</sub> = 102 min), C = 0,335 y FRA = 0,976: Q<sub>p</sub> = 0,335 · 26,6 · 0,437 · 0,976 / 3,6 = 1,05 m³/s. Manning a tubo "
  "lleno con S₀ = 0,5 % da D = 0,84 m: <b>DN 1 000</b> (capacidad 1,70 m³/s).")
P("La incertidumbre que más pesa no es la lluvia sino la fórmula del tiempo de concentración (Tabla 3): en 1,4 m de desnivel en 1,85 km "
  "cada fórmula imagina distinto el recorrido del agua. El tiempo largo hace que la obra «vea» tormentas convectivas organizadas "
  "(≈ 27 mm/h durante 1,7 h), no los chubascos de 10–20 min, más intensos pero cortos. La distribución casi no importa: a T = 25 todas "
  "interpolan.")
TAB([["Fuente de incertidumbre", "D calculado", "Rango"],
     ["t<sub>c</sub> (Kirpich, Témez, velocidad, SCS)", "0,63–0,87 m", "29 %"],
     ["IC 95 % de la IDF (30 años)", "0,71–0,91 m", "25 %"],
     ["C ± 20 %", "0,77–0,90 m", "15 %"],
     ["Distribución (GEV, Gumbel, LP3, GLO)", "0,82–0,83 m", "1 %"]],
    "Fuentes de incertidumbre del diámetro de la alcantarilla, ordenadas por su efecto (diseño D = 0,84 m → DN 1 000).", [0.5, 0.28, 0.22])
FIG(PR2 / "fig_obra_alcantarilla.png", "Geometría de la obra con el caudal de diseño: perfil longitudinal a escala (tubo DN 1 000 de "
    "12 m, S₀ = 0,5 %, terraplén del camino) y sección transversal (y<sub>n</sub> = 0,57 m, y<sub>c</sub> = 0,59 m, V = 2,27 m/s). La "
    "longitud es la mediana de las 87 alcantarillas cartografiadas en OpenStreetMap a menos de 1,5 km. Script codigo/punto2/10_geometria_obras.py.")
P("La hidráulica a sección parcial (Figura 14) muestra que con el caudal de diseño el tubo va a 57 % de su altura, a 2,3 m/s, y con un flujo "
  "casi crítico (Fr = 1,06): la pendiente crítica es 0,45 % y la del tubo 0,5 %. Cerca de Fr = 1 la superficie es ondulada e inestable, "
  "así que conviene S₀ = 0,4 % (flujo subcrítico, capacidad 1,52 m³/s) y un enrocado a la salida. Falta verificar el control a la "
  "entrada (carga aguas arriba; FHWA, 2012). La ortofoto del sitio muestra además que el cierre queda junto a un estanque con bosque: "
  "si está dentro de la microcuenca, su almacenamiento amortigua el pico y el método racional lo sobrestima.")
P("El riesgo efectivo es muy distinto del nominal (Figura 15). El Monte Carlo, que combina las réplicas bootstrap de la IDF, t<sub>c</sub> "
  "triangular, C ± 20 % y rangos de n y S₀, da una probabilidad anual de falla de 0,0036 (T efectivo ≈ 275 años) frente a 0,04 "
  "nominal: el riesgo en 30 años baja de 71 % a 10 %. No es que la incertidumbre ayude; es el salto al diámetro comercial, que con el "
  "valor central ya alcanza T ≈ 100 años. Se debe reportar el riesgo del tubo instalado, no el del cálculo.")
FIG(IMG / "problemaA_mc.webp", "Monte Carlo de la probabilidad anual de falla del DN 1 000 (réplicas bootstrap de la IDF, t<sub>c</sub>, "
    "C, n y S₀).", ancho=0.8)
P("<b>Escala.</b> El método racional supone lluvia uniforme durante t<sub>c</sub>: vale en 0,44 km² (FRA ≈ 0,98) y no en el Chicama de "
  "2 170 km², donde t<sub>c</sub> ≈ 6,4 h y la lluvia media de una hora sobre la cuenca es 33 % menor que la del pluviómetro. Allí, "
  "además, el caudal depende de la humedad previa, del almacenamiento (sección 3.3) y de que la lluvia caiga solo en la parte alta.")

H2("3.8 Problema B: urbanizar y detener")
P("Con 50 % de superficie impermeable, C pasa de 0,34 a 0,58 y t<sub>c</sub> de 102 a 41 min (Kirpich × 0,4, comprobado con cunetas a "
  "1 m/s). Los dos efectos se multiplican: con un tiempo más corto la obra ve lluvias más cortas, más intensas, y el pico sube de 1,08 a "
  "3,79 m³/s (3,5 veces). El tanque de detención, con orificio libre dimensionado para dejar salir el pico previo a tanque lleno "
  "(A<sub>o</sub> = 0,332 m², Ø ≈ 0,65 m) y 1,5 m de lámina por el nivel freático somero, necesita ≈ 6 400 m³ con la duración crítica de "
  "los trapecios (60 min) y ≈ 7 900 m³ con bloques alternos de 6 h con el pico al centro, 24 % más porque ese hietograma contiene la "
  "lluvia de 25 años de todas las duraciones (Figura 16).")
FIG(PR2 / "fig_obra_tanque.png", "Corte del tanque de detención (72,6 × 72,6 m, h<sub>máx</sub> = 1,5 m): entrada desde la red urbana, "
    "orificio de 0,65 m (O = C<sub>d</sub>A<sub>o</sub>√(2gh) ≤ 1,08 m³/s) y vertedero de excedencias propuesto. Escala vertical "
    "exagerada. Script codigo/punto2/10_geometria_obras.py.", ancho=0.85)
P("La posición del pico importa: al inicio, el pico llega con el tanque vacío (6 500 m³); al centro o al final, 2–3 h de lluvia previa "
  "ya ocuparon parte del volumen (7 700–7 900 m³). Las tormentas convectivas de Cottbus suelen parecerse al patrón de pico temprano "
  "(Huff, 1967), pero se diseñó con el más exigente. C es lo que más mueve el volumen (±20 % en C da −28 % y +29 %); el t<sub>c</sub> "
  "urbano, 5–10 %, y la duración del hietograma casi no importa más allá de 3 h. Se recomienda V ≈ 7 900 m³, A<sub>s</sub> ≈ 5 200 m² y "
  "un vertedero de excedencias (≈ 14 m de cresta para pasar 3,79 m³/s con 0,3 m de carga, por verificar) para que una tormenta mayor "
  "salga por un sitio controlado.")
P("<b>Costo esperado (bono 2.5.1).</b> Con las funciones de la guía, el mínimo del costo anual esperado está en T ≈ 500 años, pero la "
  "curva es muy plana (entre 200 y 1 000 años está a menos de 5 % del mínimo) y cae en zona de extrapolación. La conclusión robusta es "
  "que T = 25 queda corto en costo, y que el DN 1 000, que ya cubre T ≈ 100, es una buena elección.")

H2("3.9 La misma obra en Medellín: quebrada La Presidenta")
P("La Presidenta drena 13,38 km² de la ladera oriental del Valle de Aburrá (Corantioquia, 2023), con 145 estructuras hidráulicas "
  "(canales, box culverts y puentes) y crecidas súbitas frecuentes. Con el DEM GLO-30 la cuenca a la estación de nivel 101 tiene "
  "10,5 km², un cauce de 5,1 km y una pendiente de 25 %. En 427 crecidas con subida ≥ 40 cm (2013–2026) el nivel sube en 27 min "
  "(mediana; 17–57 min el 50 % central) y en 25 min en las 13 mayores de 1,5 m; el pico llega 23 min después del centroide de la "
  "lluvia (Figura 17). De las fórmulas de t<sub>c</sub>, solo Kirpich (24 min) se acerca; Témez, Bransby-Williams y Giandotti dan "
  "65–81 min. Es la comprobación que no se pudo hacer en Cottbus.")
FIG(PR2 / "fig_presidenta_respuesta.png", "Quebrada La Presidenta (SIATA): (a) la crecida mayor del registro, nivel de la estación 101 "
    "y lluvia del pluviómetro 41 cada minuto; (b) distribución del tiempo de subida en 427 crecidas frente a cuatro fórmulas de "
    "t<sub>c</sub> (las mayores de 130 min se agrupan en el último intervalo). Script codigo/punto2/14_presidenta_respuesta.py.")
P("La lluvia corta es parecida en los dos sitios: a 10 min, Cottbus, la IDF oficial de EPM y los máximos anuales de SIATA dan ≈ 125–130 "
  "mm/h para 25 años (Figura 18). La diferencia crece con la duración (25–30 % más en Medellín a 1–2 h, 60 % más en la estación Palmas, "
  "a 2 550 m). Para T = 25 la IDF de EPM queda a menos de 10 % de lo que mide SIATA, y en los años comunes SIATA mide 15–20 % más. Lo "
  "que más cambia el diseño es la duración que importa: en la parte alta de La Presidenta (0,55 km², cauce de 47 %) el t<sub>c</sub> de "
  "la FAA que exige EPM es de 10–21 min, y la misma obra tendría que llevar 5,8 m³/s en ladera con vegetación y 15,1 m³/s urbanizada "
  "(10,6 y 27,6 m³/s/km², frente a 2,4 en Cottbus). Con esos caudales la velocidad supera los 5 m/s que permite la norma, y la obra "
  "debe escalonarse o pasar a box culvert.")
FIG(PR2 / "fig_idf_comparacion.png", "Intensidades de 10, 30 y 60 minutos: IDF de Cottbus, IDF oficial de EPM (estaciones Medellín y "
    "Palmas, anexo 6.1, 2025) y máximos anuales del pluviómetro SIATA 41 con posiciones de Gringorten (11 años válidos). Script "
    "codigo/informe/16_figuras_informe.py.")
P("La diferencia no está en el método sino en la hidroclimatología y la geomorfología: en una llanura templada, con una temporada de "
  "tormentas convectivas y t<sub>c</sub> de horas, la obra falla por volumen y remanso; en una quebrada andina tropical bimodal, con "
  "laderas empinadas y t<sub>c</sub> de minutos, falla por energía, obstrucción y arrastre de sedimentos. Por eso tampoco se puede "
  "copiar el periodo de retorno: EPM pide 10 años para áreas de 10–1 000 ha y 100 años para canales que drenan más de 1 000 ha.")

H2("3.10 Deforestación y urbanización: el balance anual y los caudales pico")
P("La presión urbanística sobre las laderas actúa en dos escalas que conviene separar. En la escala anual, quitar bosque reduce la "
  "evapotranspiración (raíces menos profundas, menos interceptación y menos rugosidad) y deja más agua para el caudal (Bosch y "
  "Hewlett, 1982; Zhang et al., 2001). En la escala de la crecida, impermeabilizar y deforestar reduce la infiltración y acorta el "
  "tiempo de concentración, y los dos efectos se multiplican en el caudal pico, como en el problema B.")
P("<b>Balance anual en el Chicama.</b> Con la curva de Fu y el parámetro ω ligado a la cobertura del apartado 1.5 (calibrado para que "
  "pasar de pastizal a bosque nativo aumente la ET 19 %, como en las curvas de Zhang et al., 2001), quitar el 15 % de bosque actual sube "
  "el caudal solo de 387 a 400 mm/año (+3 %; Figura 19a). El efecto es pequeño porque la cuenca está limitada por agua: hay poca agua "
  "«sobrante» que el bosque pueda evaporar. En una cuenca húmeda la misma deforestación pesaría más.")
P("<b>Balance anual en La Presidenta.</b> El pluviómetro SIATA 41 registra 1 850 mm/año en promedio (2015–2025; de 1 125 mm en 2015 a "
  "2 909 mm en 2022, año de La Niña), más del doble que el Chicama. Con el caudal medio que simula Corantioquia (2023) para la salida "
  "(333 L/s, 785 mm/año), Q/P ≈ 0,42 y P − Q ≈ 1 070 mm/año: la ET es más del doble que en el Chicama porque sobran agua y energía "
  "(clima húmedo, sin una estación seca larga). Al urbanizar, el agua que caía sobre suelo y vegetación cae sobre techos y vías: "
  "se evapora menos y se infiltra menos, y más lluvia sale como escorrentía rápida. El caudal anual puede aumentar mientras el flujo "
  "base baja, porque disminuye la recarga del suelo. Este caudal anual es simulado (modelo de tanques de Corantioquia), y el "
  "pluviómetro está en la parte baja de una cuenca que sube a 2 800 m: son órdenes de magnitud, no un balance observado.")
FIG(PR2 / "fig_cobertura.png", "Efecto de la cobertura. (a) Chicama: ET y caudal anuales con la curva de Fu y ω según el porcentaje de "
    "bosque nativo o de plantación (modelo del apartado 1.5). (b) Parte alta de La Presidenta: caudal pico con la norma EPM para cuatro "
    "coberturas de su tabla 6-3 (I = 0,30 a 0,90; C = 0,15 + 0,85·I; t<sub>c</sub> de la FAA; IDF Medellín) frente a los picos de "
    "Cottbus. Script codigo/punto2/19_cobertura_caudales.py.")
P("<b>Caudales pico.</b> En la parte alta de La Presidenta, la norma de EPM permite seguir la secuencia de presión urbanística con sus "
  "coeficientes de impermeabilidad (Figura 19b). Pasar de ladera con vegetación (I = 0,30) a ladera sin vegetación (0,60) duplica el "
  "pico para 25 años (5,8 → 11,3 m³/s) antes de construir nada, porque C sube de 0,41 a 0,66 y el t<sub>c</sub> baja de 21 a 14 min. "
  "La urbanización residencial lo lleva a 15,1 m³/s (×2,6) y la comercial a 19,8 m³/s (×3,4), con t<sub>c</sub> de 6 min. En Cottbus "
  "urbanizar la mitad de la microcuenca multiplicó el pico por 3,5: el factor es parecido, pero en Medellín se aplica sobre un caudal "
  "base cinco veces mayor por km². Una obra dimensionada con la cobertura actual queda corta si la ladera se deforesta o se urbaniza "
  "durante su vida útil; por eso la norma de EPM pide estimar C para el final del periodo de diseño con el plan de ordenamiento.")
P("<b>¿Se ve en los datos?</b> En las 427 crecidas de La Presidenta no hay tendencia detectable entre 2014 y 2026 en el tiempo de subida, "
  "en el retardo ni en la subida de nivel por milímetro de lluvia (Kendall, p ≥ 0,16). Con 12 años, una cuenca ya muy urbanizada al "
  "comenzar el registro y los años de La Niña de 2020–2023, el efecto de los cambios recientes de cobertura no se puede separar de la "
  "variabilidad del clima. Atribuirlo exigiría mapas de cobertura por año y un registro más largo.")

H2("3.11 Cambio climático y adaptación de la obra")
P("En la IDF, a(T) = μ + σ·y<sub>T</sub>: el término de la escala es el 57 % de la lluvia de 25 años y el 66 % de la de 100 años. Por "
  "eso la variabilidad pesa más que la media: si σ aumenta 25 % con la misma media, la lluvia de diseño sube ≈ 14 %, más que con 1,5 °C "
  "de calentamiento a 7 %/°C (≈ 11 %; Figura 20). Con +1,5 °C la lluvia de 25 años de hoy pasaría a repetirse cada ≈ 15 años. El DN 1 000 "
  "resiste porque D ∝ Q<sup>3/8</sup>: un 10 % más de caudal pide 3–4 % más de diámetro, y el T efectivo del tubo pasa de ≈ 470 a ≈ 230 "
  "años (≈ 180 con +2 °C y ≈ 80 si la tasa fuera de 14 %/°C). Las pendientes positivas de los máximos de Cottbus (+2 a +8 % por década, "
  "no significativas) son compatibles con esa intensificación, pero con 30 años la incertidumbre de muestreo (±25 % en el diámetro) "
  "todavía supera la señal de 1–2 °C.")
FIG(PR2 / "fig_clima_escenarios.png", "Intensidad de duración t<sub>c</sub> = 102 min según el periodo de retorno, hoy y con tres "
    "escenarios de escalado, frente a la capacidad del DN 1 000. Es una sensibilidad, no una proyección. Script "
    "codigo/punto2/11_escenarios_clima.py.", ancho=0.72)
P("Las decisiones de adaptación que se desprenden son cuatro: diseñar con la banda de incertidumbre y no con el valor central; preferir "
  "obras que fallen sin romperse (vertedero de excedencias, terraplén que pueda desbordar); dejar el margen barato del diámetro comercial "
  "siguiente cuando el costo esperado lo justifica; y revisar la IDF cada década con la tendencia que ya muestran los datos.")

H2("3.12 Síntesis física y tabla de evidencias")
TAB([["Resultado principal", "Mecanismo propuesto", "Evidencia propia", "Fuente consultada", "Limitación o alternativa"],
     ["Demanda ≈ 1 210 mm/año, casi constante, 90 % radiativa", "Sol alto todo el año; nubes de la temporada húmeda recortan R<sub>s</sub>",
      "Ciclo de ETo y términos; sensibilidad R<sub>s</sub> −10 % → −7,3 %", "Allen et al. (1998); Hoyos (2026)",
      "ERA5-Land 10–25 % bajo frente a estación y PISCOeo"],
     ["ET ≈ 440 mm/año, 36 % de la demanda", "Limitación por agua bajo la inversión de subsidencia", "Balance P − Q; cubeta diaria con MAE 5,6 mm/mes",
      "Budyko (1974); Thornthwaite y Mather (1955)", "±160 mm por la lluvia; ΔS ≠ 0 en 2000 y 2017"],
     ["Energía en feb–mar, agua en may–sep", "Suelo lleno en la temporada húmeda, vacío en la seca", "ET/ETo mensual; umbral θ* de ERA5-Land",
      "Dingman (2015)", "Ciclo de ET de productos y de modelo"],
     ["ω de Fu bajo (1,6–1,8)", "Lluvia concentrada cuando la demanda es menor; laderas empinadas", "Budyko–Fu con y sin años anómalos",
      "Fu (1981); Zhang et al. (2001)", "Lluvia de cabecera subestimada"],
     ["Métodos de temperatura fallan −40 %", "Calibración de latitudes medias; frío con Sol intenso", "Nueve métodos y su sensibilidad",
      "Thornthwaite (1948); Milly y Dunne (2016)", "Makkink podría acertar por compensación"],
     ["t<sub>c</sub> domina el diámetro", "La duración crítica fija la intensidad", "Rangos de D por fuente de incertidumbre; La Presidenta",
      "Kirpich (1940); Chow et al. (1988)", "t<sub>c</sub> de Cottbus sin validar"],
     ["Riesgo efectivo ≈ 10 % frente a 71 %", "Salto al diámetro comercial", "Monte Carlo con bootstrap", "Chow et al. (1988)",
      "Depende de la cola de la GEV"],
     ["La misma obra en Medellín lleva 4–5 veces más caudal por km²", "t<sub>c</sub> de minutos en laderas de 47 %",
      "427 crecidas de SIATA; IDF de EPM", "EPM (2025); Corantioquia (2023)", "Nivel sin curva de gasto"],
     ["Deforestar la ladera duplica el pico; urbanizar lo multiplica por 2,6–3,4", "Menos infiltración y t<sub>c</sub> más corto", "Coberturas de la tabla 6-3 de EPM; Fu con ω según cobertura", "Bosch y Hewlett (1982); Zhang et al. (2001)", "Sin tendencia detectable en 12 años de SIATA; C de tabla"]],
    "Tabla de evidencias: resultado, mecanismo, evidencia propia, fuente y limitación.", [0.2, 0.22, 0.22, 0.17, 0.19])

# ====================================================================== 4. Conclusiones
H1("4. Conclusiones")
B("<b>La pregunta del punto 1.</b> La atmósfera pide sobre la cuenca casi el triple de lo que la cuenca le devuelve. H0 no se sostiene: "
  "el balance de la Tarea 1 no cierra porque la cuenca está limitada por agua con alternancia estacional, como plantea H1. Evapora al "
  "ritmo de la energía solo en el pico de lluvias y después vive del suelo hasta vaciarlo; la lluvia llega concentrada cuando la "
  "demanda es menor y buena parte escurre. La subestimación de la lluvia de cabecera sigue siendo una explicación alternativa.")
B("<b>Qué controla la diferencia.</b> En el año, la inversión de subsidencia: tapa la convección, seca el suelo y, con cielo despejado y "
  "aire seco, sube la demanda. Entre años la ET real casi no cambia y la variabilidad de la lluvia pasa al caudal; la relación "
  "complementaria se observa.")
B("<b>Sensibilidad.</b> La respuesta depende sobre todo de la fuente de lluvia, después del método de demanda y menos de la resolución "
  "temporal si se usan los extremos diarios. Los métodos radiativos y Hargreaves son adecuados en esta cuenca; los de temperatura "
  "sola no, y exageran la respuesta al calentamiento.")
B("<b>La pregunta del punto 2.</b> El diámetro de una obra pequeña en terreno plano lo controla más el tiempo de concentración que la "
  "lluvia o la distribución, y el riesgo real lo controla el diámetro comercial. Se recomienda una alcantarilla de concreto DN 1 000 "
  "con pendiente de 0,4–0,5 % y un tanque de detención de ≈ 7 900 m³ con orificio de 0,65 m y vertedero de excedencias si se urbaniza; "
  "la incertidumbre del diámetro calculado queda cubierta por el diámetro comercial.")
B("<b>Cobertura.</b> En la escala anual, quitar bosque aumenta poco el caudal de una cuenca limitada por agua como el Chicama. En la "
  "escala de la crecida, deforestar y urbanizar una ladera andina multiplica el pico, porque reduce la infiltración y acorta el tiempo "
  "de concentración a la vez; el diseño debe usar la cobertura esperada al final de la vida útil y no la actual.")
B("<b>Escala, región y clima.</b> El método racional es válido en la microcuenca y no en el Chicama. En una quebrada andina como La "
  "Presidenta la misma obra lleva varias veces más caudal y falla por energía; el tiempo de respuesta medido valida la fórmula de "
  "Kirpich. La intensificación de los extremos reduce el periodo de retorno efectivo de la obra, pero con 30 años de datos la "
  "incertidumbre de muestreo aún domina; la adaptación consiste en diseñar con esa incertidumbre y en que las obras fallen de forma "
  "controlada.")
B("<b>Límites.</b> Lluvia de referencia interpolada, ET mensual de productos, ERA5-Land con un valle suavizado, 30 años de extremos con "
  "cambios de instrumento, coeficientes C de tabla, t<sub>c</sub> sin validar en Cottbus y nivel sin caudal en La Presidenta. Los "
  "resultados mejor sustentados son la ET de largo plazo, la limitación por agua de la estación seca, la dominancia radiativa de la "
  "demanda, la IDF de duraciones cortas y el orden de las fuentes de incertidumbre del diseño.")

# ====================================================================== 5. Uso de IA
H1("5. Uso de inteligencia artificial")
P("Nivel A (abierto) de la política del curso. Herramienta: Claude (Anthropic), mediante Claude Code en VS Code, modelo Claude Opus 5.5, "
  "noviembre de 2026. [El grupo debe declarar también las herramientas usadas en la versión inicial del sitio de Marlin Schwebel.]")
B("<b>Tareas en que intervino:</b> revisión crítica del sitio inicial frente a la Tarea 1; adaptación de la página a la línea gráfica de la "
  "Tarea 1; balance diario del suelo; geometría e hidráulica de las obras (esquemas 2D y 3D, alcantarillas reales de OpenStreetMap, "
  "ortofoto); escenarios de cambio climático y de cobertura; búsqueda de documentos oficiales (Corantioquia, EPM) y análisis de La Presidenta con "
  "datos de SIATA; diapositivas de conceptos y de síntesis; figuras estáticas y un borrador de este informe a partir de los resultados "
  "del grupo. [El grupo debe precisar qué scripts y secciones escribió cada integrante.]")
B("<b>Verificaciones:</b> la ETo de un día se reprodujo a mano y con pyet; el balance diario se programó dos veces (Python y JavaScript "
  "de la página) con resultados idénticos; el orificio (0,6·0,332·√(2g·1,5) = 1,08 m³/s), Manning y los tirantes se comprobaron a mano; "
  "los escenarios climáticos se recalcularon en un script aparte; cada figura nueva se revisó visualmente; las áreas de La Presidenta se "
  "contrastaron con Corantioquia y los coeficientes IDF se copiaron del anexo 6.1 de EPM. [El grupo debe añadir sus propias "
  "comprobaciones.]")
B("<b>Propuestas corregidas o rechazadas:</b> en el proceso se corrigieron, entre otros, una afirmación inicial de que la pendiente de "
  "0,5 % era incoherente con el terreno (el tubo mide solo 12 m: son 6 cm), el periodo de retorno efectivo del DN 1 000 (≈ 470 y no 480 "
  "años), una afirmación demasiado general sobre la coincidencia entre la IDF de EPM y SIATA (vale para T = 25, no para T ≈ 2), un error "
  "de índices en la delimitación que daba una cuenca de 8 celdas y un error de lectura de fechas con segundos que dejaba el 99 % de los "
  "datos de SIATA como faltantes. Se descartó comparar con el río Medellín porque a esa escala el método racional no es válido.")
B("<b>Contribución de los integrantes:</b> [completar por el grupo].")
P("Las referencias de la sección 6 son las que sustentan la discusión. El grupo debe leer cada una y comprobar su DOI y páginas antes "
  "de la entrega; una referencia propuesta por la IA no sustituye esa lectura. Quedan marcadas para verificar: Bergström (1992), "
  "McMillan et al. (2012), el coeficiente de vertedero en Chow (1959) y los valores de C de la tabla 15.1.1 de Chow et al. (1988).", "aviso")

# ====================================================================== 6. Referencias
H1("6. Referencias")
REFS = [
    "Abatzoglou, J. T., Dobrowski, S. Z., Parks, S. A. y Hegewisch, K. C. (2018). TerraClimate, a high-resolution global dataset of monthly climate and climatic water balance from 1958–2015. <i>Scientific Data</i>, 5, 170191. doi:10.1038/sdata.2017.191",
    "Allen, R. G., Pereira, L. S., Raes, D. y Smith, M. (1998). <i>Crop evapotranspiration: guidelines for computing crop water requirements</i>. FAO Irrigation and Drainage Paper 56. FAO, Roma.",
    "Aybar, C., Fernández, C., Huerta, A., Lavado, W., Vega, F. y Felipe-Obando, O. (2020). Construction of a high-resolution gridded rainfall dataset for Peru from 1981 to the present day. <i>Hydrological Sciences Journal</i>, 65(5), 770–785. doi:10.1080/02626667.2019.1649411",
    "Bergström, S. (1992). <i>The HBV model: its structure and applications</i>. SMHI Reports Hydrology 4, Norrköping. [verificar]",
    "Bosch, J. M. y Hewlett, J. D. (1982). A review of catchment experiments to determine the effect of vegetation changes on water yield and evapotranspiration. <i>Journal of Hydrology</i>, 55, 3–23. doi:10.1016/0022-1694(82)90117-2",
    "Bouchet, R. J. (1963). Évapotranspiration réelle et potentielle, signification climatique. <i>IAHS Publication</i>, 62, 134–142.",
    "Budyko, M. I. (1974). <i>Climate and Life</i>. Academic Press, Nueva York.",
    "CAMELS-PE v1.0.1 (2026). Conjunto de datos de cuencas del Perú. Zenodo. doi:10.5281/zenodo.21195425",
    "Chow, V. T. (1959). <i>Open-Channel Hydraulics</i>. McGraw-Hill, Nueva York.",
    "Chow, V. T., Maidment, D. R. y Mays, L. W. (1988). <i>Applied Hydrology</i>. McGraw-Hill, Nueva York.",
    "Coles, S. (2001). <i>An Introduction to Statistical Modeling of Extreme Values</i>. Springer, Londres. doi:10.1007/978-1-4471-3675-0",
    "Corantioquia (2023). <i>Reglamentación del uso de las aguas de la quebrada La Presidenta</i>. Contrato CD 387 de 2022, Área Metropolitana del Valle de Aburrá y Corantioquia, Medellín.",
    "Dingman, S. L. (2015). <i>Physical Hydrology</i>, 3.ª ed. Waveland Press, Long Grove.",
    "EPM (2025). <i>GD-AS002 Normas de diseño de sistemas de alcantarillado</i>, versión 1, revisión 2 (actualización del anexo 6.1). Empresas Públicas de Medellín.",
    "FHWA (2012). <i>Hydraulic Design of Highway Culverts</i>, 3.ª ed. Hydraulic Design Series 5, FHWA-HIF-12-026. Federal Highway Administration, Washington.",
    "Fu, B. P. (1981). On the calculation of the evaporation from land surface. <i>Scientia Atmospherica Sinica</i>, 5(1), 23–31 (en chino).",
    "Garreaud, R. D. (2009). The Andes climate and weather. <i>Advances in Geosciences</i>, 22, 3–11. doi:10.5194/adgeo-22-3-2009",
    "Gringorten, I. I. (1963). A plotting rule for extreme probability paper. <i>Journal of Geophysical Research</i>, 68(3), 813–814. doi:10.1029/JZ068i003p00813",
    "Hamed, K. H. y Rao, A. R. (1998). A modified Mann–Kendall trend test for autocorrelated data. <i>Journal of Hydrology</i>, 204, 182–196. doi:10.1016/S0022-1694(97)00125-X",
    "Hargreaves, G. H. y Samani, Z. A. (1985). Reference crop evapotranspiration from temperature. <i>Applied Engineering in Agriculture</i>, 1(2), 96–99. doi:10.13031/2013.26773",
    "Hosking, J. R. M. y Wallis, J. R. (1997). <i>Regional Frequency Analysis: An Approach Based on L-Moments</i>. Cambridge University Press.",
    "Hoyos, C. D. (2026). <i>Evapotranspiración: el agua que vuelve al cielo</i>. Notas del curso de Hidrología, Universidad Nacional de Colombia. https://carlosdhoyos.github.io/cursos/hidrologia/evapotranspiracion.html",
    "Huff, F. A. (1967). Time distribution of rainfall in heavy storms. <i>Water Resources Research</i>, 3(4), 1007–1019. doi:10.1029/WR003i004p01007",
    "Kirpich, Z. P. (1940). Time of concentration of small agricultural watersheds. <i>Civil Engineering</i>, 10(6), 362.",
    "Klein, S. A. y Hartmann, D. L. (1993). The seasonal cycle of low stratiform clouds. <i>Journal of Climate</i>, 6, 1587–1606. doi:10.1175/1520-0442(1993)006<1587:TSCOLS>2.0.CO;2",
    "Koutsoyiannis, D., Kozonis, D. y Manetas, A. (1998). A mathematical framework for studying rainfall intensity–duration–frequency relationships. <i>Journal of Hydrology</i>, 206, 118–135. doi:10.1016/S0022-1694(98)00097-3",
    "Linacre, E. T. (1977). A simple formula for estimating evaporation rates in various climates, using temperature data alone. <i>Agricultural Meteorology</i>, 18, 409–424. doi:10.1016/0002-1571(77)90007-3",
    "Lu, J., Sun, G., McNulty, S. G. y Amatya, D. M. (2005). A comparison of six potential evapotranspiration methods for regional use in the southeastern United States. <i>Journal of the American Water Resources Association</i>, 41(3), 621–633. doi:10.1111/j.1752-1688.2005.tb03759.x",
    "Makkink, G. F. (1957). Testing the Penman formula by means of lysimeters. <i>Journal of the Institution of Water Engineers</i>, 11, 277–288.",
    "McMillan, H., Krueger, T. y Freer, J. (2012). Benchmarking observational uncertainties for hydrology: rainfall, river discharge and water quality. <i>Hydrological Processes</i>, 26, 4078–4111. doi:10.1002/hyp.9384 [verificar]",
    "Milly, P. C. D. y Dunne, K. A. (2016). Potential evapotranspiration and continental drying. <i>Nature Climate Change</i>, 6, 946–949. doi:10.1038/nclimate3046",
    "Muñoz-Sabater, J. et al. (2021). ERA5-Land: a state-of-the-art global reanalysis dataset for land applications. <i>Earth System Science Data</i>, 13, 4349–4383. doi:10.5194/essd-13-4349-2021",
    "Oudin, L., Hervieu, F., Michel, C., Perrin, C., Andréassian, V., Anctil, F. y Loumagne, C. (2005). Which potential evapotranspiration input for a lumped rainfall–runoff model? <i>Journal of Hydrology</i>, 303, 290–306. doi:10.1016/j.jhydrol.2004.08.026",
    "Priestley, C. H. B. y Taylor, R. J. (1972). On the assessment of surface heat flux and evaporation using large-scale parameters. <i>Monthly Weather Review</i>, 100(2), 81–92. doi:10.1175/1520-0493(1972)100<0081:OTAOSH>2.3.CO;2",
    "Roderick, M. L. y Farquhar, G. D. (2002). The cause of decreased pan evaporation over the past 50 years. <i>Science</i>, 298, 1410–1411. doi:10.1126/science.1075390",
    "Sankarasubramanian, A., Vogel, R. M. y Limbrunner, J. F. (2001). Climate elasticity of streamflow in the United States. <i>Water Resources Research</i>, 37(6), 1771–1781. doi:10.1029/2000WR900330",
    "SIATA (2026). Históricos de las estaciones 101 (nivel, Q. La Presidenta – Parque Lineal; doi:10.83041/JEO98W) y 41 (pluviométrica, Comisaría El Poblado; doi:10.83041/8E02OC). Sistema de Alerta Temprana del Valle de Aburrá. https://datos.siata.gov.co",
    "Témez, J. R. (1978). <i>Cálculo hidrometeorológico de caudales máximos en pequeñas cuencas naturales</i>. Dirección General de Carreteras, MOPU, Madrid.",
    "Thornthwaite, C. W. (1948). An approach toward a rational classification of climate. <i>Geographical Review</i>, 38(1), 55–94. doi:10.2307/210739",
    "Thornthwaite, C. W. y Mather, J. R. (1955). The water balance. <i>Publications in Climatology</i>, 8(1). Drexel Institute of Technology, Centerton.",
    "Weiss, L. L. (1964). Ratio of true to fixed-interval maximum rainfall. <i>Journal of the Hydraulics Division, ASCE</i>, 90(HY1), 77–82.",
    "Westra, S. et al. (2014). Future changes to the intensity and frequency of short-duration extreme rainfall. <i>Reviews of Geophysics</i>, 52, 522–555. doi:10.1002/2014RG000464",
    "Zhang, L., Dawes, W. R. y Walker, G. R. (2001). Response of mean annual evapotranspiration to vegetation changes at catchment scale. <i>Water Resources Research</i>, 37(3), 701–708. doi:10.1029/2000WR900325",
]
for r_ in REFS:
    P(r_, "ref")

doc = SimpleDocTemplate(str(SALIDA), pagesize=letter, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=1.8 * cm,
                        bottomMargin=1.8 * cm, title=TITULO, author="Grupo Tarea 2 Hidrología 2026-2")
doc.build(story, onFirstPage=pie, onLaterPages=pie)
print(f"{SALIDA} · {N_FIG[0]} figuras · {N_TAB[0]} tablas")

_pie = W.sections[0].footer.paragraphs[0]
runs(_pie, PIE + " · página ", tam=8, color=GRIS_W)
_campo_pagina(_pie)
W.core_properties.title = TITULO
W.core_properties.author = "Grupo Tarea 2 Hidrología 2026-2"
W.save(str(SALIDA_DOCX))
print(SALIDA_DOCX)
