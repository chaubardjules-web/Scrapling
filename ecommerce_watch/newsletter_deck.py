"""Génère la newsletter Veille E-commerce en PowerPoint, toujours sur la même charte.

La charte (couleurs, typos Aptos, rayons, ombres) est dans charte/charte.yaml.
Structure : une page de garde avec sommaire, puis une slide par news.
Moteur graphique et contrôle de mise en page : pptx-mcp-server
(https://github.com/knorq-ai/pptx-mcp-server).

Usage :
  python ecommerce_watch/newsletter_deck.py contenu.json sortie.pptx [--assets DOSSIER]

--assets pointe vers la banque d'images (branche newsletter-assets du repo :
index.json + images/). Les crédits photo sont repris automatiquement de index.json.

Format de contenu.json :
{
  "date": "5 nov. 2026",
  "cover_image": "images/livraison-1.jpg",          // optionnel, photo de la page de garde
  "news": [                                         // 5 à 8 news
    {
      "theme": "Agentic commerce",                  // surtitre
      "headline": "Titre d'action, 70 caractères max",
      "kpis": [{"value": "+45 %", "label": "ce que mesure le chiffre, 70 car. max"}],  // 1 ou 2
      "points": ["info clé 1", "info clé 2", "info clé 3"],   // 2 à 3 puces, 120 car. max
      "why": "Pourquoi ça compte, 1 à 2 phrases",
      "visual": {"type": "image", "file": "images/doordash-1.jpg", "focus": 0.5}
             // focus : cadrage horizontal quand la photo est rognée (0 = gauche, 1 = droite)
             // ou {"type": "chart", "chart_type": "column" | "bar" | "line" | "doughnut",
             //      "title": "...", "categories": [...], "series": [{"name": "...", "values": [...]}],
             //      "number_format": "0"}
             // ou absent : slide texte seule
      "source": "Modern Retail",
      "url": "https://..."
    }
  ]
}
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from urllib.parse import urlparse

import yaml
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from pptx_mcp_server.engine.charts import _add_chart
from pptx_mcp_server.engine.validation import check_deck_extended
from pptx_mcp_server.theme import DELOITTE, Theme, register_theme

ROOT = Path(__file__).resolve().parent
CHARTE = yaml.safe_load((ROOT / "charte" / "charte.yaml").read_text(encoding="utf-8"))
C = CHARTE["couleurs"]
F = CHARTE["typos"]
W, H = CHARTE["format"]["width"], CHARTE["format"]["height"]
M = CHARTE["marges"]
RADIUS = CHARTE["formes"]["rayon"]
SHADOW = CHARTE["formes"]["ombre"]

CHART_THEME = Theme(
    name="veille_ecommerce",
    colors={**DELOITTE.colors, "primary": f"#{C['encre']}", "accent": f"#{C['vert']}",
            "text_secondary": f"#{C['gris']}", "border": f"#{C['filet']}"},
    fonts={"title": F["texte"], "body": F["texte"]},
    sizes=dict(DELOITTE.sizes), slide={"width": W, "height": H},
    margins=dict(DELOITTE.margins), layout=dict(DELOITTE.layout), table=dict(DELOITTE.table),
    chart_colors=[f"#{c}" for c in C["graphe"]],
)
register_theme(CHART_THEME)


# ------------------------------------------------------------------ typographie

def nb(value: str) -> str:
    """Espaces insécables : 20 %, 500 M€, 10 000, 48 h, « … »."""
    value = re.sub(r"(\d) (?=(%|M€|Md€|M\$|Md\$|Md|T\$|k€|€|\$|h\b|\d{3}\b))", "\\1\u00a0", value or "")
    return value.replace("« ", "«\u00a0").replace(" »", "\u00a0»")


def write(slide, value, left, top, width, height, size, *, font=None, color=None, bold=False,
          italic=False, spacing=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=1.0,
          link=None, shrink=True):
    """Zone de texte simple, sans marge interne, réduite à la taille du cadre si besoin."""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    if shrink:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, chunk in enumerate(str(value).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line
        run = p.add_run()
        run.text = nb(chunk)
        f = run.font
        f.name, f.size, f.bold, f.italic = font or F["texte"], Pt(size), bold, italic
        f.color.rgb = RGBColor.from_string(color or C["encre"])
        if spacing:
            f._rPr.set("spc", str(int(spacing * 100)))
        if link:
            run.hyperlink.address = link
    return box


def fit_size(value: str, base: float, width_in: float, lines: int, floor: float) -> float:
    """Taille de police qui fait tenir le texte sur `lines` lignes (estimation Aptos)."""
    size = base
    while size > floor:
        chars_per_line = width_in * 72 / (size * 0.56)
        if len(value) / chars_per_line <= lines * 0.92:
            break
        size -= 1
    return size


def kicker(slide, value, left, top, width, color=None):
    write(slide, value.upper(), left, top, width, 0.3, 11, font=F["etiquette"], bold=True,
          color=color or C["vert_fonce"], spacing=1.2, shrink=False)


# ------------------------------------------------------------------ formes

def _shadow(shape):
    spPr = shape._element.spPr
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    eff = spPr.makeelement(qn("a:effectLst"), {})
    sh = eff.makeelement(qn("a:outerShdw"), {
        "blurRad": str(int(Pt(SHADOW["flou_pt"]))), "dist": str(int(Pt(SHADOW["distance_pt"]))),
        "dir": "5400000", "algn": "t", "rotWithShape": "0"})
    clr = sh.makeelement(qn("a:srgbClr"), {"val": "000000"})
    alpha = clr.makeelement(qn("a:alpha"), {"val": str(int(SHADOW["opacite"] * 100000))})
    clr.append(alpha)
    sh.append(clr)
    eff.append(sh)
    spPr.append(eff)


def card(slide, left, top, width, height, fill=None, shadow=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top),
                                 Inches(width), Inches(height))
    shp.adjustments[0] = min(0.5, RADIUS / min(width, height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor.from_string(fill or C["fond"])
    shp.line.fill.background()
    if shadow:
        _shadow(shp)
    else:
        shp.shadow.inherit = False
    return shp


def rule(slide, left, top, width, height, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor.from_string(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def hairline(slide, left, top, width):
    ln = slide.shapes.add_connector(1, Inches(left), Inches(top), Inches(left + width), Inches(top))
    ln.line.color.rgb = RGBColor.from_string(C["filet"])
    ln.line.width = Pt(0.75)
    style = ln._element.find(qn("p:style"))
    if style is not None:
        ln._element.remove(style)
    return ln


def photo(slide, path, left, top, width, height, focus=0.5):
    """Photo recadrée pour remplir le cadre, coins arrondis, ombre douce.
    focus : position horizontale à garder quand on rogne (0 = gauche, 1 = droite)."""
    with Image.open(path) as im:
        iw, ih = im.size
    pic = slide.shapes.add_picture(str(path), Inches(left), Inches(top), Inches(width), Inches(height))
    box_ratio, img_ratio = width / height, iw / ih
    if img_ratio > box_ratio:          # trop large : on rogne les côtés
        crop = 1 - box_ratio / img_ratio
        focus = min(max(focus, 0.0), 1.0)
        pic.crop_left, pic.crop_right = crop * focus, crop * (1 - focus)
    else:                              # trop haute : on rogne en haut et en bas (un peu plus en bas)
        crop = 1 - img_ratio / box_ratio
        pic.crop_top, pic.crop_bottom = crop * 0.4, crop * 0.6
    geom = pic._element.spPr.find(qn("a:prstGeom"))
    geom.set("prst", "roundRect")
    av = geom.find(qn("a:avLst"))
    if av is None:
        av = geom.makeelement(qn("a:avLst"), {})
        geom.append(av)
    gd = av.makeelement(qn("a:gd"), {"name": "adj", "fmla": f"val {int(min(50000, RADIUS / min(width, height) * 100000))}"})
    av.append(gd)
    _shadow(pic)
    return pic


def domain(url):
    try:
        return urlparse(url).netloc.removeprefix("www.")
    except Exception:
        return ""


# ------------------------------------------------------------------ assets

class Assets:
    def __init__(self, folder: str | None):
        self.folder = Path(folder) if folder else None
        self.credits = {}
        if self.folder and (self.folder / "index.json").exists():
            for img in json.loads((self.folder / "index.json").read_text(encoding="utf-8")).get("images", []):
                self.credits[img["file"]] = img.get("credit", "")

    def path(self, file: str | None) -> Path | None:
        if not file:
            return None
        p = Path(file)
        if not p.is_absolute() and self.folder:
            p = self.folder / p
        return p if p.exists() else None

    def credit(self, file: str) -> str:
        return self.credits.get(file, "")


# ------------------------------------------------------------------ slides

def footer(slide, news, page):
    src = news.get("source", "")
    if src:
        label = f"Source : {src}" + (f" · {domain(news['url'])}" if news.get("url") else "")
        write(slide, label, M["gauche"], H - 0.42, 7.0, 0.22, 9, color=C["gris_clair"],
              link=news.get("url"), shrink=False)
    write(slide, str(page), W - M["droite"] - 0.5, H - 0.42, 0.5, 0.22, 9, color=C["gris_clair"],
          align=PP_ALIGN.RIGHT, shrink=False)


def cover(prs, data, assets):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    x = M["gauche"]
    img = assets.path(data.get("cover_image"))
    text_w = 6.3 if img else W - x - M["droite"]

    kicker(s, "Deloitte Conseil · Technology & Transformation", x, 0.75, text_w, color=C["gris"])
    write(s, "Veille", x, 1.25, text_w, 0.95, 60, font=F["titre"], bold=True, shrink=False)
    write(s, "e-commerce", x, 2.05, text_w, 0.95, 60, font=F["titre"], bold=True, color=C["vert"],
          shrink=False)
    write(s, f"Édition du {data['date']}", x, 3.1, text_w, 0.45, 20, font=F["citation"], italic=True,
          color=C["gris"], shrink=False)

    news = data.get("news") or []
    top = 4.0
    row = min(0.42, 2.85 / max(len(news), 1))
    for i, n in enumerate(news):
        y = top + i * row
        write(s, f"{i + 2:02d}", x, y, 0.45, row, 12, font=F["titre"], bold=True, color=C["vert"],
              shrink=False)
        write(s, n.get("headline", ""), x + 0.5, y, text_w - 0.5, row, 12.5, color=C["encre"], shrink=False)

    if img:
        photo(s, img, 7.45, 0.65, W - 7.45 - M["droite"], H - 1.3)
        cred = assets.credit(data["cover_image"])
        if cred:
            write(s, cred, 7.45, H - 0.58, W - 7.45 - M["droite"], 0.2, 7.5, color=C["gris_clair"],
                  align=PP_ALIGN.RIGHT, shrink=False)


def kpi(slide, k, left, top, width, big=48):
    value = k.get("value", "")
    size = fit_size(value, big, width, 1, 30)
    write(slide, value, left, top, width, size / 72 * 1.15, size, font=F["titre"], bold=True,
          color=C["vert"], shrink=False)
    write(slide, k.get("label", ""), left, top + size / 72 * 1.18, width, 0.62, 11.5, color=C["gris"],
          line=1.0)


def news_slide(prs, news, page, assets):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    visual = news.get("visual") or {}
    img = assets.path(visual.get("file")) if visual.get("type") == "image" else None
    chart = visual if visual.get("type") == "chart" and visual.get("series") else None

    x = M["gauche"]
    vis_x = 7.0
    col_w = (vis_x - 0.55 - x) if (img or chart) else 8.6
    kpis = [k for k in (news.get("kpis") or []) if k.get("value")][:2]

    # surtitre + titre
    kicker(s, news.get("theme", ""), x, 0.7, col_w)
    headline = news.get("headline", "")
    hsize = fit_size(headline, 30, col_w, 2, 22)
    write(s, headline, x, 1.0, col_w, 1.0, hsize, font=F["titre"], bold=True, line=0.95)

    # chiffres clés, côte à côte, sans encadré
    y = 2.25
    if kpis:
        kw = (col_w - 0.35 * (len(kpis) - 1)) / len(kpis)
        for i, k in enumerate(kpis):
            kpi(s, k, x + i * (kw + 0.35), y, kw, big=54 if len(kpis) == 1 else 44)
        y += 1.62
    hairline(s, x, y, col_w)
    y += 0.25

    # infos clés : une phrase par ligne, puce discrète ; la police baisse si la place manque
    why = news.get("why", "")
    why_top = H - 1.72
    bottom = (why_top - 0.3) if why else (H - 0.85)
    points = (news.get("points") or [])[:3]

    def heights(size):
        per_line = (col_w - 0.22) * 72 / (size * 0.5)
        return [max(1, -(-len(p) // int(per_line))) * size / 72 * 1.25 for p in points]

    psize = 13
    while psize > 10.5 and sum(heights(psize)) + 0.14 * (len(points) - 1) > bottom - y:
        psize -= 0.5
    for p, h in zip(points, heights(psize)):
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y + psize / 72 * 0.42), Inches(0.07), Inches(0.07))
        dot.fill.solid()
        dot.fill.fore_color.rgb = RGBColor.from_string(C["vert"])
        dot.line.fill.background()
        dot.shadow.inherit = False
        write(s, p, x + 0.22, y, col_w - 0.22, h + 0.05, psize, line=1.08)
        y += h + 0.14

    # pourquoi ça compte : citation en serif italique, filet vert, calée en bas de page
    if why:
        rule(s, x, why_top, 0.045, 0.95, C["vert"])
        write(s, "Pourquoi ça compte", x + 0.25, why_top - 0.02, col_w - 0.25, 0.25, 9.5, font=F["etiquette"],
              bold=True, color=C["gris"], spacing=0.8, shrink=False)
        write(s, why, x + 0.25, why_top + 0.25, col_w - 0.25, 0.72, 14, font=F["citation"], italic=True,
              line=1.05)

    # visuel
    vis_w, vis_top, vis_h = W - vis_x - M["droite"], 0.95, H - 1.65
    if img:
        photo(s, img, vis_x, vis_top, vis_w, vis_h, focus=visual.get("focus", 0.5))
        cred = assets.credit(visual["file"])
        if cred:
            write(s, cred, vis_x, vis_top + vis_h + 0.06, vis_w, 0.2, 7.5, color=C["gris_clair"],
                  align=PP_ALIGN.RIGHT, shrink=False)
    elif chart:
        ch_h = min(vis_h, 4.6)
        ch_top = vis_top + (vis_h - ch_h) / 2
        card(s, vis_x, ch_top, vis_w, ch_h)
        write(s, chart.get("title", ""), vis_x + 0.35, ch_top + 0.3, vis_w - 0.7, 0.3, 12, bold=True,
              shrink=False)
        is_round = chart.get("chart_type") in ("pie", "doughnut")
        _add_chart(
            s, chart.get("chart_type", "column"), vis_x + 0.25, ch_top + 0.7, vis_w - 0.5, ch_h - 0.95,
            chart["categories"], chart["series"],
            legend_position="bottom" if (len(chart["series"]) > 1 or is_round) else None,
            legend_font_size=10, data_labels_show=True, data_labels_font_size=11,
            data_labels_font_color="FFFFFF" if is_round else C["encre"],
            data_labels_number_format=chart.get("number_format"),
            axis_value_visible=False, axis_value_gridlines=False, gap_width=70,
            theme=CHART_THEME,
        )
        graph = [shp for shp in s.shapes if shp.has_chart][-1].chart
        if is_round:
            for j, point in enumerate(graph.plots[0].series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = RGBColor.from_string(C["graphe"][j % len(C["graphe"])])
        for axis in ("category_axis",):
            try:
                ax = getattr(graph, axis)
                ax.format.line.color.rgb = RGBColor.from_string(C["filet"])
                ax.tick_labels.font.size = Pt(10)
                ax.tick_labels.font.color.rgb = RGBColor.from_string(C["gris"])
            except Exception:
                pass

    footer(s, news, page)


# ------------------------------------------------------------------ main

def build(data: dict, out_path: str, assets_dir: str | None = None) -> dict:
    assets = Assets(assets_dir)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    cover(prs, data, assets)
    for i, news in enumerate(data.get("news") or []):
        v = news.get("visual") or {}
        if v.get("type") == "image" and not assets.path(v.get("file")):
            print(f"  image introuvable, slide sans visuel : {v.get('file')}")
        news_slide(prs, news, i + 2, assets)
    prs.save(out_path)
    report = check_deck_extended(Presentation(out_path))
    return report.get("summary", {})


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("contenu")
    ap.add_argument("sortie")
    ap.add_argument("--assets", default=None)
    a = ap.parse_args()
    content = json.loads(Path(a.contenu).read_text(encoding="utf-8"))
    summary = build(content, a.sortie, a.assets)
    print(f"Deck écrit : {a.sortie} | contrôle de mise en page : {summary}")
