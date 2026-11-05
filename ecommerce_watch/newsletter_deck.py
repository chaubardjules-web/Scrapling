"""Génère la newsletter Veille E-commerce en PowerPoint, toujours sur le même gabarit.

Gabarit : charte Deloitte (noir, blanc, vert Deloitte), police Aptos, 16:9.
Structure : une page de garde (avec sommaire), puis une slide par news.
Moteur de mise en page : pptx-mcp-server (https://github.com/knorq-ai/pptx-mcp-server).

Usage : python ecommerce_watch/newsletter_deck.py contenu.json sortie.pptx
        python ecommerce_watch/newsletter_deck.py --icons   (liste les icônes disponibles)

Format de contenu.json :
{
  "date": "5 nov. 2026",
  "news": [                                    // 5 à 8 news
    {
      "theme": "Agentic commerce",             // thème affiché en surtitre
      "headline": "Titre d'action, 70 caractères max",
      "kpis": [                                // 1 ou 2 gros chiffres, réels et sourcés
        {"value": "+45 %", "label": "ce que mesure le chiffre (60 caractères max)"}
      ],
      "points": ["info clé 1", "info clé 2", "info clé 3"],   // 3 ou 4 puces, 160 car. max
      "why": "Pourquoi ça compte, 2 phrases max",
      "visual": {"type": "chart", "chart_type": "column", "title": "Titre du graphe",
                 "categories": ["2024", "2025"], "series": [{"name": "Ventes", "values": [1.2, 2.4]}],
                 "number_format": "0.0"}
             // ou {"type": "icon", "icon": "<clé de ICONS : panier, livraison, colis, paiement,
             //      recherche, recherche_ia, magasin, outils, tech, graphique, idee, mobile, web,
             //      data, finance, createurs, personne>"}
      "source": "Modern Retail",
      "url": "https://..."
    }
  ]
}
"""

from __future__ import annotations

import copy
import json
import re
import sys
from urllib.parse import urlparse

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from pptx_mcp_server.engine.charts import _add_chart
from pptx_mcp_server.engine.composites import _build_slide
from pptx_mcp_server.engine.icons import _add_icon
from pptx_mcp_server.engine.shapes import add_auto_fit_textbox
from pptx_mcp_server.engine.validation import check_deck_extended
from pptx_mcp_server.theme import DELOITTE, register_theme

# ------------------------------------------------------------------ charte

FONT = "Aptos"
BLACK, WHITE = "000000", "FFFFFF"
GREEN = "86BC25"        # vert Deloitte
DARK_GREEN = "26890D"
TEAL = "0D8390"
GREEN_TINT = "EEF6E1"
GRAY_BG = "F2F2F2"
GRAY_TEXT = "53565A"
GRAY_LIGHT = "97999B"

THEME = copy.deepcopy(DELOITTE)
THEME.name = "deloitte_conseil"
THEME.colors.update({
    "primary": f"#{BLACK}", "accent": f"#{GREEN}", "text_secondary": f"#{GRAY_TEXT}",
    "footnote": f"#{GRAY_LIGHT}", "dark_green": f"#{DARK_GREEN}", "bg_alt": f"#{GRAY_BG}",
})
THEME.fonts.update({"title": FONT, "body": FONT})
THEME.margins.update({"left": 0.6, "right": 0.6})
THEME.layout.update({"title_top": 0.5, "title_height": 0.6, "divider_top": 1.17,
                     "body_top": 1.4, "footer_top": 6.95})
THEME.sizes.update({"title": 24, "footnote": 9})
THEME.chart_colors = [f"#{c}" for c in (GREEN, BLACK, DARK_GREEN, TEAL, GRAY_LIGHT, "C4D600")]
register_theme(THEME)

W, H = 13.333, 7.5
L, R = 0.6, 0.6
BODY_W = W - L - R
TOP, BOTTOM = 1.45, 6.75

# Icônes vérifiées visuellement (certains identifiants de la bibliothèque ne
# correspondent pas à leur dessin) : clé utilisable dans "visual.icon" -> id moteur
ICONS = {
    "panier": "shopping_cart", "livraison": "van", "colis": "box", "paiement": "credit_card",
    "recherche": "magnifying_glass", "recherche_ia": "search", "magasin": "building",
    "outils": "tools", "tech": "gears", "graphique": "chart", "idee": "bulb",
    "mobile": "target", "web": "browser", "data": "data", "finance": "network_pc",
    "createurs": "chat", "personne": "people",
}
# Icône par défaut selon le thème quand aucune illustration n'est fournie
THEME_ICONS = {
    "agentic": "panier", "geo": "recherche_ia", "search": "recherche_ia", "social": "mobile",
    "tiktok": "mobile", "creat": "createurs", "outil": "outils", "tool": "outils",
    "retail": "magasin", "marketplace": "colis", "logisti": "livraison", "commerce": "panier",
    "beauty": "createurs", "luxe": "magasin", "paiement": "paiement", "payment": "paiement",
}


# ------------------------------------------------------------------ primitives

def rect(slide, left, top, width, height, fill, shape=MSO_SHAPE.RECTANGLE):
    shp = slide.shapes.add_shape(shape, Inches(left), Inches(top), Inches(width), Inches(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = RGBColor.from_string(fill)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def nb(value: str) -> str:
    """Espaces insécables typographiques : 20 %, 500 M€, 10 000, 48 h, « ... »."""
    value = re.sub(r"(\d) (?=(%|M€|Md€|M\$|Md\$|Md|k€|€|\$|h\b|\d{3}\b))", "\\1\u00a0", value or "")
    return value.replace("« ", "«\u00a0").replace(" »", "\u00a0»")


def text(slide, value, left, top, width, height, size, color=BLACK, bold=False,
         min_size=8, align="left", anchor="top"):
    add_auto_fit_textbox(slide, nb(value), left, top, width, height, font_name=FONT,
                         font_size_pt=size, min_size_pt=min_size, bold=bold, color_hex=color,
                         align=align, vertical_anchor=anchor)


def label(slide, value, left, top, width, color=GREEN, size=10):
    text(slide, value.upper(), left, top, width, 0.3, size, color=color, bold=True, min_size=size)


def bullet_block(slide, points, left, top, width, height):
    """Un seul bloc de puces : même taille de police partout, retrait suspendu, puce verte."""
    from pptx.oxml.ns import qn
    from pptx.enum.text import MSO_AUTO_SIZE

    chars = sum(len(p) for p in points)
    size = 14 if chars <= 420 else 13 if chars <= 520 else 12
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(0.02))
    for i, point in enumerate(points):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(9)
        p.line_spacing = 1.05
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(Inches(0.28))))
        pPr.set("indent", str(-int(Inches(0.28))))
        for tag in ("a:buClr", "a:buSzPct", "a:buFont", "a:buChar"):
            el = pPr.makeelement(qn(tag), {})
            if tag == "a:buClr":
                clr = el.makeelement(qn("a:srgbClr"), {"val": GREEN})
                el.append(clr)
            elif tag == "a:buSzPct":
                el.set("val", "90000")
            elif tag == "a:buFont":
                el.set("typeface", "Arial")
            else:
                el.set("char", "■")
            pPr.append(el)
        run = p.add_run()
        run.text = nb(point)
        run.font.name, run.font.size = FONT, Pt(size)
        run.font.color.rgb = RGBColor.from_string(BLACK)


def domain(url):
    try:
        return urlparse(url).netloc.removeprefix("www.")
    except Exception:
        return ""


def pick_icon(news):
    wanted = (news.get("visual") or {}).get("icon")
    if wanted in ICONS:
        return ICONS[wanted]
    theme = (news.get("theme") or "").lower()
    for key, icon in THEME_ICONS.items():
        if key in theme:
            return ICONS[icon]
    return ICONS["idee"]


# ------------------------------------------------------------------ slides

def cover(prs, data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string(BLACK)

    rect(slide, L, 1.6, 0.12, 2.6, GREEN)
    text(slide, "Veille", 1.0, 1.45, 5.6, 0.95, 54, WHITE, bold=True, min_size=40)
    text(slide, "E-commerce", 1.0, 2.35, 5.6, 0.95, 54, GREEN, bold=True, min_size=40)
    text(slide, f"Édition du {data['date']}", 1.0, 3.45, 5.6, 0.5, 18, WHITE, min_size=14)
    text(slide, "Agentic commerce · GEO · Social commerce · Outils", 1.0, 4.6, 5.6, 0.8, 14,
         GRAY_LIGHT, min_size=11)
    text(slide, "Deloitte Conseil  |  Technology & Transformation", L, 6.75, 6.0, 0.35, 11,
         GRAY_LIGHT, min_size=10)

    # sommaire
    x, w = 7.2, W - 7.2 - R
    label(slide, "Au sommaire", x, 1.45, w)
    news = data.get("news") or []
    row = min(0.9, 4.9 / max(len(news), 1))
    for i, n in enumerate(news):
        y = 1.9 + i * row
        text(slide, f"{i + 2:02d}", x, y, 0.6, row, 16, GREEN, bold=True, min_size=12)
        text(slide, n.get("theme", "").upper(), x + 0.65, y, w - 0.65, 0.25, 9, GRAY_LIGHT,
             bold=True, min_size=8)
        text(slide, n.get("headline", ""), x + 0.65, y + 0.24, w - 0.65, row - 0.32, 13, WHITE,
             min_size=9)


def kpi_column(slide, kpis, x, w):
    kpis = [k for k in (kpis or []) if k.get("value")][:2]
    if not kpis:
        return False
    gap = 0.25
    h = (BOTTOM - TOP - gap * (len(kpis) - 1)) / len(kpis)
    for i, k in enumerate(kpis):
        y = TOP + i * (h + gap)
        rect(slide, x, y, w, h, BLACK)
        rect(slide, x, y, w, 0.07, GREEN)
        val_h = 1.3 if len(kpis) == 1 else 0.95
        text(slide, k["value"], x + 0.3, y + 0.3, w - 0.6, val_h, 60 if len(kpis) == 1 else 44,
             GREEN, bold=True, min_size=28)
        text(slide, k.get("label", ""), x + 0.3, y + 0.35 + val_h, w - 0.6, h - val_h - 0.55, 14,
             WHITE, min_size=10)
    return True


def info_column(slide, news, x, w):
    label(slide, "Les infos clés", x, TOP, w)
    points = (news.get("points") or [])[:4]
    why_h = 1.75
    list_top, list_bottom = TOP + 0.45, BOTTOM - why_h - 0.3
    bullet_block(slide, points, x, list_top, w, list_bottom - list_top)
    y = BOTTOM - why_h
    rect(slide, x, y, w, why_h, GREEN_TINT)
    rect(slide, x, y, 0.07, why_h, DARK_GREEN)
    label(slide, "Pourquoi ça compte", x + 0.3, y + 0.15, w - 0.5, color=DARK_GREEN, size=9)
    text(slide, news.get("why", ""), x + 0.3, y + 0.47, w - 0.55, why_h - 0.6, 13, BLACK, min_size=9)


def visual_column(slide, news, x, w):
    visual = news.get("visual") or {}
    rect(slide, x, TOP, w, BOTTOM - TOP, GRAY_BG)
    if visual.get("type") == "chart" and visual.get("categories") and visual.get("series"):
        try:
            label(slide, visual.get("title") or "En chiffres", x + 0.25, TOP + 0.2, w - 0.5,
                  color=DARK_GREEN, size=9)
            _add_chart(
                slide, visual.get("chart_type", "column"), x + 0.15, TOP + 0.55, w - 0.3,
                BOTTOM - TOP - 0.75, visual["categories"], visual["series"],
                legend_position="bottom" if len(visual["series"]) > 1 else None,
                legend_font_size=9, data_labels_show=True, data_labels_font_size=10,
                data_labels_font_color=WHITE if visual.get("chart_type") in ("pie", "doughnut") else None,
                data_labels_number_format=visual.get("number_format"),
                axis_value_visible=False, axis_value_gridlines=False, gap_width=60,
                theme=THEME,
            )
            if visual.get("chart_type") in ("pie", "doughnut"):
                chart = [shp for shp in slide.shapes if shp.has_chart][-1].chart
                for j, point in enumerate(chart.plots[0].series[0].points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = RGBColor.from_string(
                        (GREEN, BLACK, DARK_GREEN, GRAY_LIGHT, TEAL)[j % 5])
            return
        except Exception as exc:  # un graphe invalide ne doit pas casser le deck
            print(f"  graphe ignoré ({news.get('headline', '')[:40]}): {exc}")
    size = min(w - 1.2, 2.4)
    cx, cy = x + w / 2, TOP + (BOTTOM - TOP) / 2
    rect(slide, cx - size / 2 - 0.35, cy - size / 2 - 0.35, size + 0.7, size + 0.7, WHITE,
         shape=MSO_SHAPE.OVAL)
    _add_icon(slide, pick_icon(news), cx - size / 2, cy - size / 2, width=size, height=size,
              color=GREEN, theme=THEME)


def news_slide(prs, news, page):
    src = news.get("source", "")
    dom = domain(news.get("url", ""))
    footer = "Source : " + " — ".join(v for v in (src, dom) if v) if (src or dom) else None
    _, idx = _build_slide(prs, {"layout": "content", "title": news.get("headline", ""),
                                "page_number": page, "source": footer, "theme": THEME.name},
                          theme=THEME)
    slide = prs.slides[idx]
    label(slide, news.get("theme", ""), L, 0.25, BODY_W)

    gap = 0.3
    kpi_w, vis_w = 3.2, 3.35
    has_kpi = bool([k for k in (news.get("kpis") or []) if k.get("value")])
    if has_kpi:
        kpi_column(slide, news.get("kpis"), L, kpi_w)
        info_x = L + kpi_w + gap
    else:
        info_x = L
    info_w = W - R - vis_w - gap - info_x
    info_column(slide, news, info_x, info_w)
    visual_column(slide, news, W - R - vis_w, vis_w)

    if news.get("url"):  # lien cliquable sur la note de source
        for shp in slide.shapes:
            if shp.has_text_frame and shp.text_frame.text.startswith("Source :"):
                for p in shp.text_frame.paragraphs:
                    for run in p.runs:
                        run.hyperlink.address = news["url"]


# ------------------------------------------------------------------ main

def build(data: dict, out_path: str) -> dict:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    cover(prs, data)
    for i, news in enumerate(data.get("news") or []):
        news_slide(prs, news, i + 2)
    for slide in prs.slides:
        for shp in slide.shapes:
            if shp.has_text_frame:
                for p in shp.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.name = FONT
    prs.save(out_path)
    report = check_deck_extended(Presentation(out_path))
    return report.get("summary", {})


if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1] == "--icons":
        print(", ".join(sorted(ICONS)))
        sys.exit(0)
    if len(sys.argv) != 3:
        sys.exit("Usage : python newsletter_deck.py contenu.json sortie.pptx")
    with open(sys.argv[1], encoding="utf-8") as f:
        content = json.load(f)
    summary = build(content, sys.argv[2])
    print(f"Deck écrit : {sys.argv[2]} | contrôle de mise en page : {summary}")
