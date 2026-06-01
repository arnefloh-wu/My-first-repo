from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

W = prs.slide_width
H = prs.slide_height

def rgb(r, g, b): return RGBColor(r, g, b)

def add_rect(slide, l, t, w, h, fill_rgb=None, border_rgb=None, border_pt=0):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    shape.line.width = Pt(border_pt)
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()
    if border_rgb and border_pt:
        shape.line.color.rgb = border_rgb
    else:
        shape.line.fill.background()
    return shape

def add_textbox(slide, l, t, w, h, text, size, bold=False, color=rgb(0,0,0), align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(l, t, w, h)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txBox

pad = Inches(0.5)

# --- Header background ---
hdr_h = Inches(1.6)
add_rect(slide, 0, 0, W, hdr_h, fill_rgb=rgb(26, 26, 46))

# Header title
add_textbox(slide, pad, Inches(0.18), W - pad, Inches(0.45),
            "Masterarbeitsthemen", 28, bold=True, color=rgb(255,255,255))

# Header subtitle
add_textbox(slide, pad, Inches(0.62), W - pad, Inches(0.3),
            "Betreute Abschlussarbeiten · Ausschreibung 2026", 11,
            color=rgb(180,180,200))

# NEOH badge background
badge_t = Inches(1.0)
badge_h = Inches(0.38)
badge_w = Inches(4.2)
add_rect(slide, pad, badge_t, badge_w, badge_h,
         fill_rgb=rgb(50,50,80), border_rgb=rgb(100,100,140), border_pt=0.5)
add_textbox(slide, pad + Inches(0.15), badge_t + Inches(0.05), badge_w, badge_h,
            "✦  Themen 1 & 2 in Kooperation mit  NEOH", 10,
            color=rgb(232, 197, 71))

# --- Thesis cards ---
topics = [
    ("01", "Werbewirkungsforschung", rgb(160,100,0), rgb(255,243,205),
     "Werbewirkungsmessung einer NEOH TV-Kampagne auf dem deutschen Markt",
     "Empirische Analyse der Werbewirkung einer NEOH Fernsehwerbekampagne in Deutschland — "
     "inkl. Messung von Bekanntheit, Einstellungsänderung und Kaufabsicht anhand geeigneter "
     "Wirkungsmodelle und Methoden der Marktforschung."),
    ("02", "Marketing Analytics", rgb(160,100,0), rgb(255,243,205),
     "Marketing Mix Modelling für NEOH",
     "Entwicklung und Anwendung eines Marketing Mix Models zur quantitativen Analyse der Treiber "
     "von Absatz und Umsatz bei NEOH — inkl. Dekomposition des Mediamix, ROI-Messung einzelner "
     "Kanäle, Handlungsempfehlungen für die Budgetallokation sowie Länderbudgetallokation."),
    ("03", "AI & Marketing", rgb(26,111,168), rgb(232,244,253),
     "Artificial Intelligence in (International) Marketing & Business",
     "Empirische Untersuchung des Einsatzes und der Wirkung von KI-Technologien im Marketing "
     "und internationalen Geschäftsumfeld — z. B. Personalisierung, Automatisierung, generative "
     "KI oder KI-gestützte Entscheidungsprozesse."),
    ("04", "Brand Growth", rgb(30,126,70), rgb(234,247,239),
     "How Small Brands Grow (Internationally)",
     "Empirische Analyse der Wachstumsstrategien kleiner Marken im nationalen und internationalen "
     "Kontext — unter Berücksichtigung von Markenbekanntheit, Distributionsaufbau, Positionierung "
     "und den Gesetzmäßigkeiten des empirischen Marketings (z. B. Ehrenberg-Bass)."),
    ("05", "Replikation", rgb(107,63,160), rgb(243,238,255),
     "Replikationsstudien",
     "Empirische Replikation bestehender Studien aus dem Bereich Marketing oder internationales "
     "Business — zur Überprüfung der Generalisierbarkeit von Befunden in neuen Kontexten, "
     "Märkten oder Zeiträumen."),
]

card_h = Inches(0.88)
card_gap = Inches(0.07)
card_top = hdr_h + Inches(0.18)
card_l = pad
card_w = W - 2 * pad

for i, (num, tag_text, tag_fg, tag_bg, title, desc) in enumerate(topics):
    t = card_top + i * (card_h + card_gap)

    # Card background + border
    add_rect(slide, card_l, t, card_w, card_h,
             fill_rgb=rgb(255,255,255), border_rgb=rgb(220,225,235), border_pt=0.75)

    # Number
    add_textbox(slide, card_l + Inches(0.12), t + Inches(0.12), Inches(0.45), card_h,
                num, 22, bold=True, color=rgb(220,220,220))

    # Tag pill
    tag_l = card_l + Inches(0.6)
    tag_t = t + Inches(0.1)
    tag_w = Inches(1.7)
    tag_pill_h = Inches(0.22)
    add_rect(slide, tag_l, tag_t, tag_w, tag_pill_h, fill_rgb=tag_bg)
    add_textbox(slide, tag_l + Inches(0.06), tag_t, tag_w, tag_pill_h,
                tag_text.upper(), 7, bold=True, color=tag_fg)

    # Title
    add_textbox(slide, card_l + Inches(0.6), t + Inches(0.33), card_w - Inches(0.75), Inches(0.28),
                title, 11, bold=True, color=rgb(26,26,46))

    # Description
    add_textbox(slide, card_l + Inches(0.6), t + Inches(0.58), card_w - Inches(0.75), Inches(0.3),
                desc, 8, color=rgb(90,106,128))

# --- Hints section ---
hint_top = card_top + 5 * (card_h + card_gap) + Inches(0.1)
hint_h = Inches(0.52)
add_rect(slide, pad, hint_top, card_w, hint_h,
         fill_rgb=rgb(247,248,250), border_rgb=rgb(228,232,238), border_pt=0.75)

add_textbox(slide, pad + Inches(0.18), hint_top + Inches(0.05), card_w, Inches(0.18),
            "ALLGEMEINE HINWEISE", 7, bold=True, color=rgb(138,150,168))
add_textbox(slide, pad + Inches(0.18), hint_top + Inches(0.22), card_w, Inches(0.14),
            "✦  Es werden nur empirische Arbeiten betreut (qualitativ oder quantitativ).    "
            "✦  Die Arbeiten können auf Deutsch oder Englisch verfasst werden.", 9,
            color=rgb(74,85,104))

# --- Footer ---
footer_h = Inches(0.28)
footer_t = H - footer_h
add_rect(slide, 0, footer_t, W, footer_h, fill_rgb=rgb(247,248,250))
add_textbox(slide, pad, footer_t + Inches(0.05), Inches(4), footer_h,
            "Kooperationspartner: NEOH GmbH", 8, color=rgb(170,170,170))
add_textbox(slide, W - Inches(2), footer_t + Inches(0.05), Inches(1.8), footer_h,
            "Stand: Mai 2026", 8, color=rgb(170,170,170), align=PP_ALIGN.RIGHT)

prs.save("/home/user/My-first-repo/masterarbeitsthemen.pptx")
print("Done")
