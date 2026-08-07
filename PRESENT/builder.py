from pathlib import Path
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

import design


def _set_text_style(paragraph, size, color, bold=False, font_name=None):
    paragraph.font.size = size
    paragraph.font.color.rgb = color
    paragraph.font.bold = bold
    paragraph.font.name = font_name or design.FONT_BODY


def _add_title(slide, text, hero=False):
    box = slide.shapes.add_textbox(
        design.MARGIN_X,
        design.MARGIN_Y,
        Inches(11.7),
        Inches(1.0),
    )
    p = box.text_frame.paragraphs[0]
    p.text = text
    _set_text_style(
        p,
        design.HERO_SIZE if hero else design.TITLE_SIZE,
        design.CHARCOAL,
        bold=True,
        font_name=design.FONT_HEADLINE,
    )
    return box


def _add_subtitle(slide, text, top=1.55):
    box = slide.shapes.add_textbox(
        design.MARGIN_X,
        Inches(top),
        Inches(8.8),
        Inches(0.7),
    )
    p = box.text_frame.paragraphs[0]
    p.text = text
    _set_text_style(p, design.BODY_SIZE, design.MUTED)
    return box


def build_hero(prs, slide_spec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = design.PEARL_WHITE
    _add_title(slide, slide_spec["headline"], hero=True)
    if slide_spec.get("subheadline"):
        _add_subtitle(slide, slide_spec["subheadline"], top=1.75)

    image = slide_spec.get("image")
    if image and Path(image).exists():
        slide.shapes.add_picture(image, Inches(8.5), Inches(0.65), width=Inches(4.1))
    return slide


def build_comparison(prs, slide_spec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = design.PEARL_WHITE
    _add_title(slide, slide_spec["headline"])

    columns = [
        (slide_spec.get("left_title", "Left"), slide_spec.get("left", []), 0.8),
        (slide_spec.get("right_title", "Right"), slide_spec.get("right", []), 6.85),
    ]

    for heading, items, x in columns:
        box = slide.shapes.add_textbox(Inches(x), Inches(1.75), Inches(5.65), Inches(4.8))
        tf = box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = heading
        _set_text_style(p, design.TITLE_SIZE, design.DEEP_BLUE, bold=True, font_name=design.FONT_HEADLINE)
        for item in items:
            p = tf.add_paragraph()
            p.text = item
            p.space_before = design.SMALL_SIZE
            _set_text_style(p, design.BODY_SIZE, design.CHARCOAL)
    return slide


def build_content(prs, slide_spec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = design.PEARL_WHITE
    _add_title(slide, slide_spec["headline"])

    if slide_spec.get("subheadline"):
        _add_subtitle(slide, slide_spec["subheadline"], top=1.45)

    box = slide.shapes.add_textbox(Inches(1.0), Inches(2.15), Inches(11.2), Inches(4.55))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True

    items = slide_spec.get("items", [])
    if not items:
        items = [slide_spec.get("body", "")]

    for index, item in enumerate(items):
        p = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        p.text = item
        p.space_after = design.SMALL_SIZE
        _set_text_style(p, design.BODY_SIZE, design.CHARCOAL)
    return slide


def build_statement(prs, slide_spec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = design.DEEP_BLUE

    box = slide.shapes.add_textbox(Inches(1.15), Inches(2.05), Inches(11.0), Inches(2.2))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = slide_spec["headline"]
    p.alignment = PP_ALIGN.CENTER
    _set_text_style(p, design.HERO_SIZE, design.WHITE, bold=True, font_name=design.FONT_HEADLINE)

    if slide_spec.get("subheadline"):
        p2 = tf.add_paragraph()
        p2.text = slide_spec["subheadline"]
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = design.BODY_SIZE
        _set_text_style(p2, design.BODY_SIZE, design.PEARL_WHITE)
    return slide


BUILDERS = {
    "hero": build_hero,
    "comparison": build_comparison,
    "content": build_content,
    "statement": build_statement,
}


def add_slides(prs, spec: dict):
    for slide_spec in spec.get("slides", []):
        slide_type = slide_spec.get("type")
        if slide_type not in BUILDERS:
            raise ValueError(f"Unsupported slide type: {slide_type}")
        BUILDERS[slide_type](prs, slide_spec)
    return prs


def build_deck(spec: dict, output_path: str):
    prs = Presentation()
    prs.slide_width = design.SLIDE_WIDTH
    prs.slide_height = design.SLIDE_HEIGHT
    add_slides(prs, spec)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path
