import os
import io
import urllib.parse
import random
import requests
from PIL import Image, ImageDraw, ImageFont
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pydantic import BaseModel, Field

# High-contrast mobile-friendly palettes
PALETTES = [
    {
        "name": "Obsidian Cyan",
        "dark_bg": RGBColor(10, 15, 29),
        "card_bg": RGBColor(19, 29, 53),
        "card_border": RGBColor(38, 56, 96),
        "text_main": RGBColor(255, 255, 255),
        "text_muted": RGBColor(165, 180, 205),
        "accent": RGBColor(56, 189, 248),
        "accent_light": RGBColor(12, 74, 110)
    },
    {
        "name": "Deep Purple",
        "dark_bg": RGBColor(15, 12, 34),
        "card_bg": RGBColor(30, 24, 66),
        "card_border": RGBColor(60, 48, 120),
        "text_main": RGBColor(255, 255, 255),
        "text_muted": RGBColor(185, 175, 225),
        "accent": RGBColor(192, 132, 252),
        "accent_light": RGBColor(88, 28, 135)
    }
]

class MetricBox(BaseModel):
    value: str = Field(description="e.g., '1.6B', '1.2%', '$50B'")
    title: str = Field(description="Short label, max 3 words")
    desc: str = Field(description="Brief 1-sentence explanation")

class PillarCard(BaseModel):
    number: str = Field(description="'01', '02', '03'")
    title: str
    desc: str

class SlideContent(BaseModel):
    slide_number: int
    layout_type: str = Field(description="'title', 'cards_grid', 'stats_row', 'split_image', 'closing'")
    eyebrow: str
    title: str
    subtitle: str | None = None
    source_citation: str | None = None
    bullet_points: list[str] = Field(default=[])
    cards: list[PillarCard] = Field(default=[])
    metrics: list[MetricBox] = Field(default=[])
    takeaway_callout: str | None = None
    ai_image_prompt: str | None = None

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    total_slides: int
    slides: list[SlideContent]

def fetch_image_secure(prompt: str | None, custom_bytes: bytes | None = None) -> io.BytesIO | None:
    if custom_bytes:
        try:
            return io.BytesIO(custom_bytes)
        except Exception:
            pass
    if not prompt:
        return None

    clean_p = urllib.parse.quote_plus(" ".join(prompt.strip().split()[:3]))
    
    # 1. Fast AI image generation
    try:
        url = f"https://image.pollinations.ai/prompt/{clean_p}?width=1024&height=640&model=flux&nologo=true"
        resp = requests.get(url, timeout=8)
        if resp.status_code == 200 and len(resp.content) > 4000:
            bio = io.BytesIO(resp.content)
            img = Image.open(bio)
            img.verify()
            bio.seek(0)
            return bio
    except Exception:
        pass

    # 2. Picsum reliable photo fallback
    try:
        resp = requests.get(f"https://picsum.photos/seed/{clean_p}/800/600", timeout=6)
        if resp.status_code == 200 and len(resp.content) > 3000:
            return io.BytesIO(resp.content)
    except Exception:
        pass

    return None

def generate_presentation_json(topic: str, slide_count: int = 10, language: str = "uz", user_notes: str | None = None) -> MasterpiecePresentation:
    client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

    lang_map = {
        "uz": "O'zbek tili (barcha matnlar 100% ravon o'zbek tilida)",
        "ru": "Русский язык (весь текст строго на русском)",
        "en": "English language"
    }

    system_instruction = (
        f"You are an executive slide architect. Generate EXACTLY {slide_count} slides. "
        "Keep text concise, punchy, and readable on mobile devices. "
        "Never make titles longer than 40 characters. "
        "Mandatory layouts: Slide 1: 'title', Slide {slide_count}: 'closing', "
        "interleave 'cards_grid', 'stats_row', and 'split_image' in between."
    )

    prompt = f"Topic: {topic}\nSlide count: {slide_count}\nLanguage: {lang_map.get(language, 'O uzbek')}\nNotes: {user_notes or 'None'}"

    res = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            response_mime_type="application/json",
            response_schema=MasterpiecePresentation,
            temperature=0.2
        )
    )
    return res.parsed

def create_pptx_file(plan: MasterpiecePresentation, output_path: str = "output.pptx", custom_img_bytes: bytes | None = None) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    theme = random.choice(PALETTES)
    blank_layout = prs.slide_layouts[6]

    for s in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # Full Slide Background
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme["dark_bg"]
        bg.line.fill.background()

        # ---------------- 1. TITLE SLIDE ----------------
        if s.layout_type == "title":
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(2.0), Inches(0.1), Inches(3.2))
            bar.fill.solid()
            bar.fill.fore_color.rgb = theme["accent"]
            bar.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.4), Inches(1.8), Inches(10.5), Inches(3.8))
            tf = tb.text_frame
            tf.word_wrap = True

            pe = tf.paragraphs[0]
            pe.text = s.eyebrow.upper()
            pe.font.size = Pt(14)
            pe.font.bold = True
            pe.font.color.rgb = theme["accent"]
            pe.space_after = Pt(14)

            pt = tf.add_paragraph()
            pt.text = s.title
            pt.font.size = Pt(44)
            pt.font.bold = True
            pt.font.color.rgb = theme["text_main"]
            pt.space_after = Pt(16)

            if s.subtitle:
                ps = tf.add_paragraph()
                ps.text = s.subtitle
                ps.font.size = Pt(20)
                ps.font.color.rgb = theme["text_muted"]
            continue

        # ---------------- HEADER FOR CONTENT SLIDES ----------------
        hb = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.0), Inches(1.2))
        hf = hb.text_frame
        hf.word_wrap = True

        he = hf.paragraphs[0]
        he.text = s.eyebrow.upper()
        he.font.size = Pt(12)
        he.font.bold = True
        he.font.color.rgb = theme["accent"]
        he.space_after = Pt(4)

        ht = hf.add_paragraph()
        ht.text = s.title
        ht.font.size = Pt(30)
        ht.font.bold = True
        ht.font.color.rgb = theme["text_main"]

        # ---------------- 2. CARDS GRID ----------------
        if s.layout_type == "cards_grid":
            items = s.cards[:3] if s.cards else [
                PillarCard(number=f"0{i+1}", title=f"Aspekt {i+1}", desc=b)
                for i, b in enumerate(s.bullet_points[:3])
            ]
            c_w = Inches(3.64)
            c_h = Inches(4.3)
            gap = Inches(0.4)

            for i, c in enumerate(items[:3]):
                cx = Inches(0.8) + i * (c_w + gap)
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.9), c_w, c_h)
                card.fill.solid()
                card.fill.fore_color.rgb = theme["card_bg"]
                card.line.color.rgb = theme["card_border"]

                ctb = slide.shapes.add_textbox(cx + Inches(0.3), Inches(2.2), c_w - Inches(0.6), c_h - Inches(0.6))
                ctf = ctb.text_frame
                ctf.word_wrap = True

                p0 = ctf.paragraphs[0]
                p0.text = c.number
                p0.font.size = Pt(14)
                p0.font.bold = True
                p0.font.color.rgb = theme["accent"]
                p0.space_after = Pt(10)

                p1 = ctf.add_paragraph()
                p1.text = c.title
                p1.font.size = Pt(20)
                p1.font.bold = True
                p1.font.color.rgb = theme["text_main"]
                p1.space_after = Pt(12)

                p2 = ctf.add_paragraph()
                p2.text = c.desc
                p2.font.size = Pt(15)
                p2.font.color.rgb = theme["text_muted"]

        # ---------------- 3. STATS ROW ----------------
        elif s.layout_type == "stats_row":
            m_items = s.metrics[:3] if s.metrics else [
                MetricBox(value="85%", title="O'SISH", desc="Yillik ijobiy dinamika"),
                MetricBox(value="10x", title="TEZLIK", desc="Operatsion optimallashuv"),
                MetricBox(value="Zero", title="XATOLIK", desc="To'liq avtomatlashtirilgan nazorat")
            ]
            m_w = Inches(3.64)
            m_h = Inches(4.3)
            gap = Inches(0.4)

            for i, m in enumerate(m_items[:3]):
                mx = Inches(0.8) + i * (m_w + gap)
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.9), m_w, m_h)
                box.fill.solid()
                box.fill.fore_color.rgb = theme["card_bg"]
                box.line.color.rgb = theme["card_border"]

                mtb = slide.shapes.add_textbox(mx + Inches(0.3), Inches(2.4), m_w - Inches(0.6), m_h - Inches(0.8))
                mtf = mtb.text_frame
                mtf.word_wrap = True

                pv = mtf.paragraphs[0]
                pv.text = m.value
                pv.font.size = Pt(46)
                pv.font.bold = True
                pv.font.color.rgb = theme["accent"] if i == 0 else theme["text_main"]
                pv.space_after = Pt(12)

                ptit = mtf.add_paragraph()
                ptit.text = m.title.upper()
                ptit.font.size = Pt(16)
                ptit.font.bold = True
                ptit.font.color.rgb = theme["text_main"]
                ptit.space_after = Pt(10)

                pd = mtf.add_paragraph()
                pd.text = m.desc
                pd.font.size = Pt(14)
                pd.font.color.rgb = theme["text_muted"]

        # ---------------- 4. SPLIT IMAGE LAYOUT ----------------
        elif s.layout_type == "split_image":
            img_io = fetch_image_secure(s.ai_image_prompt or s.title, custom_img_bytes)

            # If image exists, make 50/50 split; otherwise, use full width for text
            text_width = Inches(6.0) if img_io else Inches(11.7)

            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.9), text_width, Inches(4.3))
            tf = tb.text_frame
            tf.word_wrap = True

            for idx, pt_text in enumerate(s.bullet_points[:4]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = f"•  {pt_text}"
                p.font.size = Pt(16)
                p.font.color.rgb = theme["text_main"]
                p.space_after = Pt(16)

            if img_io:
                try:
                    slide.shapes.add_picture(img_io, Inches(7.2), Inches(1.9), width=Inches(5.3), height=Inches(4.3))
                except Exception:
                    pass

        # ---------------- 5. CLOSING SLIDE ----------------
        elif s.layout_type == "closing":
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.3))
            box.fill.solid()
            box.fill.fore_color.rgb = theme["card_bg"]
            box.line.color.rgb = theme["card_border"]

            ctb = slide.shapes.add_textbox(Inches(1.5), Inches(2.4), Inches(10.3), Inches(3.3))
            ctf = ctb.text_frame
            ctf.word_wrap = True

            p1 = ctf.paragraphs[0]
            p1.text = "STRATEGIK XULOSA"
            p1.font.size = Pt(14)
            p1.font.bold = True
            p1.font.color.rgb = theme["accent"]
            p1.space_after = Pt(14)

            p2 = ctf.add_paragraph()
            p2.text = s.title.upper()
            p2.font.size = Pt(34)
            p2.font.bold = True
            p2.font.color.rgb = theme["text_main"]
            p2.space_after = Pt(16)

            if s.takeaway_callout or s.subtitle:
                p3 = ctf.add_paragraph()
                p3.text = s.takeaway_callout or s.subtitle
                p3.font.size = Pt(18)
                p3.font.color.rgb = theme["text_muted"]

        # Footer citation
        if s.source_citation:
            fc = slide.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
            fp = fc.text_frame.paragraphs[0]
            fp.text = s.source_citation
            fp.font.size = Pt(11)
            fp.font.color.rgb = theme["text_muted"]

    prs.save(output_path)
    return output_path
