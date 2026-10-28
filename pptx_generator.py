import os
import io
import requests
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pydantic import BaseModel, Field

# 1. Pydantic Structured Schema
class MetricItem(BaseModel):
    value: str = Field(description="e.g. '80%', '10x', 'Zero'")
    label: str = Field(description="Short description under the metric")

class CardItem(BaseModel):
    title: str
    description: str

class SlideStructure(BaseModel):
    slide_number: int
    layout_type: str = Field(
        description="One of: 'title', 'cards', 'split_image', 'metrics', 'comparison', 'closing'"
    )
    eyebrow: str = Field(description="Category tag (e.g., 'ARCHITECTURAL SHIFT', 'MACROECONOMICS')")
    title: str = Field(description="Max 45 chars, punchy executive headline")
    subtitle: str | None = None
    bullet_points: list[str] = Field(default=[], description="Bullet points for split or standard slides")
    cards: list[CardItem] = Field(default=[], description="Exactly 2 or 3 cards for 'cards' or 'comparison' layouts")
    metrics: list[MetricItem] = Field(default=[], description="Exactly 3 metrics for 'metrics' layout")
    key_takeaway: str | None = Field(default=None, description="One strong synthesis sentence")
    image_query: str | None = Field(default=None, description="2-3 English keywords for Unsplash photo")

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    slides: list[SlideStructure]

# 2. Gemini Generation Function
def generate_presentation_json(topic: str, slide_count: int = 6, language: str = "uz") -> MasterpiecePresentation:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_map = {
        "uz": "O'zbek tili (barcha matnlar 100% o'zbek tilida bo'lsin)",
        "ru": "Русский язык (весь текст слайдов строго на русском языке)",
        "en": "English language (executive presentation style)"
    }
    target_language = lang_map.get(language, "O'zbek tili")

    system_instruction = (
        "You are a Principal Presentation Designer at McKinsey/BCG. "
        "Convert the user topic into a visually diverse executive slide deck. "
        "Vary the layouts across slides:\n"
        "- Slide 1: 'title'\n"
        "- Slide 2: 'cards' (3 distinct pillars/features)\n"
        "- Slide 3: 'split_image' (visual + structured bullet points)\n"
        "- Slide 4: 'metrics' (3 large numbers/stats with labels)\n"
        "- Slide 5: 'comparison' (2 structured pillar cards)\n"
        "- Slide 6: 'closing' (high-level synthesis/vision)\n"
        "Never output boring walls of text. Be concise, punchy, and follow the JSON schema strictly."
    )

    prompt = f"Topic: {topic}\nDesired slides: {slide_count}\nLanguage: {target_language}"

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            response_mime_type="application/json",
            response_schema=MasterpiecePresentation,
            temperature=0.25
        ),
    )
    return response.parsed

# Helper: Fetch thematic photo
def fetch_photo(query: str | None) -> io.BytesIO | None:
    if not query:
        return None
    try:
        clean_q = requests.utils.quote(query)
        url = f"https://loremflickr.com/800/600/{clean_q}"
        resp = requests.get(url, timeout=5)
        if resp.status_code == 200 and len(resp.content) > 2000:
            return io.BytesIO(resp.content)
    except Exception:
        pass
    return None

# 3. PPTX Builder Engine (Draws cards, metrics, colors, and banners)
def create_pptx_file(plan: MasterpiecePresentation, output_path: str = "output.pptx") -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Color Palette
    C_DARK_BG = RGBColor(9, 13, 22)       # Executive Slate/Black
    C_LIGHT_BG = RGBColor(255, 255, 255)  # Pure White
    C_CARD_BG = RGBColor(248, 250, 252)   # Slate 50
    C_CARD_BORDER = RGBColor(226, 232, 240)# Slate 200
    C_PRIMARY = RGBColor(15, 23, 42)      # Slate 900
    C_CYAN = RGBColor(2, 132, 199)        # Cyan 600
    C_CYAN_LIGHT = RGBColor(224, 242, 254)# Cyan 100
    C_MUTED = RGBColor(100, 116, 139)     # Slate 500
    C_WHITE = RGBColor(248, 250, 252)

    blank_layout = prs.slide_layouts[6]

    for s in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # ------------------- 1. TITLE SLIDE (DARK) -------------------
        if s.layout_type == "title":
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = C_DARK_BG
            bg.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(11.0), Inches(3.8))
            tf = tb.text_frame
            tf.word_wrap = True

            # Eyebrow
            p_eye = tf.paragraphs[0]
            p_eye.text = s.eyebrow.upper()
            p_eye.font.size = Pt(13)
            p_eye.font.bold = True
            p_eye.font.color.rgb = C_CYAN
            p_eye.space_after = Pt(16)

            # Title
            p_title = tf.add_paragraph()
            p_title.text = s.title
            p_title.font.size = Pt(46)
            p_title.font.bold = True
            p_title.font.color.rgb = C_WHITE
            p_title.space_after = Pt(20)

            # Subtitle
            if s.subtitle:
                p_sub = tf.add_paragraph()
                p_sub.text = s.subtitle
                p_sub.font.size = Pt(20)
                p_sub.font.color.rgb = C_MUTED

            # Decorative cyan line
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(5.8), Inches(1.0), Inches(0.06))
            line.fill.solid()
            line.fill.fore_color.rgb = C_CYAN
            line.line.fill.background()
            continue

        # ------------------- 6. CLOSING SLIDE (DARK) -------------------
        if s.layout_type == "closing":
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = C_DARK_BG
            bg.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.8), Inches(1.8), Inches(9.7), Inches(4.0))
            tf = tb.text_frame
            tf.word_wrap = True

            p_eye = tf.paragraphs[0]
            p_eye.text = s.eyebrow.upper()
            p_eye.font.size = Pt(13)
            p_eye.font.bold = True
            p_eye.font.color.rgb = C_CYAN
            p_eye.alignment = PP_ALIGN.CENTER
            p_eye.space_after = Pt(14)

            p_title = tf.add_paragraph()
            p_title.text = s.title
            p_title.font.size = Pt(40)
            p_title.font.bold = True
            p_title.font.color.rgb = C_WHITE
            p_title.alignment = PP_ALIGN.CENTER
            p_title.space_after = Pt(24)

            if s.subtitle or s.key_takeaway:
                p_sub = tf.add_paragraph()
                p_sub.text = s.key_takeaway or s.subtitle
                p_sub.font.size = Pt(18)
                p_sub.font.color.rgb = C_MUTED
                p_sub.alignment = PP_ALIGN.CENTER
            continue

        # ------------------- BASE FOR LIGHT CONTENT SLIDES -------------------
        # Header Eyebrow & Title
        ht_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.2))
        htf = ht_box.text_frame
        htf.word_wrap = True
        
        pe = htf.paragraphs[0]
        pe.text = s.eyebrow.upper()
        pe.font.size = Pt(11)
        pe.font.bold = True
        pe.font.color.rgb = C_CYAN

        pt = htf.add_paragraph()
        pt.text = s.title
        pt.font.size = Pt(28)
        pt.font.bold = True
        pt.font.color.rgb = C_PRIMARY

        # ------------------- 2. CARDS LAYOUT (3 Bento Cards) -------------------
        if s.layout_type == "cards":
            items = s.cards if s.cards else [CardItem(title=f"Pillar {i+1}", description=b) for i, b in enumerate(s.bullet_points[:3])]
            num_cards = min(len(items), 3) or 3
            card_w = Inches(3.64)
            card_h = Inches(3.8)
            gap = Inches(0.39)
            start_x = Inches(0.8)
            start_y = Inches(1.8)

            for i, card_data in enumerate(items[:3]):
                cx = start_x + i * (card_w + gap)
                # Outer Card
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, start_y, card_w, card_h)
                card.fill.solid()
                card.fill.fore_color.rgb = C_CARD_BG
                card.line.color.rgb = C_CARD_BORDER

                # Small Accent Dot/Square inside
                icon_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.3), start_y + Inches(0.35), Inches(0.45), Inches(0.45))
                icon_box.fill.solid()
                icon_box.fill.fore_color.rgb = C_CYAN_LIGHT
                icon_box.line.fill.background()

                # Card Text
                ctb = slide.shapes.add_textbox(cx + Inches(0.3), start_y + Inches(1.0), card_w - Inches(0.6), card_h - Inches(1.2))
                ctf = ctb.text_frame
                ctf.word_wrap = True

                cp1 = ctf.paragraphs[0]
                cp1.text = card_data.title
                cp1.font.size = Pt(18)
                cp1.font.bold = True
                cp1.font.color.rgb = C_PRIMARY
                cp1.space_after = Pt(10)

                cp2 = ctf.add_paragraph()
                cp2.text = card_data.description
                cp2.font.size = Pt(14)
                cp2.font.color.rgb = C_MUTED

        # ------------------- 3. SPLIT IMAGE LAYOUT -------------------
        elif s.layout_type == "split_image":
            # Left Column (Bullet List)
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.5), Inches(4.0))
            tf = tb.text_frame
            tf.word_wrap = True

            for idx, pt_text in enumerate(s.bullet_points[:4]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = f"•  {pt_text}"
                p.font.size = Pt(15)
                p.font.color.rgb = C_PRIMARY
                p.space_after = Pt(16)

            # Right Column (Picture)
            img_stream = fetch_photo(s.image_query or s.title)
            if img_stream:
                try:
                    slide.shapes.add_picture(img_stream, Inches(7.6), Inches(1.8), width=Inches(4.9), height=Inches(3.8))
                except Exception:
                    pass

        # ------------------- 4. METRICS / STATS LAYOUT -------------------
        elif s.layout_type == "metrics":
            m_items = s.metrics if s.metrics else [
                MetricItem(value="80%", label="Workflows automated"),
                MetricItem(value="10x", label="Efficiency gains achieved"),
                MetricItem(value="Zero", label="Latency in execution")
            ]
            m_w = Inches(3.64)
            m_h = Inches(2.2)
            gap = Inches(0.39)
            start_x = Inches(0.8)

            for i, m in enumerate(m_items[:3]):
                mx = start_x + i * (m_w + gap)
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.9), m_w, m_h)
                box.fill.solid()
                box.fill.fore_color.rgb = C_CARD_BG
                box.line.color.rgb = C_CARD_BORDER

                mtb = slide.shapes.add_textbox(mx + Inches(0.3), Inches(2.1), m_w - Inches(0.6), Inches(1.7))
                mtf = mtb.text_frame
                mtf.word_wrap = True

                p_val = mtf.paragraphs[0]
                p_val.text = m.value
                p_val.font.size = Pt(42)
                p_val.font.bold = True
                p_val.font.color.rgb = C_CYAN if i % 2 == 0 else C_PRIMARY
                p_val.space_after = Pt(8)

                p_lbl = mtf.add_paragraph()
                p_lbl.text = m.label
                p_lbl.font.size = Pt(13)
                p_lbl.font.color.rgb = C_MUTED

        # ------------------- 5. COMPARISON / TWO PILLARS -------------------
        elif s.layout_type == "comparison":
            items = s.cards[:2] if len(s.cards) >= 2 else [
                CardItem(title="Technical Alignment", description="Mechanistic interpretability & continuous stress-testing"),
                CardItem(title="Institutional Governance", description="Regulatory compliance & provenance standards")
            ]
            c_w = Inches(5.6)
            c_h = Inches(3.8)
            gap = Inches(0.5)

            for i, item in enumerate(items[:2]):
                cx = Inches(0.8) + i * (c_w + gap)
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), c_w, c_h)
                card.fill.solid()
                card.fill.fore_color.rgb = C_CARD_BG
                card.line.color.rgb = C_CARD_BORDER

                ctb = slide.shapes.add_textbox(cx + Inches(0.4), Inches(2.1), c_w - Inches(0.8), c_h - Inches(0.6))
                ctf = ctb.text_frame
                ctf.word_wrap = True

                cp1 = ctf.paragraphs[0]
                cp1.text = item.title
                cp1.font.size = Pt(20)
                cp1.font.bold = True
                cp1.font.color.rgb = C_PRIMARY
                cp1.space_after = Pt(12)

                cp2 = ctf.add_paragraph()
                cp2.text = item.description
                cp2.font.size = Pt(15)
                cp2.font.color.rgb = C_MUTED

        # Bottom Takeaway Callout Banner for content slides
        if s.key_takeaway:
            banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.0), Inches(11.7), Inches(0.85))
            banner.fill.solid()
            banner.fill.fore_color.rgb = C_CYAN_LIGHT
            banner.line.color.rgb = C_CYAN

            btb = slide.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.3), Inches(0.75))
            btf = btb.text_frame
            btf.word_wrap = True
            bp = btf.paragraphs[0]
            bp.text = f"💡 {s.key_takeaway}"
            bp.font.size = Pt(13)
            bp.font.bold = True
            bp.font.color.rgb = C_CYAN

    prs.save(output_path)
    return output_path
