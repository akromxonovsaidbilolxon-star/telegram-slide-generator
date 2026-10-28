import os
import io
import urllib.parse
import requests
from PIL import Image
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pydantic import BaseModel, Field

# 1. Pydantic Structured Data Schema
class MetricItem(BaseModel):
    value: str = Field(description="Short big stat, e.g. '85%', '10x', 'Zero', '$4.2B'")
    label: str = Field(description="Clear explanatory caption under the metric")

class CardItem(BaseModel):
    title: str = Field(description="Pillar or concept title")
    description: str = Field(description="Detailed 2-sentence explanation")

class SlideStructure(BaseModel):
    slide_number: int
    layout_type: str = Field(
        description="Must be one of: 'title', 'cards', 'split_image', 'metrics', 'comparison', 'closing'"
    )
    eyebrow: str = Field(description="Category tag in uppercase, e.g., 'SYSTEM ARCHITECTURE', 'MARKET IMPACT'")
    title: str = Field(description="Punchy executive headline, max 45 characters")
    subtitle: str | None = None
    bullet_points: list[str] = Field(default=[], description="3-4 comprehensive bullet points")
    cards: list[CardItem] = Field(default=[], description="2 or 3 items for 'cards' or 'comparison' layouts")
    metrics: list[MetricItem] = Field(default=[], description="Exactly 3 metrics for 'metrics' layout")
    key_takeaway: str | None = Field(default=None, description="One strong synthesis sentence")
    image_query: str | None = Field(default=None, description="2 simple English keywords for photo query, e.g. 'logistics highway'")

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    slides: list[SlideStructure]

# 2. Reliable Image Downloader
def fetch_photo(query: str | None) -> io.BytesIO | None:
    if not query:
        return None
    try:
        # Take the first two words to avoid complex URL encoding issues
        clean_keyword = urllib.parse.quote_plus("-".join(query.strip().split()[:2]).lower())
        url = f"https://picsum.photos/seed/{clean_keyword}/800/600"

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }

        resp = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
        if resp.status_code == 200 and len(resp.content) > 3000:
            img_io = io.BytesIO(resp.content)
            img = Image.open(img_io)
            img.verify()
            img_io.seek(0)
            return img_io
        else:
            print(f"[Warning] Failed to fetch image, status code: {resp.status_code}")
    except Exception as e:
        print(f"[Error] Image fetch failed: {e}")
    return None

# 3. Gemini Generation Call
def generate_presentation_json(topic: str, slide_count: int = 6, language: str = "uz") -> MasterpiecePresentation:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_map = {
        "uz": "O'zbek tili (barcha matnlar, xulosalar va punktlar 100% ravon o'zbek tilida yozilsin)",
        "ru": "Русский язык (весь контент, заголовки и выводы строго на русском языке)",
        "en": "English language (executive consulting style, concise and professional)"
    }
    target_language = lang_map.get(language, "O'zbek tili")

    system_instruction = (
        "You are an executive presentation designer at a top tier strategy consultancy. "
        "Create an in-depth, structured presentation deck matching the requested slide count and language. "
        "Ensure layout variety across slides:\n"
        "- Slide 1: 'title'\n"
        "- Slide 2: 'cards' (3 core pillars or dimensions)\n"
        "- Slide 3: 'split_image' (detailed bullet points paired with a relevant photo)\n"
        "- Slide 4: 'metrics' (3 quantitative metrics or growth markers)\n"
        "- Slide 5: 'comparison' (2 structured pillar cards)\n"
        "- Slide 6: 'closing' (strategic roadmap and concluding insight)\n"
        "Avoid vague or sparse text. Provide rich, insightful information adhering strictly to the JSON schema."
    )

    prompt = f"Topic: {topic}\nSlide Count: {slide_count}\nLanguage: {target_language}"

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

# 4. Presentation Assembly Engine
def create_pptx_file(plan: MasterpiecePresentation, output_path: str = "output.pptx") -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Color Palette
    C_DARK_BG = RGBColor(9, 13, 22)         # Obsidian slate
    C_CARD_BG = RGBColor(248, 250, 252)     # Off-white canvas
    C_CARD_BORDER = RGBColor(226, 232, 240) # Slate border
    C_PRIMARY = RGBColor(15, 23, 42)        # Charcoal title
    C_CYAN = RGBColor(2, 132, 199)          # Electric teal accent
    C_CYAN_LIGHT = RGBColor(224, 242, 254)  # Cyan banner fill
    C_MUTED = RGBColor(100, 116, 139)       # Body text slate
    C_WHITE = RGBColor(248, 250, 252)

    blank_layout = prs.slide_layouts[6]

    for s in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # ---------------- 1. TITLE SLIDE (DARK) ----------------
        if s.layout_type == "title":
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = C_DARK_BG
            bg.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(11.0), Inches(3.8))
            tf = tb.text_frame
            tf.word_wrap = True

            pe = tf.paragraphs[0]
            pe.text = s.eyebrow.upper()
            pe.font.size = Pt(13)
            pe.font.bold = True
            pe.font.color.rgb = C_CYAN
            pe.space_after = Pt(16)

            pt = tf.add_paragraph()
            pt.text = s.title
            pt.font.size = Pt(46)
            pt.font.bold = True
            pt.font.color.rgb = C_WHITE
            pt.space_after = Pt(18)

            if s.subtitle:
                ps = tf.add_paragraph()
                ps.text = s.subtitle
                ps.font.size = Pt(20)
                ps.font.color.rgb = C_MUTED

            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(5.8), Inches(1.2), Inches(0.06))
            line.fill.solid()
            line.fill.fore_color.rgb = C_CYAN
            line.line.fill.background()
            continue

        # ---------------- 6. CLOSING SLIDE (DARK) ----------------
        if s.layout_type == "closing":
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = C_DARK_BG
            bg.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10.3), Inches(3.5))
            tf = tb.text_frame
            tf.word_wrap = True

            pe = tf.paragraphs[0]
            pe.text = s.eyebrow.upper()
            pe.font.size = Pt(13)
            pe.font.bold = True
            pe.font.color.rgb = C_CYAN
            pe.alignment = PP_ALIGN.CENTER
            pe.space_after = Pt(14)

            pt = tf.add_paragraph()
            pt.text = s.title
            pt.font.size = Pt(38)
            pt.font.bold = True
            pt.font.color.rgb = C_WHITE
            pt.alignment = PP_ALIGN.CENTER
            pt.space_after = Pt(22)

            if s.key_takeaway or s.subtitle:
                ps = tf.add_paragraph()
                ps.text = s.key_takeaway or s.subtitle
                ps.font.size = Pt(18)
                ps.font.color.rgb = C_MUTED
                ps.alignment = PP_ALIGN.CENTER
            continue

        # ---------------- CONTENT SLIDE HEADER ----------------
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

        # ---------------- 2. CARDS LAYOUT ----------------
        if s.layout_type == "cards":
            items = s.cards if s.cards else [CardItem(title=f"Aspekt {i+1}", description=b) for i, b in enumerate(s.bullet_points[:3])]
            card_w = Inches(3.64)
            card_h = Inches(3.8)
            gap = Inches(0.39)

            for i, card_data in enumerate(items[:3]):
                cx = Inches(0.8) + i * (card_w + gap)
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), card_w, card_h)
                card.fill.solid()
                card.fill.fore_color.rgb = C_CARD_BG
                card.line.color.rgb = C_CARD_BORDER

                # Small indicator badge
                badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.3), Inches(2.1), Inches(0.4), Inches(0.4))
                badge.fill.solid()
                badge.fill.fore_color.rgb = C_CYAN_LIGHT
                badge.line.fill.background()

                ctb = slide.shapes.add_textbox(cx + Inches(0.3), Inches(2.7), card_w - Inches(0.6), card_h - Inches(1.1))
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
                cp2.font.size = Pt(13)
                cp2.font.color.rgb = C_MUTED

        # ---------------- 3. SPLIT IMAGE LAYOUT ----------------
        elif s.layout_type == "split_image":
            # Left side bullet points
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.5), Inches(3.9))
            tf = tb.text_frame
            tf.word_wrap = True

            for idx, pt_text in enumerate(s.bullet_points[:4]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = f"•  {pt_text}"
                p.font.size = Pt(14)
                p.font.color.rgb = C_PRIMARY
                p.space_after = Pt(14)

            # Right side image embedding
            img_stream = fetch_photo(s.image_query or s.title)
            image_inserted = False

            if img_stream:
                try:
                    slide.shapes.add_picture(
                        img_stream,
                        Inches(7.6),
                        Inches(1.8),
                        width=Inches(4.9),
                        height=Inches(3.8)
                    )
                    image_inserted = True
                except Exception as err:
                    print(f"Failed to place image: {err}")

            if not image_inserted:
                # Styled fallback card with accent border
                fallback_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.6), Inches(1.8), Inches(4.9), Inches(3.8))
                fallback_card.fill.solid()
                fallback_card.fill.fore_color.rgb = C_CARD_BG
                fallback_card.line.color.rgb = C_CYAN

                ftb = slide.shapes.add_textbox(Inches(7.8), Inches(3.0), Inches(4.5), Inches(1.5))
                ftf = ftb.text_frame
                ftf.word_wrap = True
                fp = ftf.paragraphs[0]
                fp.text = f"📊 {s.title}"
                fp.font.size = Pt(18)
                fp.font.bold = True
                fp.font.color.rgb = C_PRIMARY
                fp.alignment = PP_ALIGN.CENTER

        # ---------------- 4. METRICS / STATS LAYOUT ----------------
        elif s.layout_type == "metrics":
            m_items = s.metrics if s.metrics else [
                MetricItem(value="85%", label="Samaradorlik o'sishi"),
                MetricItem(value="10x", label="Jarayon tezlashishi"),
                MetricItem(value="Zero", label="Operatsion xatolar")
            ]
            m_w = Inches(3.64)
            m_h = Inches(2.3)
            gap = Inches(0.39)

            for i, m in enumerate(m_items[:3]):
                mx = Inches(0.8) + i * (m_w + gap)
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.9), m_w, m_h)
                box.fill.solid()
                box.fill.fore_color.rgb = C_CARD_BG
                box.line.color.rgb = C_CARD_BORDER

                mtb = slide.shapes.add_textbox(mx + Inches(0.3), Inches(2.1), m_w - Inches(0.6), Inches(1.8))
                mtf = mtb.text_frame
                mtf.word_wrap = True

                p_val = mtf.paragraphs[0]
                p_val.text = m.value
                p_val.font.size = Pt(40)
                p_val.font.bold = True
                p_val.font.color.rgb = C_CYAN if i % 2 == 0 else C_PRIMARY
                p_val.space_after = Pt(8)

                p_lbl = mtf.add_paragraph()
                p_lbl.text = m.label
                p_lbl.font.size = Pt(13)
                p_lbl.font.color.rgb = C_MUTED

        # ---------------- 5. COMPARISON LAYOUT ----------------
        elif s.layout_type == "comparison":
            items = s.cards[:2] if len(s.cards) >= 2 else [
                CardItem(title="Birinchi Yo'nalish", description="Tizimli optimallashtirish va avtomatlashtirish usullari."),
                CardItem(title="Ikkinchi Yo'nalish", description="Strategik tahlil va inson nazorati integratsiyasi.")
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
                cp2.font.size = Pt(14)
                cp2.font.color.rgb = C_MUTED

        # Bottom Takeaway Callout Banner
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
