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

# ----------------- DYNAMIC COLOR ENGINE -----------------
# Generates distinct aesthetic palettes so every user deck looks unique
PALETTES = [
    {
        "dark_bg": RGBColor(10, 15, 29),
        "card_bg": RGBColor(248, 250, 252),
        "card_border": RGBColor(226, 232, 240),
        "primary": RGBColor(15, 23, 42),
        "accent": RGBColor(2, 132, 199),         # Sapphire Cyan
        "accent_light": RGBColor(224, 242, 254),
        "muted": RGBColor(100, 116, 139),
        "white": RGBColor(248, 250, 252)
    },
    {
        "dark_bg": RGBColor(15, 23, 42),
        "card_bg": RGBColor(241, 245, 249),
        "card_border": RGBColor(203, 213, 225),
        "primary": RGBColor(30, 41, 59),
        "accent": RGBColor(99, 102, 241),        # Indigo Violet
        "accent_light": RGBColor(238, 242, 255),
        "muted": RGBColor(100, 116, 139),
        "white": RGBColor(255, 255, 255)
    },
    {
        "dark_bg": RGBColor(6, 78, 59),
        "card_bg": RGBColor(240, 253, 244),
        "card_border": RGBColor(187, 247, 208),
        "primary": RGBColor(6, 78, 59),
        "accent": RGBColor(16, 185, 129),        # Emerald Mint
        "accent_light": RGBColor(209, 250, 229),
        "muted": RGBColor(75, 85, 99),
        "white": RGBColor(255, 255, 255)
    },
    {
        "dark_bg": RGBColor(24, 24, 27),
        "card_bg": RGBColor(250, 250, 250),
        "card_border": RGBColor(228, 228, 231),
        "primary": RGBColor(24, 24, 27),
        "accent": RGBColor(217, 119, 6),          # Executive Amber
        "accent_light": RGBColor(254, 243, 199),
        "muted": RGBColor(113, 113, 122),
        "white": RGBColor(255, 255, 255)
    }
]

# ----------------- PYDANTIC SCHEMAS -----------------
class MetricItem(BaseModel):
    value: str = Field(description="Concrete number/stat, e.g., '69 ta', '20 km³', '85%'")
    label: str = Field(description="Contextual description under the metric")

class CardItem(BaseModel):
    title: str = Field(description="Pillar name or core concept")
    description: str = Field(description="Two clear sentences explaining the concept")

class SlideStructure(BaseModel):
    slide_number: int
    layout_type: str = Field(description="Must be one of: 'title', 'cards', 'split_image', 'metrics', 'comparison', 'closing'")
    eyebrow: str = Field(description="Categorical uppercase tag")
    title: str = Field(description="Punchy executive headline, max 45 chars")
    subtitle: str | None = None
    bullet_points: list[str] = Field(default=[], description="Bullet points starting with bold capitalized concepts")
    cards: list[CardItem] = Field(default=[], description="Cards for 'cards' or 'comparison' layouts")
    metrics: list[MetricItem] = Field(default=[], description="3 metrics for 'metrics' layout")
    key_takeaway: str | None = Field(default=None, description="One crisp takeaway sentence")
    image_query_en: str | None = Field(default=None, description="2-3 simple generic English keywords for stock photo search. Example: 'water dam aerial', 'hydroelectric turbine', 'modern logistics warehouse'")

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    total_slides: int
    slides: list[SlideStructure]

# ----------------- IMAGE RETRIEVAL WITH FALLBACKS -----------------
def fetch_photo(query_en: str | None, user_custom_image_bytes: bytes | None = None) -> io.BytesIO | None:
    # 1. Prioritize user-uploaded picture if provided
    if user_custom_image_bytes:
        try:
            return io.BytesIO(user_custom_image_bytes)
        except Exception:
            pass

    if not query_en:
        return None

    # Clean English query
    clean_q = " ".join(query_en.strip().replace("-", " ").split()[:3])
    
    # 2. Try Google Custom Search first
    api_key = os.environ.get("GOOGLE_SEARCH_API_KEY")
    cx = os.environ.get("GOOGLE_CSE_ID")
    if api_key and cx:
        url = "https://www.googleapis.com/customsearch/v1"
        params = {
            "key": api_key,
            "cx": cx,
            "q": clean_q,
            "searchType": "image",
            "num": 3,
            "imgSize": "large",
            "safe": "active"
        }
        try:
            resp = requests.get(url, params=params, timeout=8)
            if resp.status_code == 200:
                items = resp.json().get("items", [])
                for item in items:
                    img_url = item.get("link")
                    if img_url:
                        try:
                            headers = {"User-Agent": "Mozilla/5.0"}
                            img_resp = requests.get(img_url, headers=headers, timeout=6)
                            if img_resp.status_code == 200 and len(img_resp.content) > 3000:
                                img_io = io.BytesIO(img_resp.content)
                                img = Image.open(img_io)
                                img.verify()
                                img_io.seek(0)
                                return img_io
                        except Exception:
                            continue
        except Exception as e:
            print(f"[Google Image Error]: {e}")

    # 3. Try Pixabay API Fallback
    pixabay_key = os.environ.get("PIXABAY_API_KEY")
    if pixabay_key:
        try:
            p_url = "https://pixabay.com/api/"
            params = {
                "key": pixabay_key,
                "q": clean_q,
                "image_type": "photo",
                "orientation": "horizontal",
                "per_page": 3
            }
            resp = requests.get(p_url, params=params, timeout=6)
            if resp.status_code == 200:
                hits = resp.json().get("hits", [])
                if hits:
                    img_url = hits[0].get("webformatURL") or hits[0].get("largeImageURL")
                    img_resp = requests.get(img_url, timeout=6)
                    if img_resp.status_code == 200:
                        return io.BytesIO(img_resp.content)
        except Exception as e:
            print(f"[Pixabay Error]: {e}")

    # 4. Ultimate Unsplash/Picsum Direct Photo Redirect
    try:
        url_seed = urllib.parse.quote_plus(clean_q)
        direct_resp = requests.get(f"https://picsum.photos/seed/{url_seed}/800/600", timeout=6)
        if direct_resp.status_code == 200 and len(direct_resp.content) > 3000:
            return io.BytesIO(direct_resp.content)
    except Exception:
        pass

    return None

# ----------------- GEMINI GENERATION -----------------
def generate_presentation_json(
    topic: str,
    slide_count: int = 6,
    language: str = "uz",
    user_notes: str | None = None
) -> MasterpiecePresentation:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_map = {
        "uz": "O'zbek tili (barcha matnlar, faktlar va xulosalar 100% ravon o'zbek tilida bo'lsin)",
        "ru": "Русский язык (весь текст строго на русском языке)",
        "en": "English language (executive consulting grade)"
    }
    target_language = lang_map.get(language, "O'zbek tili")

    # Research Step using Google Search
    research_prompt = (
        f"Search Google for accurate, up-to-date facts, real metrics, laws/decrees, and data on: '{topic}'. "
        f"User custom details: {user_notes or 'None'}"
    )
    try:
        search_res = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=research_prompt,
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
                temperature=0.2
            )
        )
        research_context = search_res.text if search_res and search_res.text else ""
    except Exception as e:
        print(f"[Search Warning]: {e}")
        research_context = ""

    # Synthesis Step with Strict Length Enforcement
    system_instruction = (
        f"You are a Senior Presentation Architect at McKinsey. "
        f"You MUST generate EXACTLY {slide_count} slides in the 'slides' array. Not 6, not 8, but exactly {slide_count}. "
        f"- Slide 1 MUST be layout_type 'title'.\n"
        f"- Slide {slide_count} MUST be layout_type 'closing'.\n"
        f"- All intermediary slides (slides 2 through {slide_count - 1}) must alternate dynamically among: "
        f"'cards', 'split_image', 'metrics', and 'comparison'.\n"
        f"CRITICAL: 'image_query_en' MUST ALWAYS BE 2-3 ENGLISH WORDS describing a generic stock photo "
        f"(e.g., 'water reservoir dam', 'solar panels field', 'cargo ship harbor'). NEVER use non-English terms for image_query_en."
    )

    synthesis_prompt = (
        f"Topic: {topic}\n"
        f"Required Exact Slide Count: {slide_count}\n"
        f"Language: {target_language}\n"
        f"User Custom Instructions: {user_notes or 'Standard high-level briefing'}\n\n"
        f"Verified Facts Context:\n{research_context[:3500]}"
    )

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=synthesis_prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            response_mime_type="application/json",
            response_schema=MasterpiecePresentation,
            temperature=0.2
        ),
    )
    return response.parsed

# ----------------- PPTX BUILDER -----------------
def create_pptx_file(
    plan: MasterpiecePresentation,
    output_path: str = "output.pptx",
    user_custom_image_bytes: bytes | None = None
) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Randomly select one of the luxury palettes for visual variety per user/deck
    theme = random.choice(PALETTES)
    blank_layout = prs.slide_layouts[6]

    for s in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # 1. TITLE SLIDE
        if s.layout_type == "title":
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = theme["dark_bg"]
            bg.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(11.0), Inches(3.8))
            tf = tb.text_frame
            tf.word_wrap = True

            pe = tf.paragraphs[0]
            pe.text = s.eyebrow.upper()
            pe.font.size = Pt(13)
            pe.font.bold = True
            pe.font.color.rgb = theme["accent"]
            pe.space_after = Pt(16)

            pt = tf.add_paragraph()
            pt.text = s.title
            pt.font.size = Pt(44)
            pt.font.bold = True
            pt.font.color.rgb = theme["white"]
            pt.space_after = Pt(18)

            if s.subtitle:
                ps = tf.add_paragraph()
                ps.text = s.subtitle
                ps.font.size = Pt(20)
                ps.font.color.rgb = theme["muted"]

            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(5.8), Inches(1.2), Inches(0.06))
            line.fill.solid()
            line.fill.fore_color.rgb = theme["accent"]
            line.line.fill.background()
            continue

        # CLOSING SLIDE
        if s.layout_type == "closing":
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = theme["dark_bg"]
            bg.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.5), Inches(2.0), Inches(10.3), Inches(3.5))
            tf = tb.text_frame
            tf.word_wrap = True

            pe = tf.paragraphs[0]
            pe.text = s.eyebrow.upper()
            pe.font.size = Pt(13)
            pe.font.bold = True
            pe.font.color.rgb = theme["accent"]
            pe.alignment = PP_ALIGN.CENTER
            pe.space_after = Pt(14)

            pt = tf.add_paragraph()
            pt.text = s.title
            pt.font.size = Pt(38)
            pt.font.bold = True
            pt.font.color.rgb = theme["white"]
            pt.alignment = PP_ALIGN.CENTER
            pt.space_after = Pt(22)

            if s.key_takeaway or s.subtitle:
                ps = tf.add_paragraph()
                ps.text = s.key_takeaway or s.subtitle
                ps.font.size = Pt(18)
                ps.font.color.rgb = theme["muted"]
                ps.alignment = PP_ALIGN.CENTER
            continue

        # CONTENT SLIDE HEADER
        ht_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(1.2))
        htf = ht_box.text_frame
        htf.word_wrap = True

        pe = htf.paragraphs[0]
        pe.text = s.eyebrow.upper()
        pe.font.size = Pt(11)
        pe.font.bold = True
        pe.font.color.rgb = theme["accent"]

        pt = htf.add_paragraph()
        pt.text = s.title
        pt.font.size = Pt(28)
        pt.font.bold = True
        pt.font.color.rgb = theme["primary"]

        # 2. CARDS LAYOUT
        if s.layout_type == "cards":
            items = s.cards if s.cards else [CardItem(title=f"Aspekt {i+1}", description=b) for i, b in enumerate(s.bullet_points[:3])]
            card_w = Inches(3.64)
            card_h = Inches(3.8)
            gap = Inches(0.39)

            for i, card_data in enumerate(items[:3]):
                cx = Inches(0.8) + i * (card_w + gap)
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), card_w, card_h)
                card.fill.solid()
                card.fill.fore_color.rgb = theme["card_bg"]
                card.line.color.rgb = theme["card_border"]

                badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx + Inches(0.3), Inches(2.1), Inches(0.35), Inches(0.35))
                badge.fill.solid()
                badge.fill.fore_color.rgb = theme["accent_light"]
                badge.line.fill.background()

                ctb = slide.shapes.add_textbox(cx + Inches(0.3), Inches(2.6), card_w - Inches(0.6), card_h - Inches(1.0))
                ctf = ctb.text_frame
                ctf.word_wrap = True

                cp1 = ctf.paragraphs[0]
                cp1.text = card_data.title
                cp1.font.size = Pt(18)
                cp1.font.bold = True
                cp1.font.color.rgb = theme["primary"]
                cp1.space_after = Pt(10)

                cp2 = ctf.add_paragraph()
                cp2.text = card_data.description
                cp2.font.size = Pt(13)
                cp2.font.color.rgb = theme["muted"]

        # 3. SPLIT IMAGE LAYOUT
        elif s.layout_type == "split_image":
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.5), Inches(3.9))
            tf = tb.text_frame
            tf.word_wrap = True

            for idx, pt_text in enumerate(s.bullet_points[:4]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = f"•  {pt_text}"
                p.font.size = Pt(14)
                p.font.color.rgb = theme["primary"]
                p.space_after = Pt(14)

            # Fetch image using English search query or user's custom photo
            img_stream = fetch_photo(s.image_query_en, user_custom_image_bytes)
            image_inserted = False

            if img_stream:
                try:
                    slide.shapes.add_picture(img_stream, Inches(7.6), Inches(1.8), width=Inches(4.9), height=Inches(3.8))
                    image_inserted = True
                except Exception as e:
                    print(f"Failed to place image: {e}")

            if not image_inserted:
                fallback_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.6), Inches(1.8), Inches(4.9), Inches(3.8))
                fallback_card.fill.solid()
                fallback_card.fill.fore_color.rgb = theme["card_bg"]
                fallback_card.line.color.rgb = theme["accent"]

                ftb = slide.shapes.add_textbox(Inches(7.8), Inches(3.0), Inches(4.5), Inches(1.5))
                ftf = ftb.text_frame
                ftf.word_wrap = True
                fp = ftf.paragraphs[0]
                fp.text = f"📊 {s.title}"
                fp.font.size = Pt(18)
                fp.font.bold = True
                fp.font.color.rgb = theme["primary"]
                fp.alignment = PP_ALIGN.CENTER

        # 4. METRICS LAYOUT
        elif s.layout_type == "metrics":
            m_items = s.metrics if s.metrics else [
                MetricItem(value="85%", label="Samaradorlik"),
                MetricItem(value="10x", label="Tezlashuv"),
                MetricItem(value="Zero", label="Operatsion xatolar")
            ]
            m_w = Inches(3.64)
            m_h = Inches(2.3)
            gap = Inches(0.39)

            for i, m in enumerate(m_items[:3]):
                mx = Inches(0.8) + i * (m_w + gap)
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.9), m_w, m_h)
                box.fill.solid()
                box.fill.fore_color.rgb = theme["card_bg"]
                box.line.color.rgb = theme["card_border"]

                mtb = slide.shapes.add_textbox(mx + Inches(0.3), Inches(2.1), m_w - Inches(0.6), Inches(1.8))
                mtf = mtb.text_frame
                mtf.word_wrap = True

                p_val = mtf.paragraphs[0]
                p_val.text = m.value
                p_val.font.size = Pt(40)
                p_val.font.bold = True
                p_val.font.color.rgb = theme["accent"] if i % 2 == 0 else theme["primary"]
                p_val.space_after = Pt(8)

                p_lbl = mtf.add_paragraph()
                p_lbl.text = m.label
                p_lbl.font.size = Pt(13)
                p_lbl.font.color.rgb = theme["muted"]

        # 5. COMPARISON LAYOUT
        elif s.layout_type == "comparison":
            items = s.cards[:2] if len(s.cards) >= 2 else [
                CardItem(title="Klassik Yondashuv", description="An'anaviy operatsion modellar va ularning cheklovlari."),
                CardItem(title="Zamonaviy Yondashuv", description="Avtomatlashtirilgan raqamli boshqaruv tizimlari.")
            ]
            c_w = Inches(5.6)
            c_h = Inches(3.8)
            gap = Inches(0.5)

            for i, item in enumerate(items[:2]):
                cx = Inches(0.8) + i * (c_w + gap)
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), c_w, c_h)
                card.fill.solid()
                card.fill.fore_color.rgb = theme["card_bg"]
                card.line.color.rgb = theme["card_border"]

                ctb = slide.shapes.add_textbox(cx + Inches(0.4), Inches(2.1), c_w - Inches(0.8), c_h - Inches(0.6))
                ctf = ctb.text_frame
                ctf.word_wrap = True

                cp1 = ctf.paragraphs[0]
                cp1.text = item.title
                cp1.font.size = Pt(20)
                cp1.font.bold = True
                cp1.font.color.rgb = theme["primary"]
                cp1.space_after = Pt(12)

                cp2 = ctf.add_paragraph()
                cp2.text = item.description
                cp2.font.size = Pt(14)
                cp2.font.color.rgb = theme["muted"]

        # BOTTOM TAKEAWAY CALLOUT
        if s.key_takeaway:
            banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.0), Inches(11.7), Inches(0.85))
            banner.fill.solid()
            banner.fill.fore_color.rgb = theme["accent_light"]
            banner.line.color.rgb = theme["accent"]

            btb = slide.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.3), Inches(0.75))
            btf = btb.text_frame
            btf.word_wrap = True
            bp = btf.paragraphs[0]
            bp.text = f"💡 {s.key_takeaway}"
            bp.font.size = Pt(13)
            bp.font.bold = True
            bp.font.color.rgb = theme["accent"]

    prs.save(output_path)
    return output_path

# ----------------- SLIDE PREVIEW IMAGE GENERATOR -----------------
def generate_preview_card(plan: MasterpiecePresentation, output_img_path: str = "preview.png") -> str:
    W, H = 1200, 675
    img = Image.new("RGB", (W, H), color=(9, 13, 22))
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (W, 10)], fill=(2, 132, 199))

    try:
        font_large = ImageFont.truetype("DejaVuSans-Bold.ttf", 44)
        font_sub = ImageFont.truetype("DejaVuSans.ttf", 22)
        font_tag = ImageFont.truetype("DejaVuSans-Bold.ttf", 16)
    except Exception:
        font_large = ImageFont.load_default()
        font_sub = ImageFont.load_default()
        font_tag = ImageFont.load_default()

    draw.text((70, 70), f"EXECUTIVE DECK ({len(plan.slides)} SLIDES)", fill=(2, 132, 199), font=font_tag)

    title_text = plan.topic.upper()
    if len(title_text) > 40:
        title_text = title_text[:37] + "..."
    draw.text((70, 110), title_text, fill=(248, 250, 252), font=font_large)

    card_w, card_h = 330, 360
    y_card, gap = 220, 35

    for i, s in enumerate(plan.slides[:3]):
        x_card = 70 + i * (card_w + gap)
        draw.rounded_rectangle([(x_card, y_card), (x_card + card_w, y_card + card_h)], radius=12, fill=(18, 24, 38), outline=(30, 41, 59), width=2)
        draw.rounded_rectangle([(x_card + 20, y_card + 20), (x_card + 80, y_card + 50)], radius=6, fill=(2, 132, 199))
        draw.text((x_card + 30, y_card + 25), f"0{s.slide_number}", fill=(255, 255, 255), font=font_tag)

        s_title = s.title if len(s.title) < 28 else s.title[:25] + "..."
        draw.text((x_card + 20, y_card + 75), s_title, fill=(248, 250, 252), font=font_tag)
        draw.text((x_card + 20, y_card + 115), f"Layout: {s.layout_type.capitalize()}", fill=(100, 116, 139), font=font_sub)

    img.save(output_img_path)
    return output_img_path
