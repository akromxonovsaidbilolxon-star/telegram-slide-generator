import os
import io
import urllib.parse
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

# ----------------- THEMES CONFIGURATION -----------------
THEMES = {
    "midnight": {
        "name": "Midnight Executive",
        "dark_bg": RGBColor(9, 13, 22),
        "light_bg": RGBColor(255, 255, 255),
        "card_bg": RGBColor(248, 250, 252),
        "card_border": RGBColor(226, 232, 240),
        "primary": RGBColor(15, 23, 42),
        "accent": RGBColor(2, 132, 199),         # Cyan 600
        "accent_light": RGBColor(224, 242, 254), # Cyan 100
        "muted": RGBColor(100, 116, 139),
        "white": RGBColor(248, 250, 252)
    },
    "emerald": {
        "name": "Emerald Growth",
        "dark_bg": RGBColor(6, 78, 59),
        "light_bg": RGBColor(255, 255, 255),
        "card_bg": RGBColor(240, 253, 244),
        "card_border": RGBColor(187, 247, 208),
        "primary": RGBColor(19, 78, 74),
        "accent": RGBColor(5, 150, 105),         # Emerald 600
        "accent_light": RGBColor(209, 250, 229), # Emerald 100
        "muted": RGBColor(75, 85, 99),
        "white": RGBColor(255, 255, 255)
    },
    "minimal": {
        "name": "Modern Minimalist",
        "dark_bg": RGBColor(24, 24, 27),
        "light_bg": RGBColor(255, 255, 255),
        "card_bg": RGBColor(244, 244, 245),
        "card_border": RGBColor(228, 228, 231),
        "primary": RGBColor(24, 24, 27),
        "accent": RGBColor(217, 119, 6),          # Amber 600
        "accent_light": RGBColor(254, 243, 199),  # Amber 100
        "muted": RGBColor(113, 113, 122),
        "white": RGBColor(250, 250, 250)
    }
}

# ----------------- PYDANTIC SCHEMAS -----------------
class MetricItem(BaseModel):
    value: str = Field(description="Real-world metric, e.g. '$4.4T', '18%', '70%', '10x'")
    label: str = Field(description="Contextual explanation for the stat")

class CardItem(BaseModel):
    title: str = Field(description="Pillar name or core concept")
    description: str = Field(description="Two analytical, factual sentences")

class SlideStructure(BaseModel):
    slide_number: int
    layout_type: str = Field(description="'title', 'cards', 'split_image', 'metrics', 'comparison', 'closing'")
    eyebrow: str = Field(description="Categorical uppercase tag (e.g. 'OPERATIONAL STRATEGY')")
    title: str = Field(description="Punchy executive headline, max 45 chars")
    subtitle: str | None = None
    bullet_points: list[str] = Field(default=[], description="Bullet points starting with bold concepts (e.g. 'Autonomous Execution: ...')")
    cards: list[CardItem] = Field(default=[], description="2 or 3 distinct conceptual cards")
    metrics: list[MetricItem] = Field(default=[], description="Exactly 3 metrics for 'metrics' layout")
    key_takeaway: str | None = Field(default=None, description="One crisp takeaway sentence")
    image_query: str | None = Field(default=None, description="2-3 specific English keywords for Google Images")

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    theme: str
    slides: list[SlideStructure]

# ----------------- PIXABAY FALLBACK -----------------
def fetch_pixabay_photo(query: str | None) -> io.BytesIO | None:
    api_key = os.environ.get("PIXABAY_API_KEY")
    if not api_key or not query:
        return None

    clean_query = " ".join(query.strip().replace("-", " ").split()[:3])
    url = "https://pixabay.com/api/"
    params = {
        "key": api_key,
        "q": clean_query,
        "image_type": "photo",
        "orientation": "horizontal",
        "safesearch": "true",
        "per_page": 3
    }
    try:
        resp = requests.get(url, params=params, timeout=8)
        if resp.status_code == 200:
            hits = resp.json().get("hits", [])
            if hits:
                img_url = hits[0].get("webformatURL") or hits[0].get("largeImageURL")
                img_resp = requests.get(img_url, timeout=8)
                if img_resp.status_code == 200 and len(img_resp.content) > 3000:
                    img_io = io.BytesIO(img_resp.content)
                    img = Image.open(img_io)
                    img.verify()
                    img_io.seek(0)
                    return img_io
    except Exception as e:
        print(f"[Pixabay Fallback Error]: {e}")
    return None

# ----------------- GOOGLE CUSTOM SEARCH (IMAGE) -----------------
def fetch_google_image(query: str | None) -> io.BytesIO | None:
    api_key = os.environ.get("GOOGLE_SEARCH_API_KEY")
    cx = os.environ.get("GOOGLE_CSE_ID")

    if not api_key or not cx or not query:
        return fetch_pixabay_photo(query)

    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": api_key,
        "cx": cx,
        "q": query.strip(),
        "searchType": "image",
        "num": 3,
        "imgSize": "large",
        "imgType": "photo",
        "safe": "active"
    }

    try:
        resp = requests.get(url, params=params, timeout=10)
        if resp.status_code == 200:
            items = resp.json().get("items", [])
            for item in items:
                img_url = item.get("link")
                if not img_url:
                    continue
                try:
                    headers = {
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                    }
                    img_resp = requests.get(img_url, headers=headers, timeout=8)
                    if img_resp.status_code == 200 and len(img_resp.content) > 4000:
                        img_io = io.BytesIO(img_resp.content)
                        img = Image.open(img_io)
                        img.verify()
                        img_io.seek(0)
                        return img_io
                except Exception:
                    continue
        else:
            print(f"[Google Search Warning]: HTTP {resp.status_code}")
    except Exception as e:
        print(f"[Google Search Error]: {e}")

    # Если Google не вернул результат, пробуем Pixabay
    return fetch_pixabay_photo(query)

# ----------------- GEMINI + SEARCH GROUNDING -----------------
def generate_presentation_json(topic: str, slide_count: int = 6, language: str = "uz", theme: str = "midnight") -> MasterpiecePresentation:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_map = {
        "uz": "O'zbek tili (barcha faktlar, raqamlar, tushuntirishlar va xulosalar 100% ravon o'zbek tilida yozilsin)",
        "ru": "Русский язык (реальные даты, статистика и выводы строго на русском языке)",
        "en": "English language (executive consulting grade, factual metrics and verified dates)"
    }
    target_language = lang_map.get(language, "O'zbek tili")

    research_prompt = (
        f"Perform Google Search to find current facts, verified statistics, dates, and industry data "
        f"regarding the topic: '{topic}'. Focus on real reports, market valuations, and concrete figures."
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
        print(f"[Search Grounding Warning]: {e}")
        research_context = "Use internal up-to-date factual knowledge."

    system_instruction = (
        "You are an executive principal presentation designer at McKinsey & Company. "
        "Create an in-depth, structured presentation deck using the provided research and real facts. "
        "Enforce layout diversity:\n"
        "- Slide 1: 'title'\n"
        "- Slide 2: 'cards' (3 distinct pillars with clear descriptions)\n"
        "- Slide 3: 'split_image' (rich bullet points starting with bold capitalized concepts, paired with a photo query)\n"
        "- Slide 4: 'metrics' (3 real quantitative numbers/percentages based on research)\n"
        "- Slide 5: 'comparison' (2 structural comparative pillars)\n"
        "- Slide 6: 'closing' (strategic roadmap and synthesis)\n"
        "For 'image_query', output 2-3 specific English keywords suitable for Google Images search."
    )

    synthesis_prompt = (
        f"Topic: {topic}\n"
        f"Desired Slide Count: {slide_count}\n"
        f"Theme: {theme}\n"
        f"Target Output Language: {target_language}\n\n"
        f"Research and Facts to Ground the Deck:\n{research_context[:3000]}"
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
def create_pptx_file(plan: MasterpiecePresentation, output_path: str = "output.pptx") -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    theme = THEMES.get(plan.theme, THEMES["midnight"])
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

        # 6. CLOSING SLIDE
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

            # Вызываем Google Custom Search с Pixabay fallback
            img_stream = fetch_google_image(s.image_query or s.title)
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

        # 4. METRICS / STATS LAYOUT
        elif s.layout_type == "metrics":
            m_items = s.metrics if s.metrics else [
                MetricItem(value="85%", label="Samaradorlik ko'rsatkichi"),
                MetricItem(value="10x", label="Jarayon optimallashuvi"),
                MetricItem(value="Zero", label="Operatsion uzilishlar")
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
                CardItem(title="Klassik Yondashuv", description="An'anaviy operatsion modellar va cheklovlar."),
                CardItem(title="Intellektual Yondashuv", description="To'liq avtomatlashtirilgan o'z-o'zini tahlil qiluvchi ekotizim.")
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

    draw.text((70, 70), "EXECUTIVE SLIDE DECK (SEARCH GROUNDED)", fill=(2, 132, 199), font=font_tag)

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
