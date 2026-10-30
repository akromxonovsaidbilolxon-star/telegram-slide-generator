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

# ----------------- THEMES & PALETTES -----------------
THEMES = {
    "obsidian": {
        "name": "Obsidian Cyan",
        "dark_bg": RGBColor(10, 15, 29),
        "card_bg": RGBColor(19, 29, 53),
        "card_border": RGBColor(38, 56, 96),
        "text_main": RGBColor(255, 255, 255),
        "text_muted": RGBColor(165, 180, 205),
        "accent": RGBColor(56, 189, 248),
        "accent_light": RGBColor(12, 74, 110)
    },
    "emerald": {
        "name": "Nordic Emerald",
        "dark_bg": RGBColor(6, 20, 24),
        "card_bg": RGBColor(13, 37, 44),
        "card_border": RGBColor(21, 58, 69),
        "text_main": RGBColor(241, 245, 249),
        "text_muted": RGBColor(148, 175, 185),
        "accent": RGBColor(45, 212, 191),
        "accent_light": RGBColor(19, 78, 74)
    },
    "purple": {
        "name": "Deep Royal Purple",
        "dark_bg": RGBColor(15, 12, 34),
        "card_bg": RGBColor(30, 24, 66),
        "card_border": RGBColor(60, 48, 120),
        "text_main": RGBColor(255, 255, 255),
        "text_muted": RGBColor(185, 175, 225),
        "accent": RGBColor(192, 132, 252),
        "accent_light": RGBColor(88, 28, 135)
    }
}

# ----------------- SCHEMAS -----------------
class MetricBox(BaseModel):
    value: str = Field(description="Display stat, e.g. '1.6B', '88%', '$50B'")
    title: str = Field(description="Short label, max 3 words")
    desc: str = Field(description="Brief 1-sentence explanation")

class PillarCard(BaseModel):
    number: str = Field(description="'01', '02', '03'")
    title: str = Field(description="Headline of pillar")
    desc: str = Field(description="Two clear sentences")

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
    ai_image_prompt: str | None = Field(default=None, description="2-3 simple English keywords for stock photo search")

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    theme_key: str = "obsidian"
    total_slides: int
    slides: list[SlideContent]

class OutlineProposal(BaseModel):
    variant_a_title: str = Field(description="Short title of strategy A, e.g. 'Akademik va Nazariy Tahlil'")
    variant_a_points: list[str] = Field(description="List of slide titles for Option A")
    variant_b_title: str = Field(description="Short title of strategy B, e.g. 'Amaliy va Innovatsion Yondashuv'")
    variant_b_points: list[str] = Field(description="List of slide titles for Option B")

# ----------------- AI BRAINSTORMING HELPER -----------------
def generate_outline_options(topic: str, slide_count: int, audience_type: str, language: str = "uz") -> OutlineProposal:
    client = genai.Client(api_key=os.environ.get("GEMINI_API_KEY"))

    lang_map = {
        "uz": "O'zbek tili",
        "ru": "Русский язык",
        "en": "English"
    }

    prompt = (
        f"Mavzu: '{topic}'. Slaydlar soni: {slide_count}. Auditoriya yo'nalishi: {audience_type}.\n"
        f"Foydalanuvchi taqdimotni qanday boshlash va rejalashtirishni bilmayapti.\n"
        f"Unga aynan {slide_count} ta slayddan iborat ikkita (Variant A va Variant B) "
        f"mantiqiy, chuqur va professional reja taklif qil. "
        f"Har ikkala variantdagi bandlar soni aniq {slide_count} ta bo'lsin. "
        f"Til: {lang_map.get(language, 'O uzbek')}."
    )

    res = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=OutlineProposal,
            temperature=0.3
        )
    )
    return res.parsed

# ----------------- IMAGE RETRIEVAL WITH MULTI-PROVIDER FALLBACK -----------------
def fetch_image_secure(prompt: str | None, custom_bytes: bytes | None = None) -> io.BytesIO | None:
    if custom_bytes:
        try:
            return io.BytesIO(custom_bytes)
        except Exception:
            pass

    if not prompt:
        return None

    clean_prompt = " ".join(prompt.strip().replace("-", " ").split()[:3])

    # 1. Google Custom Search
    api_key = os.environ.get("GOOGLE_SEARCH_API_KEY")
    cx = os.environ.get("GOOGLE_CSE_ID")
    if api_key and cx:
        try:
            url = "https://www.googleapis.com/customsearch/v1"
            params = {
                "key": api_key,
                "cx": cx,
                "q": clean_prompt,
                "searchType": "image",
                "num": 2,
                "imgSize": "large",
                "safe": "active"
            }
            res = requests.get(url, params=params, timeout=6)
            if res.status_code == 200:
                items = res.json().get("items", [])
                for item in items:
                    l = item.get("link")
                    if l:
                        ir = requests.get(l, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
                        if ir.status_code == 200 and len(ir.content) > 4000:
                            bio = io.BytesIO(ir.content)
                            img = Image.open(bio)
                            img.verify()
                            bio.seek(0)
                            return bio
        except Exception:
            pass

    # 2. Pixabay API Fallback
    pix_key = os.environ.get("PIXABAY_API_KEY")
    if pix_key:
        try:
            p_res = requests.get("https://pixabay.com/api/", params={
                "key": pix_key,
                "q": clean_prompt,
                "image_type": "photo",
                "orientation": "horizontal",
                "per_page": 3
            }, timeout=5)
            if p_res.status_code == 200:
                hits = p_res.json().get("hits", [])
                if hits:
                    u = hits[0].get("webformatURL")
                    ir = requests.get(u, timeout=5)
                    if ir.status_code == 200:
                        return io.BytesIO(ir.content)
        except Exception:
            pass

    # 3. Direct Photo Stream Fallback
    try:
        clean_seed = urllib.parse.quote_plus(clean_prompt)
        resp = requests.get(f"https://picsum.photos/seed/{clean_seed}/800/600", timeout=5)
        if resp.status_code == 200 and len(resp.content) > 3000:
            return io.BytesIO(resp.content)
    except Exception:
        pass

    return None

# ----------------- CONTENT GENERATION VIA GEMINI -----------------
def generate_presentation_json(
    topic: str,
    slide_count: int = 10,
    language: str = "uz",
    user_notes: str | None = None,
    theme_key: str = "obsidian"
) -> MasterpiecePresentation:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_map = {
        "uz": "O'zbek tili (barcha matnlar 100% ravon o'zbek tilida)",
        "ru": "Русский язык (весь контент строго на русском)",
        "en": "English language (executive presentation grade)"
    }
    target_language = lang_map.get(language, "O'zbek tili")

    research_context = ""
    try:
        research_prompt = (
            f"Search Google for real verified stats, academic/economic literature, "
            f"dates, and core insights on: '{topic}'. User plan/notes: {user_notes or 'None'}"
        )
        search_res = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=research_prompt,
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())],
                temperature=0.2
            )
        )
        if search_res and search_res.text:
            research_context = search_res.text
    except Exception as e:
        print(f"[Search Warning]: {e}")

    # Detect user outline count
    outline_lines = [l for l in (user_notes or "").split("\n") if any(l.strip().startswith(f"{i}.") for i in range(1, 25))]
    target_count = len(outline_lines) if len(outline_lines) >= 5 else slide_count

    system_instruction = (
        f"You are an executive slide architect. You MUST generate EXACTLY {target_count} slides in the 'slides' array. "
        "Keep text concise, punchy, and readable on mobile devices. "
        "Never make headlines exceed 40 characters. "
        "If the user provided an outline or plan, map each point sequentially into its own slide.\n"
        f"- Slide 1 MUST be 'title'.\n"
        f"- Slide {target_count} MUST be 'closing'.\n"
        "- Intermediary slides must alternate between 'cards_grid', 'stats_row', and 'split_image'.\n"
        "- 'ai_image_prompt' MUST be 2-3 simple English keywords describing a stock photo."
    )

    synthesis_prompt = (
        f"Topic: {topic}\n"
        f"Required Exact Slides Count: {target_count}\n"
        f"Language: {target_language}\n"
        f"Theme Chosen: {theme_key}\n"
        f"User Plan / Outline:\n{user_notes or 'Standard presentation'}\n\n"
        f"Research Context:\n{research_context[:3000]}"
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
    result = response.parsed
    result.theme_key = theme_key
    return result

# ----------------- POWERPOINT BUILDER -----------------
def create_pptx_file(
    plan: MasterpiecePresentation,
    output_path: str = "output.pptx",
    custom_img_bytes: bytes | None = None
) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    theme = THEMES.get(plan.theme_key, THEMES["obsidian"])
    blank_layout = prs.slide_layouts[6]

    for s in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme["dark_bg"]
        bg.line.fill.background()

        # 1. TITLE SLIDE
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

        # HEADER FOR CONTENT SLIDES
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

        # 2. CARDS GRID
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

        # 3. STATS ROW
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

        # 4. SPLIT IMAGE LAYOUT
        elif s.layout_type == "split_image":
            img_io = fetch_image_secure(s.ai_image_prompt or s.title, custom_img_bytes)
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

        # 5. CLOSING SLIDE
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

        if s.source_citation:
            fc = slide.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
            fp = fc.text_frame.paragraphs[0]
            fp.text = s.source_citation
            fp.font.size = Pt(11)
            fp.font.color.rgb = theme["text_muted"]

    prs.save(output_path)
    return output_path

# ----------------- TELEGRAM PREVIEW CARD GENERATOR -----------------
def generate_preview_card(plan: MasterpiecePresentation, output_img_path: str = "preview.png") -> str:
    W, H = 1200, 675
    theme = THEMES.get(plan.theme_key, THEMES["obsidian"])
    
    bg_tuple = (theme["dark_bg"][0], theme["dark_bg"][1], theme["dark_bg"][2])
    card_tuple = (theme["card_bg"][0], theme["card_bg"][1], theme["card_bg"][2])
    border_tuple = (theme["card_border"][0], theme["card_border"][1], theme["card_border"][2])
    accent_tuple = (theme["accent"][0], theme["accent"][1], theme["accent"][2])

    img = Image.new("RGB", (W, H), color=bg_tuple)
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (W, 10)], fill=accent_tuple)

    try:
        font_large = ImageFont.truetype("DejaVuSans-Bold.ttf", 40)
        font_sub = ImageFont.truetype("DejaVuSans.ttf", 20)
        font_tag = ImageFont.truetype("DejaVuSans-Bold.ttf", 15)
    except Exception:
        font_large = ImageFont.load_default()
        font_sub = ImageFont.load_default()
        font_tag = ImageFont.load_default()

    draw.text((70, 65), f"AKADEMIX AI  •  {len(plan.slides)} SLIDES READY", fill=accent_tuple, font=font_tag)

    title_text = plan.topic.upper()
    if len(title_text) > 42:
        title_text = title_text[:39] + "..."
    draw.text((70, 105), title_text, fill=(248, 250, 252), font=font_large)

    card_w, card_h = 330, 370
    y_card, gap = 210, 35

    for i, s in enumerate(plan.slides[:3]):
        x_card = 70 + i * (card_w + gap)
        draw.rounded_rectangle([(x_card, y_card), (x_card + card_w, y_card + card_h)], radius=12, fill=card_tuple, outline=border_tuple, width=2)
        draw.rounded_rectangle([(x_card + 20, y_card + 20), (x_card + 75, y_card + 50)], radius=6, fill=accent_tuple)
        draw.text((x_card + 32, y_card + 25), f"{s.slide_number:02d}", fill=(255, 255, 255), font=font_tag)

        s_title = s.title if len(s.title) < 26 else s.title[:23] + "..."
        draw.text((x_card + 20, y_card + 75), s_title, fill=(248, 250, 252), font=font_tag)
        draw.text((x_card + 20, y_card + 115), f"Layout: {s.layout_type}", fill=(165, 180, 205), font=font_sub)

    img.save(output_img_path)
    return output_img_path
