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

# ----------------- THEMES & LUXURY PALETTES -----------------
PALETTES = [
    {
        "name": "Obsidian Cyan",
        "dark_bg": RGBColor(7, 11, 20),
        "card_bg": RGBColor(15, 23, 42),
        "card_border": RGBColor(30, 41, 59),
        "text_main": RGBColor(248, 250, 252),
        "text_muted": RGBColor(148, 163, 184),
        "accent": RGBColor(56, 189, 248),         # Sky 400
        "accent_badge": RGBColor(3, 105, 161),
        "pill_bg": RGBColor(12, 74, 110)
    },
    {
        "name": "Deep Space Purple",
        "dark_bg": RGBColor(11, 10, 26),
        "card_bg": RGBColor(22, 19, 48),
        "card_border": RGBColor(46, 38, 92),
        "text_main": RGBColor(248, 250, 252),
        "text_muted": RGBColor(165, 155, 209),
        "accent": RGBColor(168, 85, 247),        # Purple 500
        "accent_badge": RGBColor(107, 33, 168),
        "pill_bg": RGBColor(88, 28, 135)
    },
    {
        "name": "Nordic Teal",
        "dark_bg": RGBColor(6, 20, 24),
        "card_bg": RGBColor(13, 37, 44),
        "card_border": RGBColor(21, 58, 69),
        "text_main": RGBColor(241, 245, 249),
        "text_muted": RGBColor(148, 175, 185),
        "accent": RGBColor(45, 212, 191),        # Teal 400
        "accent_badge": RGBColor(15, 118, 110),
        "pill_bg": RGBColor(19, 78, 74)
    }
]

# ----------------- PYDANTIC STRUCTURED SCHEMA -----------------
class MetricBox(BaseModel):
    value: str = Field(description="Display stat, e.g. '170M', '88%', '+78M', '$50B'")
    title: str = Field(description="Bold short header, e.g. 'JOBS CREATED'")
    desc: str = Field(description="Context note under the stat")

class PillarCard(BaseModel):
    number: str = Field(description="e.g., '01', '02', '03'")
    title: str = Field(description="Headline e.g., 'AI AGENTS', 'MULTIMODAL AI'")
    subtitle: str = Field(description="One-line summary")
    desc: str = Field(description="2 sentences of crisp analytical context")

class TimelineStep(BaseModel):
    year: str = Field(description="e.g., 'NOW 2026', '2027-28', '2030'")
    title: str = Field(description="Phase title")
    desc: str = Field(description="Description of what happens")

class SlideContent(BaseModel):
    slide_number: int
    layout_type: str = Field(description="One of: 'title', 'cards_grid', 'stats_row', 'timeline', 'split_image', 'visual_hero', 'closing'")
    eyebrow: str = Field(description="Topic kicker, e.g., 'THE BIG PICTURE', 'REAL-WORLD APPLICATIONS'")
    title: str = Field(description="Slide headline, e.g., 'FOUR MAJOR SHIFTS'")
    subtitle: str | None = None
    source_citation: str | None = Field(default=None, description="e.g., 'Source: Stanford HAI, AI Index 2026'")
    bullet_points: list[str] = Field(default=[], description="Bullet points starting with bold action tags")
    cards: list[PillarCard] = Field(default=[], description="Pillars for 'cards_grid'")
    metrics: list[MetricBox] = Field(default=[], description="Stats for 'stats_row'")
    timeline_steps: list[TimelineStep] = Field(default=[], description="Steps for 'timeline'")
    takeaway_callout: str | None = Field(default=None, description="Key insight banner")
    ai_image_prompt: str | None = Field(default=None, description="Descriptive prompt for 3D/cinematic photo generation, e.g., 'futuristic glowing neon brain over smart city at night, 8k, octane render'")

class MasterpiecePresentation(BaseModel):
    topic: str
    language: str
    total_slides: int
    slides: list[SlideContent]

# ----------------- AI IMAGE GENERATOR & SEARCH -----------------
def generate_or_fetch_image(prompt_en: str | None, user_bytes: bytes | None = None) -> io.BytesIO | None:
    """Combines user upload, AI text-to-image (Pollinations Flux), and Google Images fallback."""
    if user_bytes:
        try:
            return io.BytesIO(user_bytes)
        except Exception:
            pass

    if not prompt_en:
        return None

    clean_prompt = prompt_en.strip().replace("\n", " ")

    # 1. Generate via Pollinations.AI (Flux.1 model)
    try:
        encoded_prompt = urllib.parse.quote_plus(clean_prompt)
        ai_img_url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=1280&height=720&model=flux&nologo=true"
        resp = requests.get(ai_img_url, timeout=12)
        if resp.status_code == 200 and len(resp.content) > 5000:
            bio = io.BytesIO(resp.content)
            img = Image.open(bio)
            img.verify()
            bio.seek(0)
            return bio
    except Exception as e:
        print(f"[AI Image Generation Warning]: {e}")

    # 2. Fallback to Google Custom Search
    api_key = os.environ.get("GOOGLE_SEARCH_API_KEY")
    cx = os.environ.get("GOOGLE_CSE_ID")
    if api_key and cx:
        try:
            search_q = " ".join(clean_prompt.split()[:4])
            url = "https://www.googleapis.com/customsearch/v1"
            params = {
                "key": api_key,
                "cx": cx,
                "q": search_q,
                "searchType": "image",
                "num": 2,
                "imgSize": "large",
                "safe": "active"
            }
            res = requests.get(url, params=params, timeout=8)
            if res.status_code == 200:
                items = res.json().get("items", [])
                for item in items:
                    l = item.get("link")
                    if l:
                        ir = requests.get(l, headers={"User-Agent": "Mozilla/5.0"}, timeout=6)
                        if ir.status_code == 200 and len(ir.content) > 4000:
                            bio = io.BytesIO(ir.content)
                            img = Image.open(bio)
                            img.verify()
                            bio.seek(0)
                            return bio
        except Exception as e:
            print(f"[Google Search Fallback Error]: {e}")

    # 3. Fallback to Pixabay
    pix_key = os.environ.get("PIXABAY_API_KEY")
    if pix_key:
        try:
            p_res = requests.get("https://pixabay.com/api/", params={
                "key": pix_key,
                "q": " ".join(clean_prompt.split()[:3]),
                "image_type": "photo",
                "orientation": "horizontal",
                "per_page": 3
            }, timeout=6)
            if p_res.status_code == 200:
                hits = p_res.json().get("hits", [])
                if hits:
                    u = hits[0].get("webformatURL")
                    ir = requests.get(u, timeout=6)
                    if ir.status_code == 200:
                        return io.BytesIO(ir.content)
        except Exception as e:
            print(f"[Pixabay Fallback Error]: {e}")

    return None

# ----------------- GEMINI CONTENT SYNTHESIZER -----------------
def generate_presentation_json(
    topic: str,
    slide_count: int = 10,
    language: str = "uz",
    user_notes: str | None = None
) -> MasterpiecePresentation:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_map = {
        "uz": "O'zbek tili (barcha matnlar, sarlavhalar va tahlillar 100% ravon o'zbek tilida bo'lsin)",
        "ru": "Русский язык (весь контент, даты и термины строго на русском языке)",
        "en": "English language (editorial executive presentation grade)"
    }
    target_language = lang_map.get(language, "O'zbek tili")

    # Step 1: Research with Google Search Grounding
    research_prompt = (
        f"Search Google for real verified stats, authoritative reports (e.g. Stanford HAI, WEF, Gartner), "
        f"dates, and exact projections for: '{topic}'. User notes: {user_notes or 'None'}"
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

    # Step 2: Structured Presentation Design
    system_instruction = (
        f"You are the Lead Art Director and Research Analyst who designed the Stanford & WEF Executive Deck. "
        f"You MUST generate EXACTLY {slide_count} slides in the 'slides' array. "
        f"Mandatory layout sequence for {slide_count} slides:\n"
        f"- Slide 1: 'title'\n"
        f"- Slide 2: 'cards_grid' (Why Now / 3-4 Key Pillars)\n"
        f"- Slide 3: 'split_image' (Core shifts with bold keywords)\n"
        f"- Slide 4: 'stats_row' (3 big numerical metrics with sources like WEF/Stanford)\n"
        f"- Slide 5: 'timeline' (Phases / Next 5 years trajectory)\n"
        f"- Slide 6: 'visual_hero' (High-impact visual concept slide)\n"
        f"- Intermediary slides up to {slide_count - 1}: alternate among 'cards_grid', 'split_image', and 'stats_row'.\n"
        f"- Slide {slide_count}: 'closing' (Strategic takeaway / References).\n\n"
        f"CRITICAL: For every slide that uses an image ('split_image' and 'visual_hero'), "
        f"'ai_image_prompt' MUST be a vivid English prompt for 3D/photorealistic AI generation "
        f"(e.g., 'hyper-realistic humanoid robot working with futuristic holographic interface in hospital, cinematic lighting, 8k')."
    )

    synthesis_prompt = (
        f"Topic: {topic}\n"
        f"Exact Slides Count Needed: {slide_count}\n"
        f"Language: {target_language}\n"
        f"User Custom Instructions: {user_notes or 'Produce an executive deck'}\n\n"
        f"Verified Facts & Source Context:\n{research_context[:3500]}"
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

# ----------------- POWERPOINT COMPILER -----------------
def create_pptx_file(
    plan: MasterpiecePresentation,
    output_path: str = "output.pptx",
    user_custom_image_bytes: bytes | None = None
) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    theme = random.choice(PALETTES)
    blank_layout = prs.slide_layouts[6]

    for s in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # Draw full dark background for every slide
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = theme["dark_bg"]
        bg.line.fill.background()

        # ---------------- 1. TITLE SLIDE ----------------
        if s.layout_type == "title":
            accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.8), Inches(0.08), Inches(3.6))
            accent_bar.fill.solid()
            accent_bar.fill.fore_color.rgb = theme["accent"]
            accent_bar.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.5), Inches(3.8))
            tf = tb.text_frame
            tf.word_wrap = True

            pe = tf.paragraphs[0]
            pe.text = f"{s.eyebrow.upper()}  •  2026"
            pe.font.size = Pt(13)
            pe.font.bold = True
            pe.font.color.rgb = theme["accent"]
            pe.space_after = Pt(14)

            pt = tf.add_paragraph()
            pt.text = s.title.upper()
            pt.font.size = Pt(46)
            pt.font.bold = True
            pt.font.color.rgb = theme["text_main"]
            pt.space_after = Pt(18)

            if s.subtitle:
                ps = tf.add_paragraph()
                ps.text = s.subtitle
                ps.font.size = Pt(20)
                ps.font.color.rgb = theme["text_muted"]

            # Number badge
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.2), Inches(1.8), Inches(0.9), Inches(0.45))
            badge.fill.solid()
            badge.fill.fore_color.rgb = theme["pill_bg"]
            badge.line.fill.background()
            btf = badge.text_frame
            bp = btf.paragraphs[0]
            bp.text = "01"
            bp.font.size = Pt(14)
            bp.font.bold = True
            bp.font.color.rgb = theme["accent"]
            bp.alignment = PP_ALIGN.CENTER
            continue

        # ---------------- 6. VISUAL HERO SLIDE (Full 16:9 Image + Overlay) ----------------
        if s.layout_type == "visual_hero":
            img_io = generate_or_fetch_image(s.ai_image_prompt or s.title, user_custom_image_bytes)
            if img_io:
                try:
                    slide.shapes.add_picture(img_io, Inches(0.8), Inches(1.5), width=Inches(11.733), height=Inches(5.2))
                except Exception:
                    pass

            # Header on top
            htb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(10.5), Inches(1.0))
            htf = htb.text_frame
            htf.word_wrap = True
            p1 = htf.paragraphs[0]
            p1.text = s.eyebrow.upper()
            p1.font.size = Pt(11)
            p1.font.bold = True
            p1.font.color.rgb = theme["accent"]

            p2 = htf.add_paragraph()
            p2.text = s.title
            p2.font.size = Pt(26)
            p2.font.bold = True
            p2.font.color.rgb = theme["text_main"]
            continue

        # ---------------- HEADER FOR ALL REMAINING SLIDES ----------------
        ht_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(10.5), Inches(1.1))
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
        pt.font.color.rgb = theme["text_main"]

        # Number badge on top right
        nb = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.7), Inches(0.4), Inches(0.8), Inches(0.4))
        nb.fill.solid()
        nb.fill.fore_color.rgb = theme["pill_bg"]
        nb.line.fill.background()
        np = nb.text_frame.paragraphs[0]
        np.text = f"{s.slide_number:02d}"
        np.font.size = Pt(13)
        np.font.bold = True
        np.font.color.rgb = theme["accent"]
        np.alignment = PP_ALIGN.CENTER

        # ---------------- 2. CARDS GRID (3-4 Bento Cards) ----------------
        if s.layout_type == "cards_grid":
            items = s.cards if s.cards else [
                PillarCard(number=f"0{i+1}", title=f"Pillar {i+1}", subtitle="Core driver", desc=b) 
                for i, b in enumerate(s.bullet_points[:4])
            ]
            num_cards = min(len(items), 4) or 3
            card_w = Inches(11.733 / num_cards - 0.25)
            card_h = Inches(4.3)
            start_y = Inches(1.8)

            for i, c in enumerate(items[:num_cards]):
                cx = Inches(0.8) + i * (card_w + Inches(0.25))

                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, start_y, card_w, card_h)
                card.fill.solid()
                card.fill.fore_color.rgb = theme["card_bg"]
                card.line.color.rgb = theme["card_border"]

                ctb = slide.shapes.add_textbox(cx + Inches(0.3), start_y + Inches(0.3), card_w - Inches(0.6), card_h - Inches(0.6))
                ctf = ctb.text_frame
                ctf.word_wrap = True

                p_num = ctf.paragraphs[0]
                p_num.text = c.number
                p_num.font.size = Pt(12)
                p_num.font.bold = True
                p_num.font.color.rgb = theme["accent"]
                p_num.space_after = Pt(8)

                p_title = ctf.add_paragraph()
                p_title.text = c.title
                p_title.font.size = Pt(18)
                p_title.font.bold = True
                p_title.font.color.rgb = theme["text_main"]
                p_title.space_after = Pt(6)

                if c.subtitle:
                    p_sub = ctf.add_paragraph()
                    p_sub.text = c.subtitle
                    p_sub.font.size = Pt(12)
                    p_sub.font.color.rgb = theme["accent"]
                    p_sub.space_after = Pt(10)

                p_desc = ctf.add_paragraph()
                p_desc.text = c.desc
                p_desc.font.size = Pt(13)
                p_desc.font.color.rgb = theme["text_muted"]

        # ---------------- 3. SPLIT IMAGE LAYOUT ----------------
        elif s.layout_type == "split_image":
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.3))
            tf = tb.text_frame
            tf.word_wrap = True

            for idx, pt_text in enumerate(s.bullet_points[:4]):
                p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                p.text = f"•  {pt_text}"
                p.font.size = Pt(14)
                p.font.color.rgb = theme["text_main"]
                p.space_after = Pt(16)

            # Insert Image
            img_io = generate_or_fetch_image(s.ai_image_prompt or s.title, user_custom_image_bytes)
            if img_io:
                try:
                    slide.shapes.add_picture(img_io, Inches(7.2), Inches(1.8), width=Inches(5.3), height=Inches(4.3))
                except Exception:
                    pass

        # ---------------- 4. STATS ROW ----------------
        elif s.layout_type == "stats_row":
            m_items = s.metrics if s.metrics else [
                MetricBox(value="170M", title="JOBS CREATED", desc="Structural growth areas"),
                MetricBox(value="92M", title="JOBS DISPLACED", desc="Routine tasks automation"),
                MetricBox(value="+78M", title="NET IMPACT", desc="Positive workforce transformation")
            ]
            box_w = Inches(3.65)
            box_h = Inches(3.8)

            for i, m in enumerate(m_items[:3]):
                mx = Inches(0.8) + i * (box_w + Inches(0.39))
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, Inches(1.8), box_w, box_h)
                box.fill.solid()
                box.fill.fore_color.rgb = theme["card_bg"]
                box.line.color.rgb = theme["card_border"]

                btb = slide.shapes.add_textbox(mx + Inches(0.3), Inches(2.2), box_w - Inches(0.6), box_h - Inches(0.8))
                btf = btb.text_frame
                btf.word_wrap = True

                pv = btf.paragraphs[0]
                pv.text = m.value
                pv.font.size = Pt(52)
                pv.font.bold = True
                pv.font.color.rgb = theme["accent"] if i != 1 else theme["text_main"]
                pv.space_after = Pt(12)

                ptit = btf.add_paragraph()
                ptit.text = m.title
                ptit.font.size = Pt(14)
                ptit.font.bold = True
                ptit.font.color.rgb = theme["text_main"]
                ptit.space_after = Pt(8)

                pdec = btf.add_paragraph()
                pdec.text = m.desc
                pdec.font.size = Pt(12)
                pdec.font.color.rgb = theme["text_muted"]

        # ---------------- 5. TIMELINE / ROADMAP ----------------
        elif s.layout_type == "timeline":
            steps = s.timeline_steps if s.timeline_steps else [
                TimelineStep(year="2026", title="AI Copilots", desc="Everyday assistance in software"),
                TimelineStep(year="2027-28", title="Autonomous Agents", desc="Multi-step complex execution"),
                TimelineStep(year="2030", title="Embodied Intelligence", desc="Robotics and smart city infrastructure")
            ]
            t_w = Inches(3.65)
            t_h = Inches(3.6)

            for i, st in enumerate(steps[:3]):
                tx = Inches(0.8) + i * (t_w + Inches(0.39))
                box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, Inches(1.8), t_w, t_h)
                box.fill.solid()
                box.fill.fore_color.rgb = theme["card_bg"]
                box.line.color.rgb = theme["card_border"]

                ttb = slide.shapes.add_textbox(tx + Inches(0.3), Inches(2.1), t_w - Inches(0.6), t_h - Inches(0.6))
                ttf = ttb.text_frame
                ttf.word_wrap = True

                py = ttf.paragraphs[0]
                py.text = st.year
                py.font.size = Pt(13)
                py.font.bold = True
                py.font.color.rgb = theme["accent"]
                py.space_after = Pt(10)

                pt = ttf.add_paragraph()
                pt.text = st.title
                pt.font.size = Pt(18)
                pt.font.bold = True
                pt.font.color.rgb = theme["text_main"]
                pt.space_after = Pt(10)

                pd = ttf.add_paragraph()
                pd.text = st.desc
                pd.font.size = Pt(13)
                pd.font.color.rgb = theme["text_muted"]

        # ---------------- CLOSING SLIDE ----------------
        elif s.layout_type == "closing":
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.8), Inches(10.3), Inches(4.3))
            box.fill.solid()
            box.fill.fore_color.rgb = theme["card_bg"]
            box.line.color.rgb = theme["card_border"]

            ctb = slide.shapes.add_textbox(Inches(2.0), Inches(2.2), Inches(9.3), Inches(3.5))
            ctf = ctb.text_frame
            ctf.word_wrap = True

            cp1 = ctf.paragraphs[0]
            cp1.text = "STRATEGIC SYNTHESIS"
            cp1.font.size = Pt(13)
            cp1.font.bold = True
            cp1.font.color.rgb = theme["accent"]
            cp1.space_after = Pt(14)

            cp2 = ctf.add_paragraph()
            cp2.text = s.title.upper()
            cp2.font.size = Pt(32)
            cp2.font.bold = True
            cp2.font.color.rgb = theme["text_main"]
            cp2.space_after = Pt(16)

            if s.subtitle or s.takeaway_callout:
                cp3 = ctf.add_paragraph()
                cp3.text = s.takeaway_callout or s.subtitle
                cp3.font.size = Pt(16)
                cp3.font.color.rgb = theme["text_muted"]

        # Footer Source Citation
        if s.source_citation:
            fc = slide.shapes.add_textbox(Inches(0.8), Inches(6.8), Inches(11.7), Inches(0.4))
            fp = fc.text_frame.paragraphs[0]
            fp.text = s.source_citation
            fp.font.size = Pt(10)
            fp.font.color.rgb = theme["text_muted"]

    prs.save(output_path)
    return output_path

# ----------------- TELEGRAM PREVIEW CARD -----------------
def generate_preview_card(plan: MasterpiecePresentation, output_img_path: str = "preview.png") -> str:
    W, H = 1200, 675
    img = Image.new("RGB", (W, H), color=(7, 11, 20))
    draw = ImageDraw.Draw(img)

    draw.rectangle([(0, 0), (W, 10)], fill=(56, 189, 248))

    try:
        font_large = ImageFont.truetype("DejaVuSans-Bold.ttf", 40)
        font_sub = ImageFont.truetype("DejaVuSans.ttf", 20)
        font_tag = ImageFont.truetype("DejaVuSans-Bold.ttf", 15)
    except Exception:
        font_large = ImageFont.load_default()
        font_sub = ImageFont.load_default()
        font_tag = ImageFont.load_default()

    draw.text((70, 65), f"EXECUTIVE SUITE  •  {len(plan.slides)} SLIDES READY", fill=(56, 189, 248), font=font_tag)

    title_text = plan.topic.upper()
    if len(title_text) > 42:
        title_text = title_text[:39] + "..."
    draw.text((70, 105), title_text, fill=(248, 250, 252), font=font_large)

    card_w, card_h = 330, 370
    y_card, gap = 210, 35

    for i, s in enumerate(plan.slides[:3]):
        x_card = 70 + i * (card_w + gap)
        draw.rounded_rectangle([(x_card, y_card), (x_card + card_w, y_card + card_h)], radius=12, fill=(15, 23, 42), outline=(30, 41, 59), width=2)
        draw.rounded_rectangle([(x_card + 20, y_card + 20), (x_card + 75, y_card + 50)], radius=6, fill=(3, 105, 161))
        draw.text((x_card + 32, y_card + 25), f"{s.slide_number:02d}", fill=(255, 255, 255), font=font_tag)

        s_title = s.title if len(s.title) < 26 else s.title[:23] + "..."
        draw.text((x_card + 20, y_card + 75), s_title, fill=(248, 250, 252), font=font_tag)
        draw.text((x_card + 20, y_card + 115), f"Layout: {s.layout_type}", fill=(148, 163, 184), font=font_sub)

    img.save(output_img_path)
    return output_img_path
