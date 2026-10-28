import os
import io
import requests
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pydantic import BaseModel, Field

class SlideItem(BaseModel):
    slide_number: int
    layout_type: str = Field(description="'title', 'content', or 'conclusion'")
    title: str
    subtitle: str | None = None
    bullet_points: list[str] = Field(description="3-4 detailed bullet points with deep explanations")
    key_takeaway: str | None = Field(description="A 1-sentence critical takeaway/summary insight for this slide")
    image_search_keyword: str | None = Field(description="2-3 English words for finding a relevant photo, e.g., 'freight truck highway'")
    speaker_notes: str | None = None

class PresentationPlan(BaseModel):
    topic: str
    language: str
    total_slides: int
    slides: list[SlideItem]

def generate_presentation_json(topic: str, slide_count: int = 6, language: str = "Uzbek") -> PresentationPlan:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    lang_prompts = {
        "uz": "O'zbek tili",
        "ru": "Русский язык",
        "en": "English language"
    }
    target_lang = lang_prompts.get(language, language)

    system_instruction = (
        "You are an elite presentation designer and subject matter expert. "
        "Create an in-depth, rich, and highly informative presentation structure. "
        "For each slide, provide thorough bullet points (not just 2 words, but complete explanatory thoughts), "
        "a distinct key takeaway, and a precise English 2-word keyword describing a relevant photo. "
        "Adhere strictly to the JSON schema."
    )

    prompt = f"Topic: {topic}\nTotal slides: {slide_count}\nTarget Output Language: {target_lang}"

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            response_mime_type="application/json",
            response_schema=PresentationPlan,
            temperature=0.3
        ),
    )
    return response.parsed

def fetch_image_bytes(keyword: str) -> io.BytesIO | None:
    """Fetches a royalty-free thematic photo using Unsplash source."""
    try:
        clean_keyword = requests.utils.quote(keyword)
        url = f"https://images.unsplash.com/photo-1500000000000?auto=format&fit=crop&w=800&q=80"
        # Direct search redirection endpoint for real high-res images:
        search_url = f"https://api.unsplash.com/photos/random" # Fallback to standard web search redirect:
        public_url = f"https://source.unsplash.com/800x600/?{clean_keyword}"
        
        # Safe fallback using Wikimedia or Unsplash direct redirect:
        resp = requests.get(f"https://loremflickr.com/800/600/{clean_keyword}", timeout=6)
        if resp.status_code == 200 and len(resp.content) > 1000:
            return io.BytesIO(resp.content)
    except Exception:
        pass
    return None

def create_pptx_file(plan: PresentationPlan, output_path: str = "output.pptx") -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    COLOR_PRIMARY = RGBColor(18, 38, 58)      # Deep Navy
    COLOR_ACCENT = RGBColor(0, 140, 130)      # Rich Cyan/Teal
    COLOR_TEXT = RGBColor(45, 55, 72)         # Slate text
    COLOR_CARD = RGBColor(241, 245, 249)      # Soft card background
    COLOR_MUTED = RGBColor(100, 116, 139)

    blank_layout = prs.slide_layouts[6]

    for slide_data in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # ----------------- TITLE SLIDE -----------------
        if slide_data.layout_type == "title":
            # Decorative Top Accent Bar
            accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.2))
            accent_bar.fill.solid()
            accent_bar.fill.fore_color.rgb = COLOR_ACCENT
            accent_bar.line.fill.background()

            box = slide.shapes.add_textbox(Inches(1.5), Inches(2.2), Inches(10.33), Inches(3.2))
            tf = box.text_frame
            tf.word_wrap = True

            p1 = tf.paragraphs[0]
            p1.text = slide_data.title
            p1.font.size = Pt(44)
            p1.font.bold = True
            p1.font.color.rgb = COLOR_PRIMARY
            p1.alignment = PP_ALIGN.CENTER

            if slide_data.subtitle:
                p2 = tf.add_paragraph()
                p2.text = slide_data.subtitle
                p2.font.size = Pt(22)
                p2.font.color.rgb = COLOR_ACCENT
                p2.alignment = PP_ALIGN.CENTER
                p2.space_before = Pt(16)

        # ----------------- CONTENT SLIDES -----------------
        else:
            # Title
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.9))
            ttf = title_box.text_frame
            ttf.word_wrap = True
            tp = ttf.paragraphs[0]
            tp.text = slide_data.title
            tp.font.size = Pt(28)
            tp.font.bold = True
            tp.font.color.rgb = COLOR_PRIMARY

            if slide_data.subtitle:
                sp = ttf.add_paragraph()
                sp.text = slide_data.subtitle
                sp.font.size = Pt(14)
                sp.font.color.rgb = COLOR_MUTED

            # Text Column Width
            has_image = bool(slide_data.image_search_keyword)
            text_width = Inches(7.0) if has_image else Inches(11.5)

            # Bullet Points
            body_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), text_width, Inches(3.8))
            btf = body_box.text_frame
            btf.word_wrap = True

            for idx, pt in enumerate(slide_data.bullet_points):
                bp = btf.paragraphs[0] if idx == 0 else btf.add_paragraph()
                bp.text = f"•  {pt}"
                bp.font.size = Pt(15)
                bp.font.color.rgb = COLOR_TEXT
                bp.space_after = Pt(12)

            # Key Takeaway Card at the bottom
            if slide_data.key_takeaway:
                card_top = Inches(5.8)
                card_width = text_width
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), card_top, card_width, Inches(1.0))
                card.fill.solid()
                card.fill.fore_color.rgb = COLOR_CARD
                card.line.color.rgb = COLOR_ACCENT

                ctf = card.text_frame
                ctf.word_wrap = True
                cp = ctf.paragraphs[0]
                cp.text = f"📌 Muhim xulosa / Ключевой вывод: {slide_data.key_takeaway}"
                cp.font.size = Pt(12)
                cp.font.color.rgb = COLOR_PRIMARY
                cp.font.bold = True

            # Download & Insert Image on the right column
            if has_image:
                img_stream = fetch_image_bytes(slide_data.image_search_keyword)
                if img_stream:
                    try:
                        slide.shapes.add_picture(img_stream, Inches(8.2), Inches(1.8), width=Inches(4.3))
                    except Exception:
                        pass

        # Speaker notes
        if slide_data.speaker_notes:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = slide_data.speaker_notes

    prs.save(output_path)
    return output_path
