import os
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pydantic import BaseModel, Field

# Схема структуры слайда для Pydantic
class SlideItem(BaseModel):
    slide_number: int
    layout_type: str = Field(description="title, content_with_bullets, or conclusion")
    title: str
    subtitle: str | None = None
    bullet_points: list[str] = []
    speaker_notes: str | None = None

# Схема всей презентации
class PresentationPlan(BaseModel):
    topic: str
    language: str
    total_slides: int
    slides: list[SlideItem]

def generate_presentation_json(topic: str, slide_count: int = 6, language: str = "O'zbek") -> PresentationPlan:
    api_key = os.environ.get("GEMINI_API_KEY")
    client = genai.Client(api_key=api_key)

    system_instruction = (
        "You are an expert presentation designer. Create a clear, engaging outline "
        "matching the requested slide count and language. Return valid structured JSON."
    )

    prompt = f"Mavzu: {topic}\nSlaydlar soni: {slide_count}\nTil: {language}"

    # Используем актуальную модель gemini-3.8-flash
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

def create_pptx_file(plan: PresentationPlan, output_path: str = "output.pptx") -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)  # Соотношение 16:9
    prs.slide_height = Inches(7.5)

    COLOR_PRIMARY = RGBColor(24, 43, 73)      # Темно-синий
    COLOR_ACCENT = RGBColor(0, 150, 136)      # Бирюзовый
    COLOR_TEXT = RGBColor(40, 40, 40)         # Графитовый темный

    blank_layout = prs.slide_layouts[6]

    for slide_data in plan.slides:
        slide = prs.slides.add_slide(blank_layout)

        # Титульный слайд
        if slide_data.layout_type == "title":
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
                p2.space_before = Pt(20)

        # Контентный слайд
        else:
            hdr_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.33), Inches(1.2))
            htf = hdr_box.text_frame
            htf.word_wrap = True
            hp = htf.paragraphs[0]
            hp.text = slide_data.title
            hp.font.size = Pt(32)
            hp.font.bold = True
            hp.font.color.rgb = COLOR_PRIMARY

            body_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.33), Inches(4.5))
            btf = body_box.text_frame
            btf.word_wrap = True

            for idx, pt in enumerate(slide_data.bullet_points):
                bp = btf.paragraphs[0] if idx == 0 else btf.add_paragraph()
                bp.text = f"•  {pt}"
                bp.font.size = Pt(18)
                bp.font.color.rgb = COLOR_TEXT
                bp.space_after = Pt(14)

        if slide_data.speaker_notes:
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = slide_data.speaker_notes

    prs.save(output_path)
    return output_path
