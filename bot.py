import os
import io
import asyncio
import aiohttp
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from aiogram.client.session.aiohttp import AiohttpSession
from pptx_generator import (
    generate_presentation_json,
    create_pptx_file,
    generate_preview_card,
    generate_outline_options,
    OutlineProposal
)

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")

# Total timeout (35s) must be larger than polling timeout (20s)
client_timeout = aiohttp.ClientTimeout(total=35.0, connect=10.0, sock_read=30.0)
session = AiohttpSession(timeout=client_timeout)
bot = Bot(token=TOKEN, session=session)
dp = Dispatcher()

class Form(StatesGroup):
    choosing_language = State()
    choosing_count = State()
    entering_topic = State()
    choosing_workflow = State()     # Express vs Brainstorm
    choosing_audience = State()     # Academic vs Business vs Report
    reviewing_outlines = State()    # Pick Variant A or B
    choosing_theme = State()        # Obsidian vs Emerald vs Purple
    waiting_custom_content = State()

def get_language_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="lang_uz"),
            InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru"),
            InlineKeyboardButton(text="🇬🇧 English", callback_data="lang_en"),
        ]
    ])

def get_count_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="5 ta slayd", callback_data="count_5"),
            InlineKeyboardButton(text="10 ta slayd", callback_data="count_10"),
            InlineKeyboardButton(text="16 ta slayd (Full)", callback_data="count_16"),
        ]
    ])

def get_workflow_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚡ Tayyor rejam bor (Tezkor)", callback_data="wf_express")],
        [InlineKeyboardButton(text="💡 Rejam yo'q, yordam bering (AI Maslahatchi)", callback_data="wf_brainstorm")]
    ])

def get_audience_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🎓 Akademik / Dars / Diplom", callback_data="aud_academic")],
        [InlineKeyboardButton(text="💼 Biznes / Startap / Pitch", callback_data="aud_business")],
        [InlineKeyboardButton(text="📊 Tahliliy / Rasmiy hisobot", callback_data="aud_report")]
    ])

def get_theme_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🌌 Obsidian Cyan", callback_data="theme_obsidian"),
            InlineKeyboardButton(text="🌿 Nordic Emerald", callback_data="theme_emerald"),
        ],
        [
            InlineKeyboardButton(text="👑 Deep Purple", callback_data="theme_purple")
        ]
    ])

def get_skip_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⏩ O'tkazib yuborish (Rasm/Qo'shimchasiz)", callback_data="skip_custom")]
    ])

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "👋 Assalomu alaykum!\n\n"
        "Taqdimot tilini tanlang:\n"
        "Выберите язык презентации:\n"
        "Select presentation language:",
        reply_markup=get_language_kb()
    )
    await state.set_state(Form.choosing_language)

@dp.callback_query(Form.choosing_language, F.data.startswith("lang_"))
async def process_language(callback: types.CallbackQuery, state: FSMContext):
    lang_code = callback.data.split("_")[1]
    await state.update_data(language=lang_code)

    await callback.message.edit_text("📊 Slaydlar sonini tanlang:", reply_markup=get_count_kb())
    await state.set_state(Form.choosing_count)
    await callback.answer()

@dp.callback_query(Form.choosing_count, F.data.startswith("count_"))
async def process_count(callback: types.CallbackQuery, state: FSMContext):
    count = int(callback.data.split("_")[1])
    await state.update_data(slide_count=count)

    await callback.message.edit_text(f"✅ {count} ta slayd tanlandi.\n\nEndi taqdimot mavzusini yozib yuboring:")
    await state.set_state(Form.entering_topic)
    await callback.answer()

@dp.message(Form.entering_topic)
async def process_topic(message: types.Message, state: FSMContext):
    topic = message.text.strip()
    await state.update_data(topic=topic)

    await message.answer(
        f"🎯 Mavzu: *{topic}*\n\n"
        "Sizda tayyor reja bormi yoki g'oyani birgalikda shakllantiramizmi?",
        parse_mode="Markdown",
        reply_markup=get_workflow_kb()
    )
    await state.set_state(Form.choosing_workflow)

# ----------------- EXPRESS WORKFLOW -----------------
@dp.callback_query(Form.choosing_workflow, F.data == "wf_express")
async def choose_express(callback: types.CallbackQuery, state: FSMContext):
    await callback.answer()
    await callback.message.edit_text(
        "📝 O'z rejangiz (REJA), shaxsiy rasmingiz yoki matnni yuboring.\n\n"
        "Rasm qo'shmoqchi bo'lsangiz, rasmni hozir yuboring (izohiga rejangizni ham kiritishingiz mumkin):\n\n"
        "Agar qo'shimchasiz davom etmoqchi bo'lsangiz, tugmani bosing:",
        reply_markup=get_skip_kb()
    )
    await state.set_state(Form.waiting_custom_content)

# ----------------- AI BRAINSTORMING WORKFLOW -----------------
@dp.callback_query(Form.choosing_workflow, F.data == "wf_brainstorm")
async def choose_brainstorm(callback: types.CallbackQuery, state: FSMContext):
    await callback.answer()
    await callback.message.edit_text(
        "💡 Taqdimot kim uchun mo'ljallangan? Yo'nalishni tanlang:",
        reply_markup=get_audience_kb()
    )
    await state.set_state(Form.choosing_audience)

@dp.callback_query(Form.choosing_audience, F.data.startswith("aud_"))
async def process_audience(callback: types.CallbackQuery, state: FSMContext):
    aud_type = callback.data.split("_")[1]
    await state.update_data(audience_type=aud_type)
    await callback.answer()

    user_data = await state.get_data()
    topic = user_data.get("topic")
    count = user_data.get("slide_count", 10)
    lang = user_data.get("language", "uz")

    wait_msg = await callback.message.edit_text("🤖 Siz uchun 2 xil professional reja varianti shakllantirilmoqda...")

    try:
        proposals: OutlineProposal = generate_outline_options(topic, count, aud_type, lang)
        
        await state.update_data(
            plan_a="\n".join([f"{i+1}. {p}" for i, p in enumerate(proposals.variant_a_points)]),
            plan_b="\n".join([f"{i+1}. {p}" for i, p in enumerate(proposals.variant_b_points)])
        )

        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔹 Variant A ni tanlash", callback_data="pick_plan_a")],
            [InlineKeyboardButton(text="🔸 Variant B ni tanlash", callback_data="pick_plan_b")]
        ])

        msg_text = (
            f"🎯 *Mavzu:* {topic}\n\n"
            f"━━━━━━━━━━━━━━━━━━━\n"
            f"🔹 *Variant A: {proposals.variant_a_title}*\n"
            + "\n".join([f"• {p}" for p in proposals.variant_a_points[:5]]) + "\n...\n\n"
            f"🔸 *Variant B: {proposals.variant_b_title}*\n"
            + "\n".join([f"• {p}" for p in proposals.variant_b_points[:5]]) + "\n..."
            f"\n━━━━━━━━━━━━━━━━━━━\n"
            f"Qaysi reja ma'qul keldi?"
        )
        await wait_msg.edit_text(msg_text, parse_mode="Markdown", reply_markup=kb)
        await state.set_state(Form.reviewing_outlines)
    except Exception as e:
        await callback.message.edit_text(f"❌ Xatolik yuz berdi: {str(e)}\nIltimos /start orqali qayta urinib ko'ring.")

@dp.callback_query(Form.reviewing_outlines, F.data.startswith("pick_plan_"))
async def select_plan(callback: types.CallbackQuery, state: FSMContext):
    choice = callback.data.split("_")[-1]
    user_data = await state.get_data()
    chosen_plan = user_data.get("plan_a") if choice == "a" else user_data.get("plan_b")
    await state.update_data(final_plan=chosen_plan)
    await callback.answer()

    await callback.message.edit_text(
        "🎨 Taqdimot dizayn uslubini tanlang:",
        reply_markup=get_theme_kb()
    )
    await state.set_state(Form.choosing_theme)

@dp.callback_query(Form.choosing_theme, F.data.startswith("theme_"))
async def select_theme(callback: types.CallbackQuery, state: FSMContext):
    theme_key = callback.data.split("_")[1]
    await state.update_data(theme_key=theme_key)
    await callback.answer()

    user_data = await state.get_data()
    final_plan = user_data.get("final_plan")
    await run_generation(callback.message, state, custom_img_bytes=None, custom_notes=final_plan)

# ----------------- EXPRESS INPUT RECEIVERS -----------------
@dp.callback_query(Form.waiting_custom_content, F.data == "skip_custom")
async def skip_custom(callback: types.CallbackQuery, state: FSMContext):
    await callback.answer()
    await run_generation(callback.message, state, custom_img_bytes=None, custom_notes=None)

@dp.message(Form.waiting_custom_content, F.photo)
async def receive_custom_photo(message: types.Message, state: FSMContext):
    photo = message.photo[-1]
    file = await bot.get_file(photo.file_id)
    photo_stream = io.BytesIO()
    await bot.download_file(file.file_path, photo_stream)
    photo_bytes = photo_stream.getvalue()
    
    caption = message.caption or ""
    await run_generation(message, state, custom_img_bytes=photo_bytes, custom_notes=caption if caption else None)

@dp.message(Form.waiting_custom_content, F.text)
async def receive_custom_notes(message: types.Message, state: FSMContext):
    await run_generation(message, state, custom_img_bytes=None, custom_notes=message.text)

# ----------------- CORE GENERATION CONTROLLER -----------------
async def run_generation(message: types.Message, state: FSMContext, custom_img_bytes=None, custom_notes=None):
    user_data = await state.get_data()
    topic = user_data.get("topic")
    slide_count = user_data.get("slide_count", 10)
    language = user_data.get("language", "uz")
    theme_key = user_data.get("theme_key", "obsidian")

    status_msg = await message.answer(f"⏳ {slide_count} ta slayd, mos rasmlar va ma'lumotlar tahlil qilinmoqda...")

    uid = message.from_user.id
    pptx_path = f"pres_{uid}.pptx"
    preview_path = f"preview_{uid}.png"

    try:
        plan = generate_presentation_json(
            topic=topic,
            slide_count=slide_count,
            language=language,
            user_notes=custom_notes,
            theme_key=theme_key
        )

        create_pptx_file(plan, pptx_path, custom_img_bytes=custom_img_bytes)
        generate_preview_card(plan, preview_path)

        preview_file = FSInputFile(preview_path)
        await message.answer_photo(
            photo=preview_file,
            caption=f"🎯 *{plan.topic}*\nSlaydlar: *{len(plan.slides)} ta*",
            parse_mode="Markdown"
        )

        doc = FSInputFile(pptx_path, filename=f"{topic[:25].replace(' ', '_')}.pptx")
        await message.answer_document(
            document=doc,
            caption=f"✅ *{len(plan.slides)} ta* slaydli premium taqdimot tayyor!\nYangi taqdimot yaratish uchun /start ni bosing.",
            parse_mode="Markdown"
        )
    except Exception as err:
        await message.answer(f"❌ Xatolik yuz berdi: {str(err)}")
    finally:
        for p in [pptx_path, preview_path]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass
        try:
            await status_msg.delete()
        except Exception:
            pass
        await state.clear()

# ----------------- APPLICATION ENTRYPOINT -----------------
async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot, polling_timeout=20, handle_signals=False)

if __name__ == "__main__":
    asyncio.run(main())
