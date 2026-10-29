import os
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from pptx_generator import generate_presentation_json, create_pptx_file, generate_preview_card

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
bot = Bot(token=TOKEN)
dp = Dispatcher()

class Form(StatesGroup):
    choosing_language = State()
    choosing_theme = State()
    choosing_count = State()
    entering_topic = State()

# ----------------- KEYBOARDS -----------------
def get_language_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="lang_uz"),
            InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru"),
            InlineKeyboardButton(text="🇬🇧 English", callback_data="lang_en"),
        ]
    ])

def get_theme_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="🌌 Midnight Executive", callback_data="theme_midnight"),
            InlineKeyboardButton(text="🌲 Emerald Growth", callback_data="theme_emerald"),
        ],
        [
            InlineKeyboardButton(text="📄 Modern Minimalist", callback_data="theme_minimal")
        ]
    ])

def get_count_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="5 ta slayd", callback_data="count_5"),
            InlineKeyboardButton(text="8 ta slayd", callback_data="count_8"),
            InlineKeyboardButton(text="10 ta slayd", callback_data="count_10"),
        ]
    ])

# ----------------- HANDLERS -----------------
@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "👋 Assalomu alaykum! / Здравствуйте! / Hello!\n\n"
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

    captions = {
        "uz": "🎨 Taqdimot dizayn uslubini (Theme) tanlang:",
        "ru": "🎨 Выберите визуальный стиль (тему) презентации:",
        "en": "🎨 Select the presentation design theme:"
    }
    await callback.message.edit_text(captions.get(lang_code, captions["uz"]), reply_markup=get_theme_kb())
    await state.set_state(Form.choosing_theme)
    await callback.answer()

@dp.callback_query(Form.choosing_theme, F.data.startswith("theme_"))
async def process_theme(callback: types.CallbackQuery, state: FSMContext):
    theme_code = callback.data.split("_")[1]
    await state.update_data(theme=theme_code)
    user_data = await state.get_data()
    lang = user_data.get("language", "uz")

    captions = {
        "uz": "📊 Slaydlar sonini tanlang:",
        "ru": "📊 Выберите количество слайдов:",
        "en": "📊 Select total number of slides:"
    }
    await callback.message.edit_text(captions.get(lang, captions["uz"]), reply_markup=get_count_kb())
    await state.set_state(Form.choosing_count)
    await callback.answer()

@dp.callback_query(Form.choosing_count, F.data.startswith("count_"))
async def process_count(callback: types.CallbackQuery, state: FSMContext):
    count = int(callback.data.split("_")[1])
    await state.update_data(slide_count=count)
    user_data = await state.get_data()
    lang = user_data.get("language", "uz")

    messages = {
        "uz": f"✅ {count} ta slayd tanlandi.\n\nEndi taqdimot mavzusini yozib yuboring (masalan: *Sun'iy intellektning transport sohasidagi o'rni*):",
        "ru": f"✅ Выбрано {count} слайдов.\n\nТеперь введите тему презентации (например: *Роль искусственного интеллекта в логистике*):",
        "en": f"✅ {count} slides selected.\n\nNow enter your presentation topic (e.g.: *AI Transformation in Logistics*):"
    }
    await callback.message.edit_text(messages.get(lang, messages["uz"]), parse_mode="Markdown")
    await state.set_state(Form.entering_topic)
    await callback.answer()

@dp.message(Form.entering_topic)
async def process_topic(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    language = user_data.get("language", "uz")
    theme = user_data.get("theme", "midnight")
    slide_count = user_data.get("slide_count", 6)
    topic = message.text.strip()

    wait_messages = {
        "uz": "⏳ Masterpiece taqdimot, infografika va rasmlar yaratilmoqda...",
        "ru": "⏳ Создаем слайды высокого уровня, инфографику и визуал...",
        "en": "⏳ Crafting executive slides, bento grids, and imagery..."
    }
    status_msg = await message.answer(wait_messages.get(language, wait_messages["uz"]))

    uid = message.from_user.id
    pptx_path = f"pres_{uid}.pptx"
    preview_path = f"preview_{uid}.png"

    try:
        # 1. Generate Structured Content via Gemini
        plan = generate_presentation_json(topic=topic, slide_count=slide_count, language=language, theme=theme)

        # 2. Build PPTX File
        create_pptx_file(plan, pptx_path)

        # 3. Generate Visual Teaser Card
        generate_preview_card(plan, preview_path)

        # 4. Send Visual Preview first
        preview_file = FSInputFile(preview_path)
        await message.answer_photo(
            photo=preview_file,
            caption=f"🎯 *{plan.topic}*\nUslub: *{theme.capitalize()}* | Slaydlar: *{len(plan.slides)} ta*",
            parse_mode="Markdown"
        )

        # 5. Send PPTX Document
        doc = FSInputFile(pptx_path, filename=f"{topic[:25].replace(' ', '_')}.pptx")
        captions = {
            "uz": "✅ PowerPoint (.pptx) fayli tayyor!\nYangi slayd yaratish uchun /start ni bosing.",
            "ru": "✅ Файл PowerPoint (.pptx) готов!\nДля новой презентации нажмите /start.",
            "en": "✅ PowerPoint (.pptx) file is ready!\nTo create another, send /start."
        }
        await message.answer_document(
            document=doc,
            caption=captions.get(language, captions["uz"])
        )
    except Exception as err:
        await message.answer(f"❌ Xatolik yuz berdi: {str(err)}")
    finally:
        for p in [pptx_path, preview_path]:
            if os.path.exists(p):
                os.remove(p)
        try:
            await status_msg.delete()
        except Exception:
            pass
        await state.clear()

async def main():
    await dp.start_polling(bot, drop_pending_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
