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
    choosing_count = State()
    entering_topic = State()
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

    captions = {
        "uz": "📊 Slaydlar sonini tanlang (10 yoki 16 ta to'liq slayd):",
        "ru": "📊 Выберите точное количество слайдов (10 или 16 слайдов):",
        "en": "📊 Select total number of slides (10 or 16 full deck):"
    }
    await callback.message.edit_text(captions.get(lang_code, captions["uz"]), reply_markup=get_count_kb())
    await state.set_state(Form.choosing_count)
    await callback.answer()

@dp.callback_query(Form.choosing_count, F.data.startswith("count_"))
async def process_count(callback: types.CallbackQuery, state: FSMContext):
    count = int(callback.data.split("_")[1])
    await state.update_data(slide_count=count)
    user_data = await state.get_data()
    lang = user_data.get("language", "uz")

    messages = {
        "uz": f"✅ {count} ta slayd tanlandi.\n\nEndi taqdimot mavzusini yozib yuboring:",
        "ru": f"✅ Выбрано {count} слайдов.\n\nТеперь введите тему презентации:",
        "en": f"✅ {count} slides selected.\n\nNow enter your presentation topic:"
    }
    await callback.message.edit_text(messages.get(lang, messages["uz"]))
    await state.set_state(Form.entering_topic)
    await callback.answer()

@dp.message(Form.entering_topic)
async def process_topic(message: types.Message, state: FSMContext):
    topic = message.text.strip()
    await state.update_data(topic=topic)

    await message.answer(
        f"📝 Mavzu: *{topic}*\n\n"
        "Agar slaydga o'zingizning rasmingizni yoki maxsus reja/talablarni kiritmoqchi bo'lsangiz, "
        "hozir rasm yoki reja matnini (REJA) yuboring.\n\n"
        "Agar shart bo'lmasa, quyidagi tugmani bosing:",
        parse_mode="Markdown",
        reply_markup=get_skip_kb()
    )
    await state.set_state(Form.waiting_custom_content)

@dp.callback_query(Form.waiting_custom_content, F.data == "skip_custom")
async def skip_custom(callback: types.CallbackQuery, state: FSMContext):
    await callback.answer()
    await run_generation(callback.message, state, custom_img_bytes=None, custom_notes=None)

@dp.message(Form.waiting_custom_content, F.photo)
async def receive_custom_photo(message: types.Message, state: FSMContext):
    photo = message.photo[-1]
    file = await bot.get_file(photo.file_id)
    photo_bytes = (await bot.download_file(file.file_path)).read()
    caption = message.caption or ""
    await run_generation(message, state, custom_img_bytes=photo_bytes, custom_notes=caption)

@dp.message(Form.waiting_custom_content, F.text)
async def receive_custom_notes(message: types.Message, state: FSMContext):
    await run_generation(message, state, custom_img_bytes=None, custom_notes=message.text)

async def run_generation(message: types.Message, state: FSMContext, custom_img_bytes=None, custom_notes=None):
    user_data = await state.get_data()
    topic = user_data.get("topic")
    slide_count = user_data.get("slide_count", 10)
    language = user_data.get("language", "uz")

    status_msg = await message.answer(f"⏳ Taqdimot, reja asosidagi slaydlar va tahlillar tayyorlanmoqda...")

    uid = message.from_user.id
    pptx_path = f"pres_{uid}.pptx"
    preview_path = f"preview_{uid}.png"

    try:
        plan = generate_presentation_json(
            topic=topic,
            slide_count=slide_count,
            language=language,
            user_notes=custom_notes
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
            caption=f"✅ *{len(plan.slides)} ta* slaydli taqdimot tayyor!\nYangi taqdimot yaratish uchun /start ni bosing.",
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

async def main():
    await dp.start_polling(bot, drop_pending_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
