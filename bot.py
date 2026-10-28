import os
import asyncio
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from pptx_generator import generate_presentation_json, create_pptx_file

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
bot = Bot(token=TOKEN)
dp = Dispatcher()

class Form(StatesGroup):
    choosing_language = State()
    entering_topic = State()

def get_language_keyboard() -> InlineKeyboardMarkup:
    buttons = [
        [
            InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="lang_uz"),
            InlineKeyboardButton(text="🇷🇺 Русский", callback_data="lang_ru"),
            InlineKeyboardButton(text="🇬🇧 English", callback_data="lang_en"),
        ]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

@dp.message(CommandStart())
async def cmd_start(message: types.Message, state: FSMContext):
    await state.clear()
    await message.answer(
        "👋 Assalomu alaykum! / Здравствуйте! / Hello!\n\n"
        "Taqdimot tilini tanlang:\n"
        "Выберите язык презентации:\n"
        "Select presentation language:",
        reply_markup=get_language_keyboard()
    )
    await state.set_state(Form.choosing_language)

@dp.callback_query(Form.choosing_language, F.data.startswith("lang_"))
async def process_language(callback: types.CallbackQuery, state: FSMContext):
    lang_code = callback.data.split("_")[1]
    await state.update_data(language=lang_code)

    messages = {
        "uz": "✅ O'zbek tili tanlandi.\n\nEndi taqdimot mavzusini yuboring (masalan: *Sun'iy intellekt kelajagi*):",
        "ru": "✅ Выбран русский язык.\n\nТеперь отправьте тему презентации (например: *Будущее искусственного интеллекта*):",
        "en": "✅ English selected.\n\nNow send the presentation topic (e.g., *Future of Artificial Intelligence*):"
    }

    await callback.message.edit_text(messages.get(lang_code, messages["uz"]), parse_mode="Markdown")
    await state.set_state(Form.entering_topic)
    await callback.answer()

@dp.message(Form.entering_topic)
async def process_topic(message: types.Message, state: FSMContext):
    user_data = await state.get_data()
    language = user_data.get("language", "uz")
    topic = message.text.strip()

    wait_messages = {
        "uz": "⏳ Taqdimot va rasmlar tayyorlanmoqda, 15-20 soniya kuting...",
        "ru": "⏳ Создаем презентацию и подбираем изображения, подождите 15-20 секунд...",
        "en": "⏳ Generating presentation and fetching visuals, please wait 15-20 seconds..."
    }
    status_msg = await message.answer(wait_messages.get(language, wait_messages["uz"]))

    temp_path = f"pres_{message.from_user.id}.pptx"

    try:
        # Generate rich presentation and build pptx
        plan = generate_presentation_json(topic=topic, slide_count=6, language=language)
        create_pptx_file(plan, temp_path)

        doc = FSInputFile(temp_path, filename=f"{topic[:25]}.pptx")
        captions = {
            "uz": f"✅ *{plan.topic}* bo'yicha taqdimot tayyor!\nYangi taqdimot uchun /start ni bosing.",
            "ru": f"✅ Презентация на тему *{plan.topic}* готова!\nДля новой презентации нажмите /start.",
            "en": f"✅ Presentation on *{plan.topic}* is ready!\nPress /start for a new presentation."
        }
        await message.answer_document(
            document=doc,
            caption=captions.get(language, captions["uz"]),
            parse_mode="Markdown"
        )
    except Exception as err:
        await message.answer(f"❌ Xatolik yuz berdi / Ошибка: {str(err)}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        try:
            await status_msg.delete()
        except Exception:
            pass
        await state.clear()

async def main():
    await dp.start_polling(bot, drop_pending_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
