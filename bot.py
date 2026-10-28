import os
import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import FSInputFile
from pptx_generator import generate_presentation_json, create_pptx_file

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
bot = Bot(token=TOKEN)
dp = Dispatcher()

@dp.message(CommandStart())
async def start_handler(message: types.Message):
    await message.answer(
        "👋 Assalomu alaykum!\n\n"
        "Menga taqdimot mavzusini yuboring (masalan: *Logistika haqida*), "
        "men sizga tayyor PowerPoint (.pptx) faylini yaratib beraman.",
        parse_mode="Markdown"
    )

@dp.message()
async def topic_handler(message: types.Message):
    topic = message.text.strip()
    status_msg = await message.answer("⏳ Taqdimot tayyorlanmoqda, 10-15 soniya kuting...")

    temp_path = f"pres_{message.from_user.id}.pptx"

    try:
        # Генерация плана через Gemini и создание файла
        plan = generate_presentation_json(topic=topic, slide_count=6, language="O'zbek")
        create_pptx_file(plan, temp_path)

        document = FSInputFile(temp_path, filename=f"{topic[:25]}.pptx")
        await message.answer_document(
            document=document,
            caption=f"✅ *{plan.topic}* bo'yicha taqdimot tayyor!",
            parse_mode="Markdown"
        )
    except Exception as err:
        await message.answer(f"❌ Xatolik yuz berdi: {str(err)}")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        try:
            await status_msg.delete()
        except Exception:
            pass

async def main():
    await dp.start_polling(bot, drop_pending_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
