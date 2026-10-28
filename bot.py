import os
import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import CommandStart
from aiogram.types import FSInputFile
from pptx_generator import generate_presentation_json, create_pptx_file

BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

@dp.message(CommandStart())
async def cmd_start(message: types.Message):
    await message.answer(
        "Assalomu alaykum! Menga taqdimot mavzusini yuboring, "
        "men sizga PowerPoint (.pptx) faylini tayyorlab beraman."
    )

@dp.message()
async def generate_slides_handler(message: types.Message):
    topic = message.text.strip()
    status_msg = await message.answer("Taqdimot tayyorlanmoqda, iltimos kuting...")

    user_id = message.from_user.id
    file_path = f"presentation_{user_id}.pptx"

    try:
        # 1. Ask Gemini to structure the presentation
        plan = generate_presentation_json(topic=topic, slide_count=6, language="O'zbek")
        
        # 2. Build PPTX
        create_pptx_file(plan, file_path)

        # 3. Send back to user
        doc = FSInputFile(file_path, filename=f"{topic[:20]}_taqdimot.pptx")
        await message.answer_document(document=doc, caption=f"✅ {plan.topic} bo'yicha taqdimot tayyor!")
    except Exception as e:
        await message.answer(f"Xatolik yuz berdi: {str(e)}")
    finally:
        # Clean up temporary local file
        if os.path.exists(file_path):
            os.remove(file_path)
        await status_msg.delete()

async def main():
    await dp.start_polling(bot, drop_pending_updates=True)

if __name__ == "__main__":
    asyncio.run(main())
