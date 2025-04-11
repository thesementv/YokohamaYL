import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram.filters import Command
import openpyxl
import re

API_TOKEN = '7928868644:AAGExX-NjIoSkzv_7MYCAsQ6c9kEkHwIokE'
EXCEL_FILE = 'tires.xlsx'

bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# Кнопка Назад
back_button = KeyboardButton(text="⬅️ Назад")

# Главное меню
main_kb = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Поиск по размеру"), KeyboardButton(text="Поиск по названию")]
    ],
    resize_keyboard=True
)

# Кнопки размеров
sizes = ["R12", "R13", "R14", "R15", "R16", "R17", "R17.5", "R18", "R19", "R19.5", "R20", "R22.5"]
size_kb = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text=s)] for s in sizes] + [[back_button]],
    resize_keyboard=True
)

def clean_name(name):
    """Удаляет (1), (2) и т.д. из конца названия"""
    return re.sub(r'\s*\(\d+\)$', '', str(name))

def get_price(item):
    """Возвращает правильную цену (аудитную, если она есть и выше розничной, иначе розничную)"""
    audit_price = item['audit'] if item['audit'] is not None else 0
    retail_price = item['retail'] if item['retail'] is not None else 0
    
    if audit_price > retail_price:
        return audit_price
    return retail_price

# Загрузка данных из Excel
def load_data():
    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, qty, reserve, retail, audit = row
        available = qty - (reserve if reserve else 0)
        if available > 0:
            cleaned_name = clean_name(name)
            data.append({
                "name": cleaned_name,
                "available": available,
                "retail": retail,
                "audit": audit
            })
    return data

@dp.message(Command("start"))
async def start_cmd(message: types.Message):
    await message.answer("Выберите режим поиска:", reply_markup=main_kb)

@dp.message(lambda message: message.text == "⬅️ Назад")
async def back_handler(message: types.Message):
    await message.answer("Главное меню:", reply_markup=main_kb)

@dp.message(lambda message: message.text == "Поиск по размеру")
async def size_search(message: types.Message):
    await message.answer("Выберите размер:", reply_markup=size_kb)

@dp.message(lambda message: message.text in sizes)
async def show_by_size(message: types.Message):
    data = load_data()
    size = message.text.upper()
    matches = [item for item in data if size in item["name"]]
    
    if not matches:
        await message.answer(f"Шины размера {size} отсутствуют в наличии.", reply_markup=size_kb)
        return
    
    result_message = f"🔍 Результаты поиска по размеру {size}:\n\n"
    for item in matches:
        price = get_price(item)
        result_message += f"{item['name']}\nДоступно: {item['available']} шт. | Цена: {price}₽\n\n"
    
    result_message += f"Всего доступно: {len(matches)} позиций"
    await message.answer(result_message, reply_markup=size_kb)

@dp.message(lambda message: message.text == "Поиск по названию")
async def name_search(message: types.Message):
    await message.answer("Введите часть названия шины:", reply_markup=ReplyKeyboardMarkup(
        keyboard=[[back_button]],
        resize_keyboard=True
    ))

@dp.message()
async def handle_text(message: types.Message):
    if message.text.startswith('/'):
        return
    
    query = message.text.strip().lower()
    data = load_data()
    results = [item for item in data if query in item["name"].lower()]
    
    if results:
        result_message = f"🔍 Результаты поиска по запросу '{query}':\n\n"
        for item in results:
            price = get_price(item)
            result_message += f"{item['name']}\nДоступно: {item['available']} шт. | Цена: {price}₽\n\n"
        
        result_message += f"Всего доступно: {len(results)} позиций"
        
        await message.answer(result_message, reply_markup=ReplyKeyboardMarkup(
            keyboard=[[back_button]],
            resize_keyboard=True
        ))
    else:
        await message.answer("Ничего не найдено.", reply_markup=ReplyKeyboardMarkup(
            keyboard=[[back_button]],
            resize_keyboard=True
        ))

async def main():
    await dp.start_polling(bot)

if __name__ == '__main__':
    asyncio.run(main())
