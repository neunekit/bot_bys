from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile
import pandas as pd
from datetime import datetime

API_TOKEN = "7761443704:AAF_W6YPmBx-Ba7iMyW42s7HUBtvSbwD0Uk"  # Ваш токен

# Инициализация бота и диспетчера
bot = Bot(token=API_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# Определение состояний
class Form(StatesGroup):
    choosing_card = State()
    waiting_for_description = State()
    waiting_for_amount = State()
    confirming_clear = State()

# Функция для создания главного меню
def create_main_menu():
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Добавить запись", callback_data="add_record")],
        [InlineKeyboardButton(text="Скачать отчет", callback_data="download_report")],
        [InlineKeyboardButton(text="Очистить отчет", callback_data="clear_report")]
    ])
    return keyboard

# Функция для выбора карты
def create_card_menu():
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Тратка", callback_data="card_tratka")],
        [InlineKeyboardButton(text="Production", callback_data="card_production")]
    ])
    return keyboard

# Обработчик команды /start
@dp.message(F.text == "/start")
async def cmd_start(message: Message):
    await message.answer("Добро пожаловать! Выберите действие:", reply_markup=create_main_menu())

# Обработчик для кнопки "Добавить запись"
@dp.callback_query(F.data == "add_record")
async def add_record(callback_query: CallbackQuery, state: FSMContext):
    await callback_query.message.answer("Выберите карту для записи:", reply_markup=create_card_menu())
    await state.set_state(Form.choosing_card)

# Обработчик выбора карты
@dp.callback_query(F.data.startswith("card_"))
async def choose_card(callback_query: CallbackQuery, state: FSMContext):
    card = "Тратка" if callback_query.data == "card_tratka" else "Production"
    await state.update_data(card=card)
    await callback_query.message.answer("Введите описание траты:")
    await state.set_state(Form.waiting_for_description)

# Обработчик описания
@dp.message(Form.waiting_for_description)
async def get_description(message: Message, state: FSMContext):
    await state.update_data(description=message.text)
    await message.answer("Введите сумму:")
    await state.set_state(Form.waiting_for_amount)

# Обработчик суммы
@dp.message(Form.waiting_for_amount)
async def get_amount(message: Message, state: FSMContext):
    data = await state.get_data()
    card = data.get("card")
    description = data.get("description")
    try:
        amount = float(message.text)
    except ValueError:
        await message.answer("Сумма должна быть числом. Попробуйте еще раз.")
        return

    # Получение текущей даты и времени
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Сохранение данных в Excel
    try:
        df = pd.read_excel("expenses.xlsx")
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Дата и время", "Карта", "Описание", "Сумма"])

    new_row = {"Дата и время": current_datetime, "Карта": card, "Описание": description, "Сумма": amount}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel("expenses.xlsx", index=False)

    await message.answer("Запись добавлена! Возвращаюсь в главное меню.", reply_markup=create_main_menu())
    await state.clear()

# Обработчик для кнопки "Скачать отчет"
@dp.callback_query(F.data == "download_report")
async def download_report(callback_query: CallbackQuery):
    try:
        file = FSInputFile("expenses.xlsx")
        await callback_query.message.answer_document(file)
    except FileNotFoundError:
        await callback_query.message.answer("Отчет еще не создан. Добавьте хотя бы одну запись.")

# Обработчик для кнопки "Очистить отчет"
@dp.callback_query(F.data == "clear_report")
async def clear_report(callback_query: CallbackQuery, state: FSMContext):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Да", callback_data="confirm_clear")],
        [InlineKeyboardButton(text="Нет", callback_data="cancel_clear")]
    ])
    await callback_query.message.answer("Вы уверены, что хотите очистить отчет?", reply_markup=keyboard)
    await state.set_state(Form.confirming_clear)

# Обработчик подтверждения очистки
@dp.callback_query(F.data == "confirm_clear")
async def confirm_clear(callback_query: CallbackQuery, state: FSMContext):
    # Очистка файла
    df = pd.DataFrame(columns=["Дата и время", "Карта", "Описание", "Сумма"])
    df.to_excel("expenses.xlsx", index=False)
    await callback_query.message.answer("Отчет успешно очищен!", reply_markup=create_main_menu())
    await state.clear()

# Обработчик отмены очистки
@dp.callback_query(F.data == "cancel_clear")
async def cancel_clear(callback_query: CallbackQuery, state: FSMContext):
    await callback_query.message.answer("Очистка отчета отменена.", reply_markup=create_main_menu())
    await state.clear()

# Запуск бота
async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
