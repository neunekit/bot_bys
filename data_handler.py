import pandas as pd
from datetime import datetime

FILE_NAME = "expenses.xlsx"

def save_data(data):
    try:
        # Попытка открыть существующий файл
        df = pd.read_excel(FILE_NAME)
    except FileNotFoundError:
        # Если файл не найден, создаем новый DataFrame
        df = pd.DataFrame(columns=["Дата", "Карта", "На что", "Сумма"])
    
    # Добавляем новую запись
    new_row = {
        "Дата": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Карта": data['card'],
        "На что": data['spending'],
        "Сумма": data['amount']
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    
    # Сохраняем данные в файл
    df.to_excel(FILE_NAME, index=False)
