import pandas as pd
import os
from tqdm import tqdm  # Для индикатора прогресса

# Получаем текущую рабочую директорию
folder_path = os.getcwd()

# Инициализация списка для хранения данных
all_data = []

# Функция для преобразования времени
def parse_time(time_str):
    if isinstance(time_str, str):
        if len(time_str) == 5:  # Формат HH:MM
            time_str = '00:' + time_str  # Добавляем часы
        return pd.to_datetime(time_str, format='%H:%M:%S', errors='coerce')
    return pd.NaT

# Проход по всем файлам в текущей папке с индикатором выполнения
for filename in tqdm(os.listdir(folder_path), desc="Обработка файлов"):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        
        # Чтение файла Excel
        df = pd.read_excel(file_path)

        # Добавляем имя файла как столбец даты
        df['Дата'] = filename.replace('.xlsx', '')  # Удаляем расширение для получения даты

        # Преобразуем дату и время в один столбец datetime
        df['ДатаВремя'] = pd.to_datetime(df['Дата'] + ' ' + df['Время'], format='%Y%m%d %H:%M:%S', errors='coerce')

        # Добавляем данные в общий список
        all_data.append(df)

# Объединение всех данных в один DataFrame
combined_data = pd.concat(all_data)

# Сортировка данных по времени
combined_data = combined_data.sort_values(by='ДатаВремя')

# Сброс индекса
combined_data.reset_index(drop=True, inplace=True)

# Логика проверки превышения работы более 20 часов без промывки
files_exceeding_20_hours = []
start_time = None
last_time = None
total_work_duration = pd.Timedelta(0)
exceeding_periods = []
paused_duration = pd.Timedelta(0)  # Для учета пауз

for index, row in combined_data.iterrows():
    if row['UT101_U'] > 300:
        # Начинаем отсчет времени работы, если он не начат
        if start_time is None:
            start_time = row['ДатаВремя']
            last_time = start_time
        else:
            # Если установка работает, обновляем последнее время работы
            last_time = row['ДатаВремя']
    elif row['UT101_I'] <= -1000 and row['UT101_U'] <= -20:
        # Промывка началась, сбрасываем счетчик времени
        if start_time is not None:
            # Рассчитываем длительность работы до промывки, исключая паузы более 20 минут
            work_duration = last_time - start_time - paused_duration
            total_work_duration += work_duration
            if total_work_duration > pd.Timedelta(hours=20):
                # Если суммарное время работы без промывки превышает 20 часов
                files_exceeding_20_hours.append(row['Дата'])
                exceeding_periods.append((start_time, last_time, total_work_duration))
            # Сбрасываем значения
            start_time = None
            last_time = None
            total_work_duration = pd.Timedelta(0)
            paused_duration = pd.Timedelta(0)
    elif start_time is not None and (row['UT101_U'] < 30 and row['UT101_I'] < 30):
        # Проверяем, если установка на паузе
        pause_duration = row['ДатаВремя'] - last_time
        if pause_duration <= pd.Timedelta(minutes=20):
            # Если пауза меньше или равна 20 минутам, добавляем её к общему времени паузы
            paused_duration += pause_duration
            last_time = row['ДатаВремя']  # Обновляем последнее время работы на конец паузы
        else:
            # Если пауза больше 20 минут, сбрасываем счетчик времени и учитываем накопленное время работы
            if start_time is not None:
                work_duration = last_time - start_time - paused_duration
                total_work_duration += work_duration  # Учитываем время работы без учета простоя перед промывкой
                if total_work_duration > pd.Timedelta(hours=20):
                    files_exceeding_20_hours.append(row['Дата'])
                    exceeding_periods.append((start_time, last_time, total_work_duration))
            # Сбрасываем значения
            start_time = None
            last_time = None
            total_work_duration = pd.Timedelta(0)
            paused_duration = pd.Timedelta(0)

# Проверка на финальный период, если цикл закончился без промывки
if start_time is not None:
    last_row_time = combined_data['ДатаВремя'].iloc[-1]
    work_duration = last_row_time - start_time - paused_duration
    total_work_duration += work_duration  # Учитываем только время работы без остановок перед промывкой
    if total_work_duration > pd.Timedelta(hours=20):
        files_exceeding_20_hours.append(combined_data['Дата'].iloc[-1])
        exceeding_periods.append((start_time, last_row_time, total_work_duration))

# Функция для форматирования времени и продолжительности
def format_time_duration(duration):
    total_seconds = int(duration.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours} часов {minutes} минут"

# Вывод списка файлов, которые соответствуют условиям
if files_exceeding_20_hours:
    print("Файлы, где значение 'UT101_U' превышает 300 более 20 часов без промывки:")
    for f, (start, end, duration) in zip(files_exceeding_20_hours, exceeding_periods):
        formatted_duration = format_time_duration(duration)
        print(f"{f}: Начало работы: {start.strftime('%Y-%m-%d %H:%M')}, Конец работы: {end.strftime('%Y-%m-%d %H:%M')}, Длительность работы без промывки: {formatted_duration}")
else:
    print("Нет файлов, соответствующих условиям.")
