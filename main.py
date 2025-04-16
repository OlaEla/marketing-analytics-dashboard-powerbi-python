import pandas as pd
import numpy as np
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
# import locale
import re
import datetime
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, date
from io import StringIO
import xml.etree.ElementTree as ET
import os

# Загружаем данные из файла
file_path = r".\data\CopyПремиум_авто_.xlsx"
sheets = pd.read_excel(file_path, sheet_name=None)  # Загружаем все листы в словарь

# Извлекаем данные по каждому листу
df_sprav = sheets['Справочник']
df_marketing = sheets['Маркетинговые данные']
df_crm = sheets['Данные из CRM']

# Проверим первые несколько строк для понимания структуры
# print("Справочник:")
# print(df_sprav.head())

# print("\nМаркетинговые данные:")
# print(df_marketing.head())

# print("\nCRM данные:")
# print(df_crm.head())


# Обработка CRM-данных: заменяем мусорные значения на NaN
df_crm['Просчет стоимости модели'] = df_crm['Просчет стоимости модели'].replace(["-", "", "Нет данных"], np.nan)
df_crm['Приход к диллеру'] = df_crm['Приход к диллеру'].replace(["-", "", "Нет данных"], np.nan)
df_crm['Продажа'] = df_crm['Продажа'].replace(["-", "", "Нет данных"], np.nan)

# Заполняем NaN нулями и приводим к целочисленному типу
df_crm[['Просчет стоимости модели', 'Приход к диллеру', 'Продажа']] = (
    df_crm[['Просчет стоимости модели', 'Приход к диллеру', 'Продажа']]
    .fillna(0)
    .astype(int)
)

# Проверочный вывод
# print("CRM данные после обработки:")
# print(df_crm.head())

# print("\nПроверка на наличие NaN:")
# print(df_crm.isna().sum())





# Очистка данных в справочнике
df_sprav.columns = df_sprav.columns.str.strip()

# # Цена: убрать пробелы и привести к float
# df_sprav['Цена'] = df_sprav['Цена'].astype(str).str.replace(' ', '').astype(float)
# Только числовой формат без форматирования
# df_sprav['Цена'] = df_sprav['Цена'].astype(str).str.replace(r'\s+', '', regex=True).astype(float)

# Цена: убираем пробелы и оставляем тип int
# df_sprav['Цена'] = df_sprav['Цена'].str.replace(r'\s+', '', regex=True).astype(int)

df_sprav['Цена'] = df_sprav['Цена'].astype(int)



# Маржинальность: из процентов → доли
df_sprav['Маржинальность'] = df_sprav['Маржинальность'].astype(str).str.replace('%', '').astype(float)

# # Валюта: привести к единому виду
# df_sprav['Валюта'] = df_sprav['Валюта'].astype(str).str.strip().str.lower()
# df_sprav['Валюта'] = df_sprav['Валюта'].replace({'$': 'usd', 'евро': 'eur', 'рубль': 'rub'})

# Валюта: сначала заполняем пропуски, потом обрабатываем
df_sprav['Валюта'] = df_sprav['Валюта'].ffill()  # корректная замена fillna(method='ffill')
df_sprav['Валюта'] = df_sprav['Валюта'].astype(str).str.strip().str.lower()
df_sprav['Валюта'] = df_sprav['Валюта'].replace({'$': 'usd', 'евро': 'eur', 'рубль': 'rub'})
# print("\nВалюты после очистки:")
# print(df_sprav[['Модель', 'Валюта']])
# print("✅ Тип данных в столбце 'Цена':", df_sprav['Цена'].dtype)




# Заполнение пропусков в валюте значениями сверху (из-за объединённых ячеек)

df_sprav['Валюта'] = df_sprav['Валюта'].ffill()


# Марка и модель: убрать лишние пробелы и привести к единообразию
df_sprav['Марка'] = df_sprav['Марка'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
df_sprav['Модель'] = df_sprav['Модель'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

# Ручная замена "Bм W" на "BMW"
df_sprav['Марка'] = df_sprav['Марка'].replace({'Bм W': 'BMW'})

# Приводим все варианты написания Mercedes к единому стилю: 'Mercedes'
df_sprav['Марка'] = df_sprav['Марка'].apply(
    lambda x: 'Mercedes' if x.strip().lower().startswith('mercedes') else x
)

# # Проверка
# print(df_sprav)
# # print(df_sprav.head())
# print("\nПропущенные значения:")
# print(df_sprav.isnull().sum())

# print("\n🔎 Первые строки маркетинговых данных:")
# print(df_marketing.head())

# # Проверка дубликатов
# print("\n📌 Кол-во полных дубликатов строк:", df_marketing.duplicated().sum())

# # Удаление дубликатов 
# df_marketing = df_marketing.drop_duplicates()

# # Проверка пропусков
# print("\n❓ Пропущенные значения по столбцам:")
# print(df_marketing.isna().sum())

# # Проверка типов
# print("\n🔧 Типы данных:")
# print(df_marketing.dtypes)

# # Убедимся, что Client ID — строка (для слияния с CRM)
# df_marketing['Client ID'] = df_marketing['Client ID'].astype(str)

# # Проверка уникальных значений в категориальных столбцах
# cat_cols = ['Domain', 'City', 'Region', 'Source']

# for col in cat_cols:
#     if col in df_marketing.columns:
#         print(f"\n🔤 Уникальные значения в '{col}':")
#         print(df_marketing[col].value_counts(dropna=False))

# print("\n✅ Маркетинговые данные готовы для анализа!")
# print(df_marketing.head())

# print("🔎 Первые строки маркетинговых данных:")
# print(df_marketing.head())

# # Проверка пропущенных значений
# print("\n❓ Пропущенные значения по столбцам:")
# print(df_marketing.isna().sum())

# # Проверка типов данных
# print("\n🔧 Типы данных:")
# print(df_marketing.dtypes)

# # Очистка числовых столбцов
# # Столбец Goal Value содержит символ "₽", пробелы и может быть строкой
# df_marketing['Goal Value'] = (
#     df_marketing['Goal Value']
#     .astype(str)
#     .str.replace("₽", "", regex=False)
#     .str.replace(" ", "", regex=False)
#     .replace("", np.nan)
#     .astype(float)
# )

# print("\n✅ Маркетинговые данные очищены и готовы к работе!")
# print(df_marketing.head())

# Удаляем дубликаты
df_marketing = df_marketing.drop_duplicates()

# Очищаем и преобразуем столбец 'Goal Value'
df_marketing.loc[:, 'Goal Value'] = (
    df_marketing['Goal Value']
    .astype(str)
    .str.replace("₽", "", regex=False)
    .str.replace(" ", "", regex=False)
    .replace(["", "-", "–", "None", "nan", "NaN"], np.nan)
    .astype(float)
)

# Преобразуем числовые поля, содержащие текстовые символы
numeric_cols = ['Avg. Session Duration', 'Bounce Rate']
for col in numeric_cols:
    df_marketing.loc[:, col] = (
        df_marketing[col]
        .astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace("%", "", regex=False)
        .replace(["", "-", "–", "None", "nan", "NaN"], np.nan)
        .astype(float)
    )

# print("\n✅ Обработка завершена. Проверка первых строк:")
# print(df_marketing.head())

# Преобразуем столбец 'Date' из строк с русскими названиями месяцев в формат datetime64[ns].
# Это нужно для корректной работы с датой: группировок, фильтрации, построения графиков.

# 1. Создаем копию, чтобы не менять оригинальный датафрейм
df_marketing_cleaned = df_marketing.copy()

# 2. Словарь русские → английские месяцы
ru_to_en = {
    'января': 'January',
    'февраля': 'February',
    'марта': 'March',
    'апреля': 'April',
    'мая': 'May',
    'июня': 'June',
    'июля': 'July',
    'августа': 'August',
    'сентября': 'September',
    'октября': 'October',
    'ноября': 'November',
    'декабря': 'December'
}

# 3. Заменяем названия месяцев
for ru, en in ru_to_en.items():
    df_marketing_cleaned.loc[:, 'Date'] = df_marketing_cleaned['Date'].str.replace(ru, en, regex=False)

# 4. Удаляем " г." в конце
df_marketing_cleaned.loc[:, 'Date'] = df_marketing_cleaned['Date'].str.replace(r' г\.$', '', regex=True)

# 5. Преобразуем в datetime
# df_marketing_cleaned.loc[:, 'Date'] = pd.to_datetime(df_marketing_cleaned['Date'], format='%d %B %Y', errors='coerce')

# Преобразуем столбец 'Date' в datetime64[ns]
df_marketing_cleaned['Date'] = pd.to_datetime(df_marketing_cleaned['Date'], format='%d %B %Y', errors='coerce')

# Убедимся, что преобразовали без ошибок
# print(df_marketing_cleaned['Date'].dtype)  # Должен быть datetime64[ns]

# # 6. Проверяем, что всё получилось
# if pd.api.types.is_datetime64_any_dtype(df_marketing_cleaned['Date']):
#     df_marketing_cleaned.loc[:, 'Date'] = df_marketing_cleaned['Date'].dt.date
#     print("✅ Столбец 'Date' успешно преобразован к типу date.")
# else:
#     print("❌ Ошибка: столбец 'Date' не распознан как datetime. Пример плохих строк:")
#     print(df_marketing_cleaned[df_marketing_cleaned['Date'].isna()].head())

# 6. Проверяем, что всё получилось
# if df_marketing_cleaned['Date'].apply(lambda x: isinstance(x, datetime.date)).all():
#     print("✅ Столбец 'Date' содержит только значения типа datetime.date.")
# else:
#     print("❌ Ошибка: в столбце 'Date' есть значения, не являющиеся датой.")
#     print(df_marketing_cleaned[~df_marketing_cleaned['Date'].apply(lambda x: isinstance(x, datetime.date))].head())



# 7. Проверка результата
# print(df_marketing_cleaned.head())
# print(df_marketing_cleaned['Date'].head())
# print(df_marketing_cleaned['Date'].dtype)

# === ВЫДЕЛЕНИЕ МАРКИ И МОДЕЛИ ИЗ МАРКЕТИНГОВЫХ ДАННЫХ ===

# Приводим столбцы в справочнике к нужному виду
df_sprav['Марка'] = df_sprav['Марка'].str.strip().str.lower()
df_sprav['Модель'] = df_sprav['Модель'].str.strip().str.lower()

# Создаем списки марок и моделей
brands = df_sprav['Марка'].unique().tolist()
models = df_sprav['Модель'].unique().tolist()

# Добавим замену опечаток вручную (например, 'mersedes' → 'mercedes')
brand_corrections = {
    'mersedes': 'mercedes',
    'merc': 'mercedes',
}

# Приводим нужные столбцы в маркетинговых данных к нижнему регистру
columns_to_check = ['Domain', 'Keyword', 'Source', 'Goal Completion Location', 'Source conv.']
for col in columns_to_check:
    df_marketing_cleaned[col] = df_marketing_cleaned[col].astype(str).str.lower()

# Функция для поиска марки
def extract_brand(text):
    for wrong, correct in brand_corrections.items():
        if wrong in text:
            return correct
    for brand in brands:
        if brand in text:
            return brand
    return None

# Функция для поиска модели
def extract_model(text):
    for model in models:
        pattern = r'\b' + re.escape(model) + r'\b'
        if re.search(pattern, text):
            return model
    return None

# Создаем новые столбцы
df_marketing_cleaned['Марка'] = None
df_marketing_cleaned['Модель'] = None

# Заполняем их на основе всех указанных столбцов
for i, row in df_marketing_cleaned.iterrows():
    combined_text = ' '.join([str(row[col]) for col in columns_to_check])

    brand_found = extract_brand(combined_text)
    model_found = extract_model(combined_text)

    df_marketing_cleaned.at[i, 'Марка'] = brand_found
    df_marketing_cleaned.at[i, 'Модель'] = model_found

# Просмотр первых результатов
# print("🔍 Найденные строки с маркой или моделью (первые 10):\n")
# print(df_marketing_cleaned[['Марка', 'Модель']].head(10))


# === Загрузка курса ЦБ РФ на сегодня ===

# URL для получения данных о курсах валют с сайта ЦБ РФ
url = "https://www.cbr.ru/scripts/XML_daily.asp"

# Делаем запрос на сайт
response = requests.get(url)

# Проверяем успешность запроса
if response.status_code == 200:
    # Преобразуем данные в XML
    # xml_data = response.content
    xml_data = response.content.decode('windows-1251')
    # Загружаем XML в DataFrame
    data = pd.read_xml(xml_data)

    # Отбираем только нужные валюты
    currency_data = data[data['CharCode'].isin(['USD', 'EUR'])]
    # Переименовываем столбцы для удобства
    # currency_data.columns = ['Валюта', 'Курс']

    # Выводим тип данных в столбце 'Курс' перед преобразованием
    # print("Тип данных в столбце 'Курс - CharCode' до преобразования:", currency_data['CharCode'].dtype)
    # print("Тип данных в столбце 'Value' до преобразования:", currency_data['Value'].dtype)


    # Преобразуем столбец 'Курс' в числовой тип данных (с удалением запятой и конвертацией)
    # currency_data['Курс'] = currency_data['Курс'].str.replace(',', '.').astype(float)
        # Оставляем только нужные столбцы: CharCode, Name и Value
    currency_data = currency_data[['CharCode', 'Name', 'Value']]  

    # Переименовываем столбцы для удобства
    currency_data.columns = ['Валюта', 'Название', 'Курс']

    # Заменяем запятую на точку и приводим к числовому типу
    currency_data['Курс'] = currency_data['Курс'].str.replace(',', '.').astype(float)

    rub_row = pd.DataFrame([{
        'Валюта': 'RUB',
        'Название': 'Российский рубль',
        'Курс': 1.0
    }])

    currency_data = pd.concat([currency_data, rub_row], ignore_index=True)

    # Сбрасываем индекс, чтобы индексы начинались с 0
    currency_data = currency_data.reset_index(drop=True)

    # Выводим тип данных в столбце 'Валюта' после преобразованием
    print("Тип данных в столбце 'Валюта' после преобразования:", currency_data['Валюта'].dtype)
    print("Тип данных в столбце 'Курс' после преобразования:", currency_data['Курс'].dtype)

    # Выводим результат (код валюты и его значение)
    # print(currency_data[['CharCode', 'Value']])
    print(currency_data)
else:
    print("Ошибка при загрузке данных с сайта ЦБ РФ.")



    # Объедините все 4 получившиеся таблицы


# объединение справочника и курсов валют
# Добавим в df_sprav рублёвую цену и рублёвую маржу

# Приводим все валюты в df_sprav к верхнему регистру
df_sprav['Валюта'] = df_sprav['Валюта'].str.upper()

# Создаем словарь курсов валют
rates = currency_data.set_index('Валюта')['Курс'].to_dict()

# Создаем новую колонку 'Цена в рублях'
df_sprav['Цена в рублях'] = df_sprav.apply(
    lambda row: row['Цена'] * rates[row['Валюта']] if row['Валюта'] in rates else row['Цена'],
    axis=1
)

# Добавляем колонку 'Маржа в рублях'
df_sprav['Маржа в рублях'] = df_sprav['Цена в рублях'] * df_sprav['Маржинальность']

df_sprav['Цена в рублях'] = df_sprav['Цена в рублях'].round(2)
df_sprav['Маржа в рублях'] = df_sprav['Маржа в рублях'].round(2)

print("\ndf_sprav после создания колонки 'Цена в рублях' и 'Маржа в рублях'")
print(df_sprav.head())
# print(df_sprav)

# print(df_crm.columns)


# Сохраним справочник с 'Цена в рублях' и 'Маржа в рублях'
# df_sprav.to_excel(r'.\data\df_sprav.xlsx', index=False)


# Перед объединением найдем лишнюю строку: Нужно найти строку с Client ID, которая присутствует в одной таблице, но отсутствует в другой.

# Переименуем столбец 'Город' в 'City' в таблице CRM, чтобы привести их к одинаковым именам
df_crm = df_crm.rename(columns={'Город': 'City'})

# Подсчитываем частоту появления городов в каждом наборе
city_counts_crm = df_crm['City'].value_counts()
city_counts_marketing = df_marketing['City'].value_counts()

# Получаем уникальные города в каждом наборе
unique_cities_crm = city_counts_crm.index
unique_cities_marketing = city_counts_marketing.index

# Выводим количество уникальных городов
# print(f"Уникальные города в CRM: {len(unique_cities_crm)}")
# print(f"Уникальные города в Marketing: {len(unique_cities_marketing)}")

# 1. Города, которые есть в CRM, но отсутствуют в Marketing
cities_in_crm_not_in_marketing = city_counts_crm[~city_counts_crm.index.isin(city_counts_marketing.index)]

# 2. Города, которые есть в Marketing, но отсутствуют в CRM
cities_in_marketing_not_in_crm = city_counts_marketing[~city_counts_marketing.index.isin(city_counts_crm.index)]

# 3. Города, которые есть в обоих наборах, но с разной частотой
common_cities = city_counts_crm.index.intersection(city_counts_marketing.index)
different_frequency_cities = common_cities[city_counts_crm[common_cities] != city_counts_marketing[common_cities]]

# Вывод результатов
print("\nГорода, которые есть в CRM, но отсутствуют в Marketing:")
print(cities_in_crm_not_in_marketing)

print("\nГорода, которые есть в Marketing, но отсутствуют в CRM:")
print(cities_in_marketing_not_in_crm)

print("\nГорода, которые есть в обоих наборах, но с разной частотой:")
print(different_frequency_cities)

# Дополнительно выводим количество этих городов с разной частотой
for city in different_frequency_cities:
    print(f"Город: {city}, CRM: {city_counts_crm[city]}, Marketing: {city_counts_marketing[city]}")


# Для корректного объединения таблиц проверяем значения и количество строк и удаляем лишнее.


# 1. Сравним количество строк
print(f"Всего строк в CRM: {len(df_crm)}")
print(f"Всего строк в маркетинговых данных: {len(df_marketing)}")
print(f"Разница: {len(df_crm) - len(df_marketing)} строка")

# 2. Анализ по городу Krasnodar
krasnodar_crm = df_crm[df_crm['City'] == 'Krasnodar']
krasnodar_marketing = df_marketing[df_marketing['City'] == 'Krasnodar']

print(f"\nСтрок с Krasnodar в CRM: {len(krasnodar_crm)}")
print(f"Строк с Krasnodar в маркетинге: {len(krasnodar_marketing)}")
print(f"Разница по Krasnodar: {len(krasnodar_crm) - len(krasnodar_marketing)} строка")

# 3. Найдем Client ID, которые есть в CRM но нет в маркетинге для Krasnodar
crm_ids = set(krasnodar_crm['Client ID'])
marketing_ids = set(krasnodar_marketing['Client ID'])
extra_ids = crm_ids - marketing_ids

print("\nClient ID, которые есть в CRM (Krasnodar), но нет в маркетинге:")
print(extra_ids)

# 4. Если нашли такой ID, выведем полную строку
if len(extra_ids) > 0:
    problem_id = extra_ids.pop()
    problem_row = krasnodar_crm[krasnodar_crm['Client ID'] == problem_id]
    
    print("\nНайдена проблемная строка:")
    print(problem_row)
    
    # Покажем индекс этой строки для последующего удаления
    print(f"\nИндекс проблемной строки: {problem_row.index[0]}")
    
    # Можно также показать соседние строки для контекста
    print("\nСоседние строки в CRM:")
    start_idx = max(0, problem_row.index[0] - 2)
    end_idx = min(len(df_crm), problem_row.index[0] + 3)
    print(df_crm.iloc[start_idx:end_idx])
else:
    print("\nНе найдено Client ID в CRM (Krasnodar), которых нет в маркетинге. Нужен другой подход.")

# 5. Альтернативный метод: сравнение по встречаемости каждого Client ID
if len(extra_ids) == 0:
    print("\nПробуем альтернативный метод сравнения частот Client ID...")
    
    # Считаем частоту каждого Client ID для Krasnodar в обеих таблицах
    crm_freq = krasnodar_crm['Client ID'].value_counts()
    marketing_freq = krasnodar_marketing['Client ID'].value_counts()
    
    # Объединяем частоты
    freq_comparison = pd.DataFrame({
        'CRM': crm_freq,
        'Marketing': marketing_freq
    }).fillna(0)
    
    # Находим Client ID с разницей в частоте
    freq_comparison['Difference'] = freq_comparison['CRM'] - freq_comparison['Marketing']
    discrepancies = freq_comparison[freq_comparison['Difference'] > 0]
    
    if not discrepancies.empty:
        print("\nНайдены Client ID с разной частотой:")
        print(discrepancies)
        
        # Берем первый такой Client ID
        problem_id = discrepancies.index[0]
        problem_rows = krasnodar_crm[krasnodar_crm['Client ID'] == problem_id]
        
        print("\nПроблемные строки:")
        print(problem_rows)
        
        # Предлагаем удалить первую из этих строк
        print(f"\nРекомендуется удалить строку с индексом: {problem_rows.index[0]}")
    else:
        print("\nНе найдено Client ID с разной частотой. Проблема в другом.")




# Удаление строки по индексу
df_crm_clean = df_crm.drop(61937)

# Проверка
print(f"После удаления: {len(df_crm_clean)} строк в CRM")
print(f"Теперь строк с Krasnodar: {len(df_crm_clean[df_crm_clean['City'] == 'Krasnodar'])}")

# # Сохранение
# output_path = r".\data\Premium_Auto_Cleaned.xlsx"
# with pd.ExcelWriter(output_path) as writer:
#     sheets['Справочник'].to_excel(writer, sheet_name='Справочник', index=False)
#     df_marketing.to_excel(writer, sheet_name='Маркетинговые данные', index=False)
#     df_crm_clean.to_excel(writer, sheet_name='Данные из CRM', index=False)

# print(f"\nОчищенные данные сохранены в: {output_path}")


# 2: логическая связка остальных таблиц

# # 2. Проверяем ключевые столбцы перед объединением
# print("Уникальных Client ID в CRM:", df_crm_clean['Client ID'].nunique())
# print("Уникальных Client ID в маркетинге:", df_marketing['Client ID'].nunique())
# print("Уникальных пар (Client ID + City) в CRM:", 
#       df_crm_clean.groupby(['Client ID', 'City']).ngroups)
# print("Уникальных пар (Client ID + City) в маркетинге:", 
#       df_marketing.groupby(['Client ID', 'City']).ngroups)

# # 3. Объединяем по составному ключу
# df_crm_merged = pd.merge(
#     df_crm_clean,
#     df_marketing,
#     on=['Client ID', 'City'],
#     how='left',  # Сохраняем все записи из CRM
#     indicator=True,  # Добавляем столбец с информацией об источнике
#     suffixes=('_crm', '_marketing')
# )

# # 4. Анализируем результат объединения
# print("\nРезультат объединения:")
# print(f"Всего строк: {len(df_crm_merged)}")
# print("\nРаспределение по источнику:")
# print(df_crm_merged['_merge'].value_counts())

# # 5. Проверяем несопоставленные записи (если нужно)
# if 'left_only' in df_crm_merged['_merge'].value_counts():
#     print("\nНесопоставленные записи из CRM:")
#     print(df_crm_merged[df_crm_merged['_merge'] == 'left_only'][['Client ID', 'City']].drop_duplicates())



# 1. Выводим количество строк перед объединением
print(f"Строк в CRM: {len(df_crm_clean)}")
print(f"Строк в Маркетинге: {len(df_marketing_cleaned)}")


# # Проверяем типы данных ключевых столбцов:
# print("\nТипы данных ключевых столбцов:")
# print("CRM Client ID:", df_crm_clean['Client ID'].dtype)
# print("Marketing Client ID:", df_marketing_cleaned['Client ID'].dtype)
# print("CRM City:", df_crm_clean['City'].dtype)
# print("Marketing City:", df_marketing_cleaned['City'].dtype)


# 2. Простое объединение по Client ID и City (left join)
# df_crm_merged = pd.merge(
#     df_crm_clean,
#     df_marketing_cleaned,
#     on=['Client ID', 'City'],
#     how='left',
#     suffixes=('_crm', '_mkt')
# )

# 2_1 Предупреждение: количество строк изменилось. Применяем inner join...

# df_crm_merged = pd.merge(
#     df_crm_clean,
#     df_marketing_cleaned,
#     on=['Client ID', 'City'],
#     how='inner',
#     suffixes=('_crm', '_mkt')
# )
# print(f"Теперь строк: {len(df_crm_merged)} (только совпадающие записи)")


# реализуем гарантированно точное объединение 1:1 без повторного использования Client ID, где:
# Каждая строка маркетинговых данных получит не более одной соответствующей записи из CRM
# Ни один Client ID из CRM не будет использован повторно
# Количество строк полностью совпадёт с исходными маркетинговыми данными


# # Создаем копию CRM данных для пометки использованных записей
# df_crm_temp = df_crm_clean.copy()
# df_crm_temp['used'] = False  # Флаг "использован"

# # Результирующий DataFrame
# result_rows = []

# # Проходим по каждой маркетинговой записи
# for _, mkt_row in df_marketing_cleaned.iterrows():
#     client_id = mkt_row['Client ID']
    
#     # Ищем первую неиспользованную запись в CRM
#     crm_match = df_crm_temp[(df_crm_temp['Client ID'] == client_id) & 
#                            (~df_crm_temp['used'])]
    
#     if not crm_match.empty:
#         # Берем первую подходящую запись
#         crm_data = crm_match.iloc[0].to_dict()
#         df_crm_temp.at[crm_match.index[0], 'used'] = True  # Помечаем как использованную
#     else:
#         # Если нет подходящей записи - заполняем пустыми значениями
#         crm_data = {col: None for col in df_crm_clean.columns}
    
#     # Объединяем строки
#     merged_row = {**mkt_row.to_dict(), **{
#         'CRM_'+col: val for col, val in crm_data.items() 
#         if col not in ['Client ID', 'used']
#     }}
#     result_rows.append(merged_row)

# # Создаем итоговый DataFrame
# df_crm_merged = pd.DataFrame(result_rows)

# # Удаляем технические столбцы
# if 'used' in df_crm_merged:
#     df_crm_merged.drop('used', axis=1, inplace=True)


# # 1. Проверка количества строк
# assert len(df_crm_merged) == len(df_marketing_cleaned), "Количество строк изменилось!"

# # 2. Проверка уникальности Client ID из CRM
# used_ids = df_crm_temp[df_crm_temp['used']]['Client ID']
# assert len(used_ids) == len(set(used_ids)), "Есть повторно использованные ID!"

# # 3. Статистика объединения
# print(f"Успешно сопоставлено: {len(used_ids)}/{len(df_marketing_cleaned)} записей")
# print(f"Неиспользованные записи CRM: {len(df_crm_temp) - len(used_ids)}")

# Шаг 1: Добавляем порядковый номер для каждого Client ID в обеих таблицах
df_marketing_cleaned['merge_id'] = df_marketing_cleaned.groupby('Client ID').cumcount()
df_crm_clean['merge_id'] = df_crm_clean.groupby('Client ID').cumcount()

# Шаг 2: Объединяем по Client ID и порядковому номеру
df_crm_merged = pd.merge(
    df_marketing_cleaned,
    df_crm_clean.drop(columns=['City']),  # Исключаем столбец City из CRM, чтобы не было конфликта
    left_on=['Client ID', 'merge_id'],
    right_on=['Client ID', 'merge_id'],
    how='left'
)

# Шаг 3: Удаляем временный столбец
df_crm_merged.drop('merge_id', axis=1, inplace=True)

print(f"Строк после объединения: {len(df_crm_merged)} (исходно: {len(df_marketing_cleaned)})")



# 1. Проверяем уникальность использования Client ID из CRM
crm_used = df_crm_merged[df_crm_merged['Приход к диллеру'].notna()]['Client ID'].value_counts()
print("Максимальное использование Client ID из CRM:", crm_used.max())  # Должно быть 1

# 2. Пример для анализа
sample_id = df_marketing_cleaned['Client ID'].iloc[0]
print("\nПример объединения для Client ID:", sample_id)
print(df_crm_merged[df_crm_merged['Client ID'] == sample_id][['Client ID', 'Приход к диллеру']].head())

# 3. Проверяем количество строк после объединения
print(f"\nСтрок после объединения: {len(df_crm_merged)}")
print("Ожидаемо - должно совпадать с количеством строк в CRM")

# 4. Быстрая проверка первых строк
print("\nПервые 3 строки объединенной таблицы:")
print(df_crm_merged.head(3))

# Проверка столбцов в финальной таблице
print("Столбцы в объединённой таблице (полный список):")
for col in df_crm_merged.columns:
    print(f"- {col}")

# Альтернативный вариант (вывод списком с переносом строк)
print("\nСтолбцы (компактный вид):")
print(*df_crm_merged.columns.tolist(), sep='\n')


# Критически важные проверки после объединения


# Проверка 1: Сохранение количества строк CRM
assert len(df_crm_merged) == len(df_crm_clean), "Количество строк изменилось после объединения!"

# # Проверка 2: Нет дублирования строк
# assert df_crm_merged.duplicated(subset=['Client ID', 'City']).sum() == 0, "Есть дубликаты после объединения!"

# Проверка 3: Все города из CRM сохранились
assert set(df_crm_clean['City']) == set(df_crm_merged['City']), "Потерялись некоторые города!"

# Проверка 4: Нет пропусков в ключевых полях CRM
assert df_crm_merged[['Client ID', 'City']].isnull().sum().sum() == 0, "Есть пропуски в ключевых полях!"

# Столбцы в currency_data
print("\nСтолбцы в currency_data:")
print(currency_data.columns)

# Проверим столбцы снова после преобразования
print("\nСтолбцы в df_crm_merged после приведения к нижнему регистру:", df_crm_merged.columns)
print("\nСтолбцы в df_sprav после приведения к нижнему регистру:", df_sprav.columns)

# Далее объединяем данные с df_sprav на основе столбцов Марка и Модель

# Объединяем с таблицей справочника по 'Марка' и 'Модель'
# df_merged_final = df_crm_merged.merge(df_sprav, on=['Марка', 'Модель'], how='left')

# Попробуем выполнить объединение
df_merged_final = df_crm_merged.merge(df_sprav, on=['Марка', 'Модель'], how='left')


# Проверка результата
print("\nКоличество строк после объединения с таблицей справочника:", len(df_merged_final))
print(df_merged_final[['Марка', 'Модель', 'Цена', 'Валюта', 'Маржа в рублях']].head())


# Выводим несколько строк для проверки
# print(df_merged_final.head())

# # Объединяем с таблицей валют по столбцу 'Валюта'
# df_final_with_currency = df_merged_final.merge(currency_data, left_on='Валюта', right_on='Валюта', how='left')

# # Выводим результат
# print(df_final_with_currency.head())

# Добавляем строку вручную
# rub_row = pd.DataFrame([{'валюта': 'RUB', 'Название': 'Российский рубль', 'Курс': 1.0}])
# currency_data = pd.concat([currency_data, rub_row], ignore_index=True)

# Приводим к верхнему регистру для корректного объединения
# df_merged_final['Валюта'] = df_merged_final['Валюта'].str.strip().str.upper()
# currency_data['Валюта'] = currency_data['Валюта'].str.strip().str.upper()

# Проводим объединение
df_final_with_currency = df_merged_final.merge(currency_data, on='Валюта', how='left')

# Столбцы в df_final_with_currency
print("Столбцы в df_final_with_currency:")
print(df_final_with_currency.columns)

print("\nПроверка объединения df_final_with_currency:")
print(df_final_with_currency.head())
# Проверим столбцы после объединения
# print("\nстолбцы после объединения df_merged_final.merge(currency_data, on='Валюта', how='left'):")
# print(df_merged_final.columns)

# # Проверочный вывод
# print(df_final_with_currency[['Валюта', 'Название', 'Курс']].drop_duplicates())
# print(df_final_with_currency.head())
# print(currency_data[currency_data['Валюта'].isna()])
# # print(currency_data.loc[17])
# print(df_final_with_currency.iloc[34])

# print(df_final_with_currency[df_final_with_currency['Валюта'].isna()])
# print(df_final_with_currency[df_final_with_currency['Название'].isna()])
# print(df_final_with_currency[df_final_with_currency['Курс'].isna()])
# print(df_final_with_currency.isna().sum())

# Проверка столбцов в финальной таблице
print("Столбцы в объединённой таблице со СПРАВОЧНИКОМ и ВАЛЮТАМИ (полный список):")
for col in df_final_with_currency.columns:
    print(f"- {col}")

# Альтернативный вариант (вывод списком с переносом строк)
print("\nСтолбцы в объединённой таблице со СПРАВОЧНИКОМ и ВАЛЮТАМИ (компактный вид):")
print(*df_final_with_currency.columns.tolist(), sep='\n')


# ПРОВЕРКИ

# # Проверим, что марка и модель правильно добавлены
# print(df_final_with_currency[['Марка', 'Модель', 'Цена', 'Маржа в рублях']].head())

# # Проверим, сколько записей без данных по марке и модели
# missing_mark_model = df_final_with_currency[df_final_with_currency['Марка'].isna() | df_final_with_currency['Модель'].isna()]
# print(f"Записи без данных по марке и модели: {missing_mark_model.shape[0]}")

# # Проверим, сколько записей без данных по цене/маржинальности
# missing_price_margin = df_final_with_currency[df_final_with_currency['Цена'].isna() | df_final_with_currency['Маржа в рублях'].isna()]
# print(f"Записи без данных по цене и марже: {missing_price_margin.shape[0]}")

# # Проверим, что валюта корректно добавлена
# missing_currency = df_final_with_currency[df_final_with_currency['Валюта'].isna()]
# print(f"Записи без данных по валютам: {missing_currency.shape[0]}")

# # Проверим уникальные значения валют
# print(df_final_with_currency['Валюта'].unique())

# # Посмотреть первые строки маркетинговых данных, где нет марок и моделей
# print(df_final_with_currency[df_final_with_currency['Марка'].isna()].head())

# missing_data = df_final_with_currency[df_final_with_currency['Марка'].isna()]
# print(missing_data[['Марка', 'Модель', 'Цена']].drop_duplicates())

# 1. Проверка пропусков в исходных таблицах
# Проверим, сколько пропущенных значений в каждой из исходных таблиц
print("Проверка пропусков в CRM данных:")
print(df_crm.isna().sum())

print("\nПроверка пропусков в маркетинговых данных:")
print(df_marketing.isna().sum())

print("\nПроверка пропусков в справочнике:")
print(df_sprav.isna().sum())

print("\nПроверка пропусков в таблице с валютами:")
print(currency_data.isna().sum())

# 2. Проверка уникальности ключевых столбцов в исходных таблицах
# print("\nУникальные Client ID в CRM:")
# print(df_crm['Client ID'].nunique())
print("\nВсе строки в столбце Client ID(включая дубликаты) в df_crm")
print(len(df_crm['Client ID']))

# print("\nУникальные Client ID в маркетинговых данных:")
# print(df_marketing['Client ID'].nunique())
print("\nВсе строки в столбце Client ID(включая дубликаты) в df_marketing")
print(len(df_marketing['Client ID']))

# print("\nУникальные Client ID в финальной таблице после объединения:")
# print(df_final_with_currency['Client ID'].nunique())

# код вернет количество ненулевых (не NaN) значений в столбце Client ID.
print("\nКоличество ненулевых (не NaN) значений в столбце Client ID в финальной таблице после объединения")
print(df_final_with_currency['Client ID'].count())

# все строки в столбце Client ID(включая дубликаты)
print("\nВсе строки в столбце Client ID(включая дубликаты) в df_final_with_currency")
print(len(df_final_with_currency['Client ID']))

# 3. Проверка правильности объединения данных
# Объединим CRM и маркетинговые данные по Client ID
# df_crm_marketing_merged = df_crm.merge(df_marketing, on='Client ID', how='left')

# Проверим количество строк и пропуски после первого объединения
# print("\nКоличество строк после объединения CRM и маркетинга:", df_crm_marketing_merged.shape[0])
# print("Записи с пропущенными данными по кампаниям (Goal Completions, Campaign):")
# print(df_crm_marketing_merged[['Campaign', 'Goal Completions']].isna().sum())

# Объединим с таблицей справочника по марке и модели
# df_sprav_merged = df_crm_marketing_merged.merge(df_sprav, on=['Марка', 'Модель'], how='left')

# Проверим количество строк и пропуски после второго объединения
# print("\nКоличество строк после объединения с справочником:", df_sprav_merged.shape[0])
# print("Записи с пропущенными данными по цене и марже (Цена, Маржа):")
# print(df_sprav_merged[['Цена', 'Маржа']].isna().sum())

# 1. Заполнение пропусков и сортировка
df_final_with_currency['Device Category'] = df_final_with_currency['Device Category'].fillna('(none)')
df_final_with_currency['Source'] = df_final_with_currency['Source'].fillna('(none)')
df_final_with_currency.sort_values(['Client ID', 'Date'], inplace=True)

# 2. Функция для создания цепочек 
def create_final_chains(df):
    # Вспомогательная функция для цепочек
    def make_chain(values):
        chain = ' => '.join(values)
        return f"{chain} => {values[0]}" if len(values) == 1 else chain
    
    # Создаем временный DataFrame для группировки
    grouped = df.groupby('Client ID')
    
    # Device и Source цепочки
    chains = pd.DataFrame({
        'Device_Chain': grouped['Device Category'].apply(
            lambda x: make_chain(x.astype(str).tolist())
        ),
        'Source_Chain': grouped['Source'].apply(
            lambda x: make_chain(x.astype(str).tolist())
        ),
        'Total_Conversions': grouped['Конверсия'].sum(),
        'User_Touches': grouped.size()
    }).reset_index()
    
    # Комбинированная цепочка 

    chains['Touchpoint_Chain'] = grouped.apply(
        lambda g: ' => '.join(f"{d}|{s}" for d, s in zip(
            g['Device Category'], 
            g['Source']
        )) if len(g) > 1 else f"{g['Device Category'].iloc[0]}|{g['Source'].iloc[0]}"
    ).values
    
    return chains

# 3. Создаем цепочки
print("Создание финальных цепочек...")
chains_df = create_final_chains(df_final_with_currency)

# 4. Объединяем с основными данными
df_final_with_currency = df_final_with_currency.merge(
    chains_df,
    on='Client ID',
    how='left'
)

# 5. Проверка результата
print("\nФинальный результат:")
sample = df_final_with_currency.drop_duplicates('Client ID').sample(3)
for _, row in sample.iterrows():
    print(f"\nУстройства: {row['Device_Chain']}")
    print(f"Источники: {row['Source_Chain']}")
    print(f"Комбинированная: {row['Touchpoint_Chain']}")




print("\nПроверка целостности данных:")
print(f"Всего строк: {len(df_final_with_currency)}")
print(f"Уникальных цепочек: {df_final_with_currency['Touchpoint_Chain'].nunique()}")
print("Примеры цепочек:")
print(df_final_with_currency['Touchpoint_Chain'].sample(3).to_string(index=False))


# # Выведем ещё примеры для проверки
# sample = df_final_with_currency.sample(3)[[
#     'Client ID',
#     'Device_Chain',
#     'Source_Chain',
#     'Touchpoint_Chain',
#     'Total_Conversions',
#     'Has_Conversion',
#     'Total_Touches_Chain',
#     'User_Touches'
# ]]

# print("Проверка добавленных данных:")
# print(sample.to_string(index=False))

# 1. Проверим существующие столбцы
print("Существующие столбцы:", df_final_with_currency.columns.tolist())

# 2. Если каких-то столбцов нет, добавим их
required_columns = {
    'Total_Conversions': lambda df: df.groupby('Touchpoint_Chain')['Конверсия'].transform('sum'),
    'Has_Conversion': lambda df: df.groupby('Touchpoint_Chain')['Конверсия'].transform('max') > 0,
    'Total_Touches_Chain': lambda df: df.groupby('Touchpoint_Chain')['Client ID'].transform('count'),
    'User_Touches': lambda df: df.groupby('Client ID')['Client ID'].transform('count')
}

for col, func in required_columns.items():
    if col not in df_final_with_currency.columns:
        df_final_with_currency[col] = func(df_final_with_currency)
        print(f"Добавлен столбец: {col}")

# 3. Проверка формата цепочек
def validate_chains(df):
    # Проверка разделителей
    assert all(' => ' in x for x in df['Device_Chain'].dropna()), "Ошибка в Device_Chain"
    assert all(' => ' in x for x in df['Source_Chain'].dropna()), "Ошибка в Source_Chain"
    assert all('|' in x for x in df['Touchpoint_Chain'].dropna()), "Ошибка в Touchpoint_Chain"
    
    # Проверка метрик
    assert all(df['Total_Conversions'] >= 0), "Ошибка в Total_Conversions"
    assert df['Has_Conversion'].isin([True, False]).all(), "Ошибка в Has_Conversion"
    assert all(df['Total_Touches_Chain'] > 0), "Ошибка в Total_Touches_Chain"
    assert all(df['User_Touches'] > 0), "Ошибка в User_Touches"

try:
    validate_chains(df_final_with_currency)
    print("\nВсе проверки пройдены успешно!")
except AssertionError as e:
    print(f"\nОшибка проверки: {e}")

# 4. Пример данных для визуальной проверки
print("\nПример данных (проверьте формат):")
sample = df_final_with_currency.drop_duplicates('Touchpoint_Chain').sample(3)[[
    'Client ID',
    'Device_Chain',
    'Source_Chain',
    'Touchpoint_Chain',
    'Total_Conversions',
    'Has_Conversion',
    'Total_Touches_Chain',
    'User_Touches'
]]
print(sample.to_string(index=False))


# Цепочки касаний успешно добавлены в 3 форматах:
# Device_Chain (например: "desktop => desktop")
# Source_Chain (например: "google => google")
# Touchpoint_Chain (например: "desktop|google")
# Все требуемые метрики присутствуют:
# Total_Conversions - сумма конверсий по цепочкам
# Has_Conversion - признак конверсии (True/False)
# Total_Touches_Chain - количество касаний для каждой цепочки
# User_Touches - количество касаний для каждого пользователя

# Новые метрики:
# Total_Conversions - сумма конверсий для каждого пользователя
# User_Touches - количество касаний пользователя
# Total_Touches_Chain - сколько раз встречалась каждая цепочка
# Has_Conversion - флаг наличия конверсии

print("Проверка данных:")
print("- Уникальных Client ID:", df_final_with_currency['Client ID'].nunique())
print("- Первые 5 Device Category:", df_final_with_currency['Device Category'].unique()[:5])
print("- Первые 5 Source:", df_final_with_currency['Source'].unique()[:5])

# # Проверяем результат 
# print("\nДобавили недостающие данные: цепочки касаний, сумма конверсий по цепочкам, признак конверсии по цепочкам, сумма касаний для каждой цепочки и каждого пользователя:")
# df_final_with_currency[['Client ID', 'Touchpoint Chain']].head()

# # Проверим, есть ли столбец 'Touchpoint Chain' и сколько строк
# print(df_final_with_currency.columns)
# print(df_final_with_currency.shape)

# # Проверим несколько строк, включая 'Touchpoint Chain'
# print(df_final_with_currency[['Client ID', 'Touchpoint Chain']].head(10))

# # Проверим, есть ли пустые значения в 'Touchpoint Chain'
# print(df_final_with_currency['Touchpoint Chain'].isna().sum())

# # Посмотрим на несколько строк после группировки по 'Client ID'
# print(df_final_with_currency.groupby('Client ID')['Touchpoint Chain'].first().head(10))

# Сохраняем датасет

# df_final_with_currency.to_excel(r'.\data\Premium_Auto_Cleaned.xlsx', index=False)

# output_path = r".\data\Premium_Auto_Cleaned.xlsx"
# with pd.ExcelWriter(output_path) as writer:
#     sheets['Справочник'].to_excel(writer, sheet_name='Справочник', index=False)
#     df_marketing.to_excel(writer, sheet_name='Маркетинговые данные', index=False)
#     df_crm_clean.to_excel(writer, sheet_name='Данные из CRM', index=False)

# print(f"\nОчищенные данные сохранены в: {output_path}")


# Cохраняем датасет для работы в Jupyter Notebook

# df_final_with_currency.to_csv('processed_data.csv', index=False)

# # Указываем полный путь к папке data
# data_folder = r".\data"

# # Сохраняем основной датафрейм
# df_final_with_currency.to_csv(os.path.join(data_folder, 'processed_data.csv'), index=False)