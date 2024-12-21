# Импортируем библиотеку pandas для работы с данными
import pandas as pd  
# Устанавливаем библиотеку fuzzywuzzy для работы с нечетким сравнением строк
!pip install fuzzywuzzy
from fuzzywuzzy import process  # Импортируем модуль process из fuzzywuzzy
# Устанавливаем библиотеку unidecode для транслитерации
!pip install unidecode
from unidecode import unidecode  # Импортируем функцию unidecode для транслитерации
# Устанавливаем библиотеку googletrans для перевода текста
!pip install googletrans==4.0.0-rc1
from googletrans import Translator  # Импортируем класс Translator для перевода
# Устанавливаем библиотеку langdetect для определения языка текста
!pip install langdetect
from langdetect import detect, DetectorFactory  # Импортируем функции detect и DetectorFactory
тов

# Загрузка данных из файла
dn = pd.read_excel('Merged.xlsx')

# Объединяем колонки 'Фамилия', 'Имя', 'Отчество' в один столбец
dn['full_name'] = dn['Фамилия'] + ' ' + dn['Имя'] + ' ' + dn['Отчество']

# Удаляем старые колонки
dn = dn.drop(columns=['Фамилия', 'Имя', 'Отчество'])

# Удаление строк с NaN значениями
dn.dropna(how='any', inplace=True)

# Сохранение результата в новый файл
df = dn.to_excel('Changed.xlsx', index=False)

def alp(name):
    # Проверяем, состоит ли строка только из кириллических символов
    if all('а' <= char <= 'я' or 'А' <= char <= 'Я' or char == 'ё' or char == 'Ё' for char in name if char.isalpha()):
        return 'Кириллица' # Возвращаем 'Кириллица', если все символы кириллические
    # Проверяем, состоит ли строка только из латинских символов
    elif all('a' <= char <= 'z' or 'A' <= char <= 'Z' for char in name if char.isalpha()):
        return 'Латиница' # Возвращаем 'Латиница', если все символы латинские
    else:
        return 'Смешанный или неизвестный' # Возвращаем 'Смешанный или неизвестный', если есть другие символы

# Чтение существующего файла Excel
df = pd.read_excel('Changed.xlsx')

# Получение списка имен, исключая пустые значения
names = df['full_name'].dropna()

# Применяем функцию определения алфавита к каждому имени и добавляем результат в новый столбец
df['Alphabet'] = names.apply(alp)

# Запись обновленного DataFrame обратно в файл Excel
df.to_excel('Changed.xlsx', index=False)

# Вывод результата
print(df[['full_name', 'Alphabet']])

def translate_name(name, alphabet):
    # Создаем экземпляр класса Translator для перевода
    translator = Translator()
    if alphabet == 'Кириллица':
        # Перевод на английский
        translated = unidecode(name)  # Нормализация символов
        return translated
    elif alphabet == 'Латиница':
        name = unidecode(name) # Нормализация символов
        # Перевод на русский
        translated = translator.translate(name, src='en', dest='ru')
        return translated.text
    else:
        return name  # В ином случае возвращаем оригинал

# Применяем функцию определения алфавита
df['Alphabet'] = names.apply(alp)

# Применяем функцию перевода имен
df['Eng'] = names.combine(df['Alphabet'], lambda name, alphabet: name if alphabet == 'Латиница' else translate_name(name, alphabet))
df['Rus'] = names.combine(df['Alphabet'], lambda name, alphabet: name if alphabet == 'Кириллица' else translate_name(name, alphabet))

# Запись обновленного DataFrame обратно в файл Excel
df.to_excel('Changed.xlsx', index=False)

# Вывод результата
print(df[['full_name', 'Eng', 'Rus']])

def cap(name):
    # Преобразуем строку, делая заглавной первую букву каждого слова
    name = name.lower()  # Приведение к нижнему регистру
    name = ' '.join(name.split())  # Удаление лишних пробелов
    # Преобразование первой буквы каждого слова в заглавную
    return ' '.join(word.capitalize() for word in name.split())

# Чтение существующего файла Excel
df = pd.read_excel('Changed.xlsx')

# Получение списка имен, исключая пустые значения
namer = df['Rus'].dropna()
namee = df['Eng'].dropna()

# Применяем функцию капитализации к каждому имени и добавляем результат в новый столбец
df['CapR'] = namer.apply(cap)
df['CapE'] = namee.apply(cap)

# Запись обновленного DataFrame обратно в файл Excel
df.to_excel('Changed.xlsx', index=False)

# Вывод результата
print(df[['full_name', 'CapR', 'CapE']])

# Указываем столбцы для удаления
col_rem = ['full_name', 'Alphabet', 'Eng', 'Rus']

# Удаляем указанные столбцы
df.drop(columns = col_rem, inplace = True, errors = 'ignore')

# Указываем словарь с новыми именами столбцов
newcol = {
    'CapR': 'Русский',
    'CapE': 'English',
}

# Переименовываем столбцы
df.rename(columns=newcol, inplace=True)

# Указываем новый порядок столбцов
nw = ['Русский', 'English', 'Университет']

# Переупорядочиваем столбцы
df = df[nw]

# Запись обновленного DataFrame в новый файл Excel
df.to_excel('Result.xlsx', index=False)
