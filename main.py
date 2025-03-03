import re
import string
import pandas as pd

SHEET_URL = "ссылка"
OUTPUT_FILE = "output.xlsx"
INDEX_COL_NAME = "ID проекта"
SELECTED_COLUMNS = (
    "Название проекта",
    "Краткое описание",
    "Подробное описание",
    "Ключевые слова",
) # Важно правильно указывать колонки
SPLIT = re.compile(r'[\s-]')
LIST_ITEMS = re.compile(r"(?:\d(?:\)|\.))|(?:\b\w\)\b)") # убирает пункты у списков (например 1., 2), б)).
PUNCTUATION = re.compile(r"[^\w]")


def tokenize(text):
    return re.split(SPLIT, text)

def list_items_filter(tokens):
    return [("" if LIST_ITEMS.fullmatch(token) else token) for token in tokens]

def punctuation_filter(tokens):
    return [PUNCTUATION.sub("", token) for token in tokens]


def get_word_count(text):
    if pd.isna(text):
        return 0
    tokens = tokenize(str(text))
    # Если в ячейке число, то pandas автоматом присвоит числовой тип, поэтому надо явно изменят тип на string
    tokens = punctuation_filter(tokens)
    return len([token for token in tokens if token])

def get_char_count(text):
    if pd.isna(text):
        return 0
    tokens = tokenize(str(text))
    tokens = punctuation_filter(tokens)
    return sum([len(token) for token in tokens])

def main():
    try:
        all_sheets = pd.read_excel(SHEET_URL, sheet_name=None, index_col=INDEX_COL_NAME)

        with pd.ExcelWriter(OUTPUT_FILE) as writer:
            for sheet_name, df in all_sheets.items():
                for col in SELECTED_COLUMNS: # Цикл для добавления счетчиков слов
                    df[f"Счётчик слов {col}"] = df[col].apply(get_word_count)
                for col in SELECTED_COLUMNS: # Цикл для добавления счетчиков символов
                    df[f"Счётчик символов {col}"] = df[col].apply(get_char_count)
                # print(f"Лист: {sheet_name}") # Для вывода всех листов
                # print(df)
                # print("\n")
                df.to_excel(writer, sheet_name=sheet_name)
    except PermissionError:
        print("Закройте файл excel")


if __name__ == "__main__":
    main()
