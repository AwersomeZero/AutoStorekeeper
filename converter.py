
"""
   MADE BY: AwersomeZero
   For PepeLand with love
"""

import os
import glob
import pandas as pd
from math import ceil

SOURCE_FOLDER = 'lists'
RESULT_FOLDER = 'tables'

def parse_txt_to_list(filepath):
    """
    Читает текстовый файл и извлекает данные.
    """
    print(f"Чтение файла: {filepath}")
    processed_data = []
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                stripped_line = line.strip()

                # Строки, которые начинаются с '|' и не являются разделителями '+'
                if stripped_line.startswith('|') and not stripped_line.startswith('|+'):
                    # Разделяем по '|'
                    columns = stripped_line.split('|')
                    cleaned_columns = [col.strip() for col in columns if col.strip()]
                    if len(cleaned_columns) == 4:
                        row_to_keep = cleaned_columns[:-2]
                        processed_data.append(row_to_keep)
    except FileNotFoundError:
        print(f"ОШИБКА: Файл не найден: {filepath}")
        input()
        return None
    except Exception as e:
        print(f"ОШИБКА: Не удалось прочитать файл {filepath}. {e}")
        input()
        return None

    return processed_data


def process_files():
    """
    Находит все .txt файлы и сохраняет в .xlsx.
    """
    # Путь к папке 'lists'
    script_dir = os.path.dirname(os.path.abspath(__file__))
    lists_dir = os.path.join(script_dir, SOURCE_FOLDER)
    sorting_list = pd.read_excel(os.path.join(script_dir, "conf", "sorting_list.xlsx"), sheet_name=0).values
    sorting_dict = {}
    for i in range (len(sorting_list)):
        sorting_dict[sorting_list[i][0]] = [i, sorting_list[i][1]]
    if not os.path.isdir(lists_dir):
        print(f"ОШИБКА: Папка '{SOURCE_FOLDER}' не найдена по пути: {lists_dir}")
        print("Убедитесь, что папка 'lists' находится в той же директории, что и скрипт.")
        return

    # Находим все .txt файлы в этой папке
    txt_files = glob.glob(os.path.join(lists_dir, '*.txt'))

    if not txt_files:
        print(f"В папке '{lists_dir}' не найдено .txt файлов для обработки.")
        return

    print(f"Найдено файлов для обработки: {len(txt_files)}")

    for txt_file_path in txt_files:
        data = parse_txt_to_list(txt_file_path)
        data[0] = ['Local_ID'] + data[0] + ['Стаки'] + ['Сундуки'] + ['Бочки']
        total_chests = 0
        total_barrels = 0
        for i in range(1, len(data)-1):
            indx = sorting_dict[data[i][0]][0]
            stacks = ceil(int(data[i][1])/sorting_dict[data[i][0]][1])
            chests, barrel = [ceil(stacks / 27), 0] if stacks > 27 else [0, 1]
            total_chests += chests
            total_barrels += barrel
            data[i] = [int(indx)] + [data[i][0]] + [int(data[i][1])] + [stacks] + [chests] + [barrel]
        if not data or len(data) < 2:  # Нужен хотя бы заголовок и одна строка данных
            print(f"В файле {txt_file_path} не найдено данных (или только заголовок). Файл пропущен.")
            continue
        # Создание DataFrame, сохранение в xls
        try:
            header = data[0]
            data_big_rows = [row for row in data[1:-1] if row != header and int(row[2]) > 3456]
            data_less_rows = [row for row in data[1:-1] if row != header and int(row[2]) <= 3456]
            big_df = pd.DataFrame(data_big_rows, columns=header)
            less_df = pd.DataFrame(data_less_rows, columns=header)
            less_df = less_df.sort_values(by=['Local_ID'])
            df = pd.concat([big_df, less_df])
            df[['Итого сундуков', 'Итого бочек']] = ''
            df.iloc[0, 6] = total_chests
            df.iloc[0, 7] = total_barrels
            base_filename = os.path.basename(txt_file_path)
            excel_filename = os.path.splitext(base_filename)[0] + '.xlsx'
            output_path = os.path.join(script_dir, RESULT_FOLDER, excel_filename)
            # Сохраняем в Excel
            df.iloc[:, 1:].to_excel(output_path, index=False)
            print(f"Файл успешно сохранен: {output_path}\n")

        except Exception as e:
            print(f"ОШИБКА: Не удалось сохранить Excel файл для {txt_file_path}. {e}\n")
            input()


# Входная точка
if __name__ == "__main__":
    process_files()
    print("--- Обработка завершена ---")
    print("Нажмите Enter для выхода")
    input()