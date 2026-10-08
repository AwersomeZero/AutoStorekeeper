"""
ASK Uthility - AutoStorekeeper-v4.0.0 (beta 1)
MADE BY: AwersomeZero
For PepeLand with love

- Улучшена архитектура и читаемость (модульность, типизация, документация).
- Исправлены критические ошибки (KeyError при неизвестных блоках, сбой при пустых файлах, потеря последней строки из-за сдвига индексов).
- Добавлена поддержка JSON-файлов со списком материалов (Litematica JSON export) с автоматическим сопоставлением ID предметов и выводом их русских названий.
- Добавлена поддержка английских названий, нечувствительность к регистру и безопасный фоллбэк.
- Cтрогая сортировка по трем группам и Local_ID внутри каждой группы.
"""

import json
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set

import pandas as pd

# Папки по умолчанию
SOURCE_FOLDER = 'lists'
RESULT_FOLDER = 'tables'
CONF_FOLDER = 'conf'
SORTING_FILENAME = 'sorting_list.xlsx'

# Вместимость хранилищ в Minecraft
DOUBLE_CHEST_SLOTS = 54  # слотов в двойном сундуке
SHULKER_BOX_SLOTS = 27   # слотов в шалкербоксе

# Значения по умолчанию для неизвестных предметов
DEFAULT_STACK_SIZE = 1
UNKNOWN_LOCAL_ID = 999_999


@dataclass(frozen=True)
class ItemSortInfo:
    """Информация о предмете из базы сортировки."""
    local_id: int
    stack_size: int
    item_id: str
    ru_name: str
    en_name: str


class SortingDatabase:
    """База данных порядка сортировки и размеров стаков предметов Minecraft."""

    def __init__(
        self,
        ru_lookup: Dict[str, ItemSortInfo],
        en_lookup: Dict[str, ItemSortInfo],
        id_lookup: Dict[str, ItemSortInfo],
    ):
        self.ru_lookup = ru_lookup
        self.en_lookup = en_lookup
        self.id_lookup = id_lookup

        # Словари в нижнем регистре
        self.ru_lower = {k.lower(): v for k, v in ru_lookup.items()}
        self.en_lower = {k.lower(): v for k, v in en_lookup.items()}
        self.id_lower = {k.lower(): v for k, v in id_lookup.items()}

        # Словарь ID без префикса "minecraft:" (например: "pale_oak_log")
        self.short_id_lookup: Dict[str, ItemSortInfo] = {}
        for k, v in id_lookup.items():
            if ':' in k:
                short_k = k.split(':', 1)[1].lower()
                if short_k not in self.short_id_lookup:
                    self.short_id_lookup[short_k] = v

        self._warned_items: Set[str] = set()

    def resolve_item(self, item_query: str) -> Tuple[int, str, int]:
        """
        Ищет предмет по ID (например 'minecraft:pale_oak_log') или по названию (RU / EN).
        Возвращает кортеж: (local_id, ru_display_name, stack_size).
        Если предмет не найден, возвращает (UNKNOWN_LOCAL_ID, исходная_строка, DEFAULT_STACK_SIZE).
        """
        clean = item_query.strip()
        lower = clean.lower()

        info: Optional[ItemSortInfo] = None

        # 1. Поиск по точному ID (например "minecraft:pale_oak_log")
        if clean in self.id_lookup:
            info = self.id_lookup[clean]
        elif lower in self.id_lower:
            info = self.id_lower[lower]

        # 2. Поиск ID без/с префиксом "minecraft:"
        elif lower in self.short_id_lookup:
            info = self.short_id_lookup[lower]
        elif ('minecraft:' + lower) in self.id_lower:
            info = self.id_lower['minecraft:' + lower]

        # 3. Поиск по русскому названию
        elif clean in self.ru_lookup:
            info = self.ru_lookup[clean]
        elif lower in self.ru_lower:
            info = self.ru_lower[lower]

        # 4. Поиск по английскому названию
        elif clean in self.en_lookup:
            info = self.en_lookup[clean]
        elif lower in self.en_lower:
            info = self.en_lower[lower]

        if info is not None:
            # Приоритетно используем русское название
            display_name = info.ru_name if info.ru_name else clean
            return info.local_id, display_name, info.stack_size

        if clean not in self._warned_items:
            self._warned_items.add(clean)
            print(f"  [!] ПРЕДУПРЕЖДЕНИЕ: Предмет/ID '{clean}' не найден в базе сортировки. "
                  f"Назначен размер стака {DEFAULT_STACK_SIZE}, ID {UNKNOWN_LOCAL_ID}.")
        return UNKNOWN_LOCAL_ID, clean, DEFAULT_STACK_SIZE

    def get_item_info(self, item_name: str) -> Tuple[int, int]:
        """Для обратной совместимости: возвращает (local_id, stack_size)."""
        local_id, _, stack_size = self.resolve_item(item_name)
        return local_id, stack_size


def get_base_dir() -> Path:
    """
    Определение корневой рабочей директории.
    Корректно работает как при обычном запуске .py, так и в собранном .exe через PyInstaller (legacy)
    (независимо от того, лежит ли файл рядом с .exe или в подпапке _internal).
    """
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        # Если lists или conf лежат рядом с exe, используем директорию exe
        if (exe_dir / SOURCE_FOLDER).exists() or (exe_dir / CONF_FOLDER).exists():
            return exe_dir
        # Если exe в корне, а данные внутри _internal
        if (exe_dir / '_internal' / CONF_FOLDER).exists():
            return exe_dir
        return exe_dir
    return Path(__file__).resolve().parent


def find_conf_file(base_dir: Path) -> Path:
    """Ищет файл sorting_list.xlsx по стандартным путям."""
    candidates = [
        base_dir / CONF_FOLDER / SORTING_FILENAME,
        base_dir / '_internal' / CONF_FOLDER / SORTING_FILENAME,
        Path(__file__).resolve().parent / CONF_FOLDER / SORTING_FILENAME,
    ]
    for p in candidates:
        if p.is_file():
            return p
    raise FileNotFoundError(
        f"Файл базы сортировки '{SORTING_FILENAME}' не найден ни в одной из директорий:\n" +
        "\n".join(f" - {p}" for p in candidates)
    )


def load_sorting_database(conf_path: Path) -> SortingDatabase:
    """Загружает базу сортировки из Excel-файла."""
    try:
        df = pd.read_excel(conf_path, sheet_name=0)
    except Exception as e:
        raise RuntimeError(f"Не удалось прочитать файл базы сортировки {conf_path}: {e}")

    ru_lookup: Dict[str, ItemSortInfo] = {}
    en_lookup: Dict[str, ItemSortInfo] = {}
    id_lookup: Dict[str, ItemSortInfo] = {}

    for idx, row in df.iterrows():
        ru_name = str(row.get('RU_Name', '')).strip()
        en_name = str(row.get('EN_name', '')).strip()
        item_id = str(row.get('ID', '')).strip()

        try:
            size_val = int(row.get('Size', DEFAULT_STACK_SIZE))
            if size_val <= 0:
                size_val = DEFAULT_STACK_SIZE
        except (ValueError, TypeError):
            size_val = DEFAULT_STACK_SIZE

        info = ItemSortInfo(
            local_id=int(idx),
            stack_size=size_val,
            item_id=item_id,
            ru_name=ru_name,
            en_name=en_name,
        )

        # Сохраняем первое вхождение для предотвращения перезаписи дубликатами
        if ru_name and ru_name not in ru_lookup:
            ru_lookup[ru_name] = info
        if en_name and en_name not in en_lookup:
            en_lookup[en_name] = info
        if item_id and item_id not in id_lookup:
            id_lookup[item_id] = info

    return SortingDatabase(ru_lookup, en_lookup, id_lookup)


def parse_material_list(filepath: Path) -> List[Tuple[str, int]]:
    """
    Читает текстовый файл MaterialList из Litematica и извлекает предметы и их количество.
    Корректно фильтрует рамки таблиц, заголовки и повторяющийся футер Litematica.
    """
    content = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1251'):
        try:
            with open(filepath, 'r', encoding=enc) as f:
                content = f.readlines()
            break
        except (UnicodeDecodeError, FileNotFoundError):
            continue

    if content is None:
        raise RuntimeError(f"Не удалось прочитать файл {filepath} (ошибка кодировки или файл отсутствует)")

    items: List[Tuple[str, int]] = []
    for line in content:
        stripped = line.strip()
        if not stripped.startswith('|') or stripped.startswith('|+'):
            continue

        # Разбиваем по '|'
        columns = [col.strip() for col in stripped.strip('|').split('|')]
        if len(columns) < 2:
            continue

        item_name = columns[0]
        total_str = columns[1]

        # Пропускаем заголовки (в начале и в конце таблицы)
        if item_name.lower() in ('item', 'предмет', 'название', 'наименование') or \
           total_str.lower() in ('total', 'всего', 'кол-во', 'количество'):
            continue

        # Очищаем число от разделителей разрядов
        clean_total = total_str.replace(' ', '').replace(',', '').replace('\xa0', '')
        if not clean_total.isdigit():
            continue

        total = int(clean_total)
        if item_name:
            items.append((item_name, total))

    return items


def parse_json_material_list(filepath: Path) -> List[Tuple[str, int]]:
    """
    Парсинг JSON-файла со списком материалов из Litematica.
    Поддерживает как объект с ключом 'Materials', так и прямой список элементов.
    Возвращает список кортежей (идентификатор_или_название_предмета, количество).
    """
    content = None
    for enc in ('utf-8-sig', 'utf-8', 'cp1251'):
        try:
            with open(filepath, 'r', encoding=enc) as f:
                content = json.load(f)
            break
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue

    if content is None:
        raise RuntimeError(f"Не удалось прочитать или декодировать JSON-файл {filepath}")

    raw_items = None
    if isinstance(content, dict):
        raw_items = content.get('Materials') or content.get('materials') or content.get('Items') or content.get('items')
        if raw_items is None:
            raw_items = content
    else:
        raw_items = content

    items: List[Tuple[str, int]] = []

    if isinstance(raw_items, list):
        for entry in raw_items:
            if not isinstance(entry, dict):
                continue
            item_id = entry.get('Item') or entry.get('item') or entry.get('ID') or entry.get('id') or entry.get('Name') or entry.get('name')
            total_val = entry.get('Total') if 'Total' in entry else entry.get('total')
            if total_val is None:
                total_val = entry.get('Count') if 'Count' in entry else entry.get('count', 0)

            if item_id is None:
                continue

            try:
                total = int(str(total_val).replace(' ', '').replace(',', '').replace('\xa0', ''))
            except (ValueError, TypeError):
                continue

            clean_id = str(item_id).strip()
            if clean_id:
                items.append((clean_id, total))

    elif isinstance(raw_items, dict):
        for k, v in raw_items.items():
            try:
                total = int(str(v).replace(' ', '').replace(',', '').replace('\xa0', ''))
            except (ValueError, TypeError):
                continue
            clean_id = str(k).strip()
            if clean_id:
                items.append((clean_id, total))

    return items


def parse_material_file(filepath: Path) -> List[Tuple[str, int]]:
    """Определяет тип файла (.json или .txt) и парсит его в список материалов."""
    if filepath.suffix.lower() == '.json':
        return parse_json_material_list(filepath)
    return parse_material_list(filepath)


def parse_txt_to_list(filepath: str) -> Optional[List[List[str]]]:
    """
    Функция для обратной совместимости с оригинальным интерфейсом converter.py.
    Возвращает список строк [['Item', 'Total'], [предмет, количество], ...].
    """
    try:
        raw_items = parse_material_file(Path(filepath))
        return [['Item', 'Total']] + [[name, str(count)] for name, count in raw_items]
    except Exception as e:
        print(f"ОШИБКА: Не удалось прочитать файл {filepath}. {e}")
        return None


def convert_materials_to_dataframe(
    materials: List[Tuple[str, int]],
    sorting_db: SortingDatabase,
) -> pd.DataFrame:
    """
    Преобразует список материалов в DataFrame со столбцами:
    Local_ID, Item, Total, Стаки, Даблчесты, Шалкербоксы.
    
    Для JSON-файлов сопоставляет ID блоков с их русскими названиями.
    Сортировка по трем группам:
    1. Более 27 стаков (stacks > 27)
    2. От 1 до 27 стаков (total >= stack_size и stacks <= 27)
    3. Меньше 1 стака (total < stack_size)
    Внутри каждой группы проходит сортировка по порядку Local_ID из sorting_list.xlsx.
    """
    if not materials:
        return pd.DataFrame(columns=['Local_ID', 'Item', 'Total', 'Стаки', 'Даблчесты', 'Шалкербоксы'])

    rows = []
    for item_raw, total in materials:
        local_id, ru_display_name, stack_size = sorting_db.resolve_item(item_raw)

        stacks = math.ceil(total / stack_size) if stack_size > 0 else 0
        chests = math.ceil(stacks / DOUBLE_CHEST_SLOTS) if stacks > 0 else 0
        shulkerboxes = math.ceil(stacks / SHULKER_BOX_SLOTS) if stacks > 0 else 0

        # Определение группы:
        # Группа 1: Более 27 стаков
        # Группа 2: От 1 до 27 стаков
        # Группа 3: Меньше 1 стака
        if stacks > 27:
            group_priority = 1
        elif total >= stack_size and stacks <= 27:
            group_priority = 2
        else:
            group_priority = 3

        rows.append({
            'Local_ID': local_id,
            'Item': ru_display_name,
            'Total': total,
            'Стаки': stacks,
            'Даблчесты': chests,
            'Шалкербоксы': shulkerboxes,
            '_group': group_priority,
        })

    df = pd.DataFrame(rows)
    # Сортировка:
    # 1. По группе (1 -> 2 -> 3)
    # 2. Внутри группы строго по Local_ID из creative inventory
    # 3. При совпадении Local_ID — по убыванию Total
    df = df.sort_values(by=['_group', 'Local_ID', 'Total'], ascending=[True, True, False])
    df = df.drop(columns=['_group']).reset_index(drop=True)
    return df


def save_dataframe_to_excel(df: pd.DataFrame, output_path: Path) -> None:
    """Сохраняет DataFrame в Excel с автоподбором ширины колонок."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
    except Exception:
        df.to_excel(output_path, index=False)


def process_single_file(file_path: Path, output_dir: Path, sorting_db: SortingDatabase) -> bool:
    """Обрабатывает один файл MaterialList (.txt или .json)."""
    print(f"Чтение файла: {file_path}")
    try:
        materials = parse_material_file(file_path)
    except Exception as e:
        print(f"ОШИБКА: Не удалось прочитать {file_path}. {e}")
        return False

    if not materials:
        print(f"В файле {file_path} не найдено строк с материалами. Файл пропущен.\n")
        return False

    try:
        df = convert_materials_to_dataframe(materials, sorting_db)
        excel_filename = file_path.stem + '.xlsx'
        output_path = output_dir / excel_filename
        save_dataframe_to_excel(df, output_path)
        print(f"Файл успешно сохранен: {output_path} (обработано предметов: {len(df)})\n")
        return True
    except Exception as e:
        print(f"ОШИБКА: Не удалось сохранить Excel для {file_path}. {e}\n")
        return False


def process_files() -> None:
    """Находит все .txt и .json файлы в папке lists и преобразует их в .xlsx в папке tables."""
    base_dir = get_base_dir()

    # Поиск базы данных сортировки
    try:
        conf_file = find_conf_file(base_dir)
        print(f"Загрузка базы сортировки: {conf_file}")
        sorting_db = load_sorting_database(conf_file)
    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
        return

    # Определение папки с входными файлами
    lists_dir = base_dir / SOURCE_FOLDER
    if not lists_dir.is_dir() and (base_dir / '_internal' / SOURCE_FOLDER).is_dir():
        lists_dir = base_dir / '_internal' / SOURCE_FOLDER

    if not lists_dir.is_dir():
        print(f"ОШИБКА: Папка '{SOURCE_FOLDER}' не найдена по пути: {lists_dir}")
        print("Убедитесь, что папка 'lists' находится в той же директории, что и скрипт/программа.")
        return

    # Папка для результатов
    tables_dir = base_dir / RESULT_FOLDER

    # Поиск входных файлов (.txt и .json)
    input_files: List[Path] = []
    for ext in ('*.txt', '*.json'):
        input_files.extend(lists_dir.glob(ext))
    input_files = sorted(input_files, key=lambda p: p.name.lower())

    if not input_files:
        print(f"В папке '{lists_dir}' не найдено файлов (.txt, .json) для обработки.")
        return

    print(f"Найдено файлов для обработки: {len(input_files)}\n")

    successful_count = 0
    for input_file in input_files:
        if process_single_file(input_file, tables_dir, sorting_db):
            successful_count += 1

    print(f"Успешно обработано: {successful_count} из {len(input_files)} файлов.")


if __name__ == "__main__":
    process_files()
    print("--- Обработка завершена ---")
    if sys.stdin and sys.stdin.isatty():
        input("Нажмите Enter для выхода\n")