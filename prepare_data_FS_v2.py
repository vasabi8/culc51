# -*- coding: utf-8 -*-
"""
Підготовка даних фінансової звітності боржників (data_FS.xlsx)
Приведення показників Звіту про фінансові результати до річного виміру

Джерела даних:
  - Source_3BX_3VX_{date}.xlsx (аркуші Source_db_3BX, Source_db_3VX)
  - Dov_3bx_3bv.xlsx (довідники)

Методи приведення до річного виміру (довідник F115):
  1 — Не проводиться (дані вже річні)
  2 — За методом ковзної річної суми: Q007_1 + (Q007_2 - Q007_3)
  3 — З використанням формули: значення × 4 / квартал

Версія 2 (вересень 2026) — виправлення за результатами звірки з Excel-розрахунком
(culcPost351_Source_MovingAnnual_v2.2) та незалежним еталоном, див. prepare_data_FS_v2_changes.md:
  [V2-1] дублікати рядків 3BX на одну дату об'єднуються (перше непорожнє значення поля);
         на кожного боржника — один рядок з найпізнішою датою Q007_1/Q007_2/Q007_3;
  [V2-2] розмір BigMed/Small — за формою звітності F110, за її відсутності — за F059 (F059=3 → BigMed);
  [V2-3] F115=2: згортання пар прибуток/збиток 2090/2095, 2190/2195, 2350/2355 — для всіх розмірів,
         2290/2295 — для BigMed; для малих 2290 = нетто (зі знаком), 2295 = 0;
  [V2-4] у розрахунок потрапляють лише боржники з FMC моделі інтегрального показника (11–15, 21–25);
  [V2-5] колонки 3VX шукаються без урахування пробілів у назві (напр. 'F110 ');
  [V2-6] п.9–10 Додатка 7: квартальна звітність приводиться до річного виміру незалежно від F115 —
         ковзна річна сума за наявності Q007_2 і Q007_3, інакше × 4 / K (у т.ч. F115 = 1 з квартальною датою).

Результат: data_FS_{ReportingDate}__CulcVer_{Date}.xlsx
  Аркуші:
    - FS_Annualized        — Приведені до річного виміру (Баланс + Звіт про фін. результати)
    - Q007_1_За останній ЗП — Дані за останній звітний період (Баланс + Звіт про фін. результати + Рух грош. коштів)
    - Q007_2_річна ФЗ для розрах ІнтегрПоказн — Річні дані для розрахунку (Баланс + Звіт про фін. результати + Рух грош. коштів)
    - Q007_3_За аналогічний ЗП попер року — Дані за аналогічний період попереднього року (Баланс + Звіт про фін. результати + Рух грош. коштів)
    - Info                 — Метадані розрахунку
"""

import pandas as pd
import numpy as np
import os
import sys
import glob
from datetime import datetime

# ============================================================================
# НАЛАШТУВАННЯ
# ============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================================
# ШЛЯХ ДО ФАЙЛУ ДЖЕРЕЛА Source_3BX_3VX_*.xlsx
# Вкажіть повний шлях до файлу або залиште None для автопошуку в папці uploads/
# Приклад: SOURCE_FILE = r"C:\Data\Source_3BX_3VX_2025-07-01.xlsx"
# ============================================================================
SOURCE_FILE = os.path.join(BASE_DIR, "source", "Source_3BX_3VX_2026-01-01.xlsx")   # [V2] назва файлу, що є у source/

DOV_FILE = os.path.join(BASE_DIR, "const_coef", "dov_3bx_3bv.xlsx")

# [V2-4] Коди FMC моделей розрахунку інтегрального показника боржника - юридичної особи
#        (11–15 — великі/середні, 21–25 — малі). Боржники з іншими FMC (42, 43, 44, 45, 90, '#')
#        оцінюються не за Z-моделлю і до data_FS не включаються.
FMC_SCORING = {'11', '12', '13', '14', '15', '21', '22', '23', '24', '25'}
APPLY_FMC_FILTER = True

# [V2-2] Визначення розміру (як у Excel: dov_F110 / dov_F059, колонка «BigMed / Small culc»)
F110_BIGMED = {'1', '2', '5'}   # НП(С)БО 1, консолідована, трансформована з МСФЗ у НП(С)БО 1
F110_SMALL = {'3', '4', '6'}    # НП(С)БО 25, спрощена, трансформована з МСФЗ у НП(С)БО 25
F059_BIGMED = {'1', '3'}        # 3 — мале підприємство, що звітує за НП(С)БО 1


def get_output_path(reporting_date_str: str) -> str:
    """
    Генерація шляху до вихідного файлу.

    Формат: data_FS_{ReportingDate_Quarter}__CulcVer_{YYYY-MM-DD}.xlsx
    Приклад: data_FS_2025-07-01__CulcVer_2026-02-11.xlsx

    Args:
        reporting_date_str: дата ReportingDate_Quarter у форматі 'YYYY-MM-DD'

    Returns:
        str: повний шлях до вихідного файлу
    """
    calc_date = datetime.now().strftime('%Y-%m-%d')
    filename = f"data_FS_{reporting_date_str}__CulcVer_{calc_date}.xlsx"
    return os.path.join(BASE_DIR, "outputs", filename)


def extract_reporting_date_from_source(source_path: str) -> str:
    """
    Витягнення дати ReportingDate_Quarter з назви файлу джерела.

    Наприклад: Source_3BX_3VX_20250701.xlsx → '2025-07-01'

    Args:
        source_path: шлях до файлу джерела

    Returns:
        str: дата у форматі 'YYYY-MM-DD'
    """
    basename = os.path.basename(source_path)
    # Витягнути дату з назви файлу (шаблон: Source_3BX_3VX_YYYYMMDD.xlsx)
    name_no_ext = os.path.splitext(basename)[0]
    parts = name_no_ext.split('_')
    # Останній елемент — дата YYYYMMDD
    date_str = parts[-1]
    try:
        dt = datetime.strptime(date_str, '%Y%m%d')
        return dt.strftime('%Y-%m-%d')
    except ValueError:
        # Якщо не вдалося розпарсити — повернути як є
        return date_str


# Імена колонок Q007
COL_Q007_1 = ('Q007_1 Дата квартальної/ річної фінансової звітності боржника/ групи, за останній звітній період')
COL_Q007_2 = ('Q007_2 Дата річної фінансової звітності боржника/ групи, що використовувалась банком для розрахунку інтегрального показника')
COL_Q007_3 = ('Q007_3 Дата квартальної фінансової звітності боржника/ групи, за аналогічний звітній період попереднього року')
COL_Q026_BX = ("Q026 Належність боржника до групи під спільним контролем/ групи пов'язаних контрагентів")


def _find_q026_column(df_columns) -> str:
    """
    Динамічний пошук колонки Q026 у DataFrame.
    Вирішує проблему різних типів апострофів (U+0027 vs U+2019)
    у назві колонки 'пов'язаних'.

    Args:
        df_columns: колонки DataFrame (df.columns)

    Returns:
        str: фактична назва колонки Q026 або COL_Q026_BX як fallback
    """
    for col in df_columns:
        if isinstance(col, str) and col.startswith('Q026'):
            return col
    return COL_Q026_BX


def find_source_file(base_dir: str) -> str:
    """Автопошук файлу Source_3BX_3VX_*.xlsx."""
    patterns = [
        os.path.join(base_dir, "source", "Source_3BX_3VX_*.xlsx"),
        os.path.join(base_dir, "source_3BX_3VX_*.xlsx"),
    ]
    for pattern in patterns:
        files = glob.glob(pattern)
        if files:
            files.sort(reverse=True)
            return files[0]
    raise FileNotFoundError(
        "Файл Source_3BX_3VX_*.xlsx не знайдено. "
        "Покладіть його у папку uploads/ або задайте SOURCE_FILE вручну."
    )


def load_reference_data(dov_path: str) -> dict:
    """
    Завантаження довідника Dovidnyk_3BX: маппінг A3B → ITEM_id.

    Returns:
        dict з ключами:
          'mapping'         — {A3B_id: ITEM_id}
          'balance_items'   — set ITEM_id, що починаються на '1' (Баланс)
          'income_items'    — set ITEM_id, що починаються на '2' (Звіт про фін. результати)
          'cashflow_items'  — set ITEM_id, що починаються на '3' (Звіт про рух грошових коштів)
          'balance_a3b'     — set A3B-колонок для Балансу
          'income_a3b'      — set A3B-колонок для Звіт про фін. результати
          'cashflow_a3b'    — set A3B-колонок для Руху грошових коштів
    """
    dov = pd.read_excel(dov_path, sheet_name='Dovidnyk_3BX')
    mapping = dict(zip(dov['ID'], dov['ITEM_id'].astype(str)))

    balance_items = {v for v in mapping.values() if v.startswith('1')}
    income_items = {v for v in mapping.values() if v.startswith('2')}
    cashflow_items = {v for v in mapping.values() if v.startswith('3')}

    balance_a3b = {k for k, v in mapping.items() if v.startswith('1')}
    income_a3b = {k for k, v in mapping.items() if v.startswith('2')}
    cashflow_a3b = {k for k, v in mapping.items() if v.startswith('3')}

    print(f"  Довідник: {len(mapping)} записів A3B → ITEM_id")
    print(f"  Баланс (ITEM починається з '1'): "
          f"{len(balance_items)} унікальних показників, {len(balance_a3b)} колонок A3B")
    print(f"  Звіт про фін. результати (ITEM починається з '2'): "
          f"{len(income_items)} унікальних показників, {len(income_a3b)} колонок A3B")
    print(f"  Звіт про рух грошових коштів (ITEM починається з '3'): "
          f"{len(cashflow_items)} унікальних показників, {len(cashflow_a3b)} колонок A3B")

    return {
        'mapping': mapping,
        'balance_items': balance_items,
        'income_items': income_items,
        'cashflow_items': cashflow_items,
        'balance_a3b': balance_a3b,
        'income_a3b': income_a3b,
        'cashflow_a3b': cashflow_a3b,
    }


def load_source_data(source_path: str) -> tuple:
    """
    Завантаження Source_db_3BX та Source_db_3VX.

    Returns:
        (df_bx, df_vx) — DataFrames
    """
    print(f"\n  Завантаження {os.path.basename(source_path)}...")
    df_bx = pd.read_excel(source_path, sheet_name='Source_db_3BX')
    df_vx = pd.read_excel(source_path, sheet_name='Source_db_3VX')

    print(f"    Source_db_3BX: {len(df_bx)} рядків, {len(df_bx.columns)} колонок")
    print(f"    Source_db_3VX: {len(df_vx)} рядків, {len(df_vx.columns)} колонок")

    # Перевірка наявності ключових колонок у Source_db_3BX
    col_q026_found = _find_q026_column(df_bx.columns)
    q026_exists = col_q026_found in df_bx.columns
    if q026_exists:
        non_empty = df_bx[col_q026_found].notna() & (df_bx[col_q026_found] != '')
        print(f"    Q026 знайдено: {non_empty.sum()} непорожніх значень з {len(df_bx)}")
    else:
        print(f"      Q026 не знайдено у Source_db_3BX! Доступні схожі колонки:")
        q026_like = [c for c in df_bx.columns if 'Q026' in str(c) or 'спільним' in str(c)]
        for c in q026_like:
            print(f"        «{c}»")
        if not q026_like:
            print("        (жодної схожої колонки не знайдено)")

    col_rdq_check = 'ReportingDate'
    if col_rdq_check in df_bx.columns:
        non_empty_rdq = df_bx[col_rdq_check].notna().sum()
        print(f"     ReportingDate знайдено: {non_empty_rdq} непорожніх значень")
    else:
        print(f"     ReportingDate не знайдено у Source_db_3BX")

    return df_bx, df_vx


def determine_quarter(q007_1_date) -> int:
    """
    Визначення номера кварталу за датою Q007_1.

    Дата 01.01 → річна (4-й квартал / рік) → quarter=4
    Дата 01.04 → Q1 → quarter=1
    Дата 01.07 → Q2 → quarter=2
    Дата 01.10 → Q3 → quarter=3

    Returns:
        int: номер кварталу (1–4), або 4 якщо дата річна
    """
    if pd.isna(q007_1_date):
        return 4
    dt = pd.to_datetime(q007_1_date)
    month = dt.month
    quarter_map = {1: 4, 4: 1, 7: 2, 10: 3}
    return quarter_map.get(month, 4)


def is_annual_date(q007_1_date) -> bool:
    """Перевірка, чи дата Q007_1 є річною (місяць = 1, тобто 01.01.yyyy)."""
    if pd.isna(q007_1_date):
        return True
    dt = pd.to_datetime(q007_1_date)
    return dt.month == 1


def get_year_quarter_label(q007_1_date, reporting_quarter_date) -> str:
    """
    Формування мітки Year_Quarter (наприклад, '2025_Q1', '2024_Y').

    Args:
        q007_1_date: дата Q007_1 (дата фін. звітності боржника)
        reporting_quarter_date: дата звітного кварталу (ReportingDate)

    Returns:
        str: мітка у форматі 'YYYY_QN' або 'YYYY_Y'
    """
    if pd.notna(q007_1_date):
        dt = pd.to_datetime(q007_1_date)
    elif pd.notna(reporting_quarter_date):
        dt = pd.to_datetime(reporting_quarter_date)
    else:
        return ''

    month = dt.month
    if month == 1:
        year = dt.year - 1
        return f"{year}_Y"
    else:
        quarter_map = {4: 'Q1', 7: 'Q2', 10: 'Q3'}
        q_label = quarter_map.get(month, f'M{month}')
        return f"{dt.year}_{q_label}"


def _norm_code(value) -> str:
    """[V2] Нормалізація коду довідника: 1, 1.0, '1', ' 1 ' → '1'; NaN/None → ''."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    s = str(value).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def map_f059_to_big_small(f059_val, f110_val=None) -> tuple:
    """
    Визначення розміру → (Big_Small_id, Big_Small_name).

    [V2-2] Big_Small_name визначається спершу за формою звітності F110
    (1, 2, 5 → 'Big'; 3, 4, 6 → 'Small'), а за її відсутності — за F059
    (1, 3 → 'Big'; 2 → 'Small'), як у Excel-розрахунку (dov_F110 / dov_F059).
    F059=3 — мале підприємство, що складає звітність за НП(С)БО 1 (форми великих).
    Big_Small_id залишається кодом F059 (далі виводиться як 'F059').
    """
    f059 = _norm_code(f059_val) or '2'
    f110 = _norm_code(f110_val)
    if f110 in F110_BIGMED:
        is_big = True
    elif f110 in F110_SMALL:
        is_big = False
    else:
        is_big = f059 in F059_BIGMED
    try:
        bs_id = int(f059)
    except ValueError:
        bs_id = 2
    return (bs_id, 'Big' if is_big else 'Small')


def _collapse_period_rows(df_bx, date_col, a3b_cols):
    """
    [V2-1] Рядки 3BX одного типу дати (Q007_1 / Q007_2 / Q007_3) → один рядок на ID_KEY.

    У джерелі трапляються дублікати ID_KEY_date: типово «повний» рядок (Баланс + ЗФР + ЗРГК)
    і рядок лише зі Звітом про рух грошових коштів. Вибір першого/останнього з них втрачає дані,
    тому рядки на одну дату об'єднуються: для кожного поля — перше непорожнє значення,
    пріоритет — у рядка з більшою кількістю заповнених полів. Якщо дат кілька — береться
    найпізніша (як LARGE(...;1) в Excel). Порядок боржників — як у джерелі.

    Returns:
        (DataFrame, кількість об'єднаних рядків-дублікатів, кількість відкинутих ранніх дат)
    """
    rows = df_bx[df_bx[date_col].notna()].copy()
    rows['_ord'] = np.arange(len(rows))
    rows['_n_filled'] = rows[a3b_cols].notna().sum(axis=1)
    first_ord = rows.groupby('ID_KEY')['_ord'].min()
    rows = rows.sort_values(['ID_KEY', date_col, '_n_filled'],
                            ascending=[True, True, False], kind='stable')
    n_dup = int(rows.duplicated(['ID_KEY', date_col]).sum())
    merged = rows.groupby(['ID_KEY', date_col], sort=False, as_index=False).first()
    latest = merged.sort_values(date_col, kind='stable').drop_duplicates('ID_KEY', keep='last')
    n_older = len(merged) - len(latest)
    latest = (latest.assign(_ord=latest['ID_KEY'].map(first_ord))
                    .sort_values('_ord', kind='stable')
                    .drop(columns=['_ord', '_n_filled'])
                    .reset_index(drop=True))
    return latest, n_dup, n_older


def _extract_raw_row(bx_row, a3b_cols, a3b_to_item) -> dict:
    """
    Витягнення значень ITEM_id з рядка 3BX для заданих A3B-колонок.

    Args:
        bx_row: рядок DataFrame Source_db_3BX
        a3b_cols: список A3B-колонок для витягування
        a3b_to_item: словник маппінгу A3B → ITEM_id

    Returns:
        dict: {ITEM_id: value}
    """
    result = {}
    for a3b_col in a3b_cols:
        item_id = a3b_to_item.get(a3b_col)
        if item_id is None:
            continue
        val = bx_row.get(a3b_col, 0)
        result[item_id] = float(val) if pd.notna(val) else 0.0
    return result


def _build_meta_row(idx, bx_row, info, q007_1_date, bs_id, bs_name, quarter_num, f115, col_q026=None) -> dict:
    """
    Побудова метаданих рядка (однакові для всіх аркушів).

    Returns:
        dict: метадані боржника
    """
    return {
        '№': idx + 1,
        'ReportingDate': bx_row.get('ReportingDate', ''),
        'NKB': bx_row.get('BANK_NKB', ''),
        'Bank_name': bx_row.get('BANK_NAME', ''),
        'comp_id': bx_row.get('ID_Comp', ''),
        'comp_name': info['comp_name'],
        'Big_Small_id': bs_id,
        'Big_Small_name': bs_name,
        'kved_2n': '',
        'kved_4n': '',
        'Sector': info.get('K115_1', ''),
        'Year_Quarter': get_year_quarter_label(q007_1_date, bx_row.get('ReportingDate')),
        'ReportingDate_FS': q007_1_date if pd.notna(q007_1_date) else '',
        'Q026_НалежністьБоржникаДо_ГСК_ГПК': bx_row.get(col_q026 or COL_Q026_BX, ''),
        'F115_method': f115,
        'F110_form': info.get('F110', ''),
        'Quarter': quarter_num,
    }


def _parse_f115(value, default: int = 1) -> int:
    """
    Безпечне перетворення значення поля F115 у ціле число.

    Обробляє:
      - числові значення (1, 2, 3)
      - рядкові числові значення ('1', '2', '3')
      - значення NaN / None / порожній рядок → default
      - некоректні рядки ('#', 'N/A', текст тощо) → default з попередженням

    Args:
        value: вхідне значення поля F115
        default: значення за замовчуванням (1 = «Не проводиться»)

    Returns:
        int: код методу приведення (1, 2 або 3)
    """
    if pd.isna(value) or value == '' or value is None:
        return default
    try:
        return int(float(str(value).strip()))
    except (ValueError, TypeError):
        print(f"  ️  F115: некоректне значення «{value}» — замінено на {default} (за замовчуванням)")
        return default


def prepare_data_fs(source_path: str, dov_path: str) -> dict:
    """
    Основна функція: підготовка data_FS із приведенням до річного виміру
    та формуванням аркушів Q007_1, Q007_2, Q007_3.

    Args:
        source_path: шлях до Source_3BX_3VX_*.xlsx
        dov_path: шлях до Dov_3bx_3bv.xlsx

    Returns:
        dict з ключами:
          'annualized' — DataFrame (FS_Annualized)
          'q007_1'     — DataFrame (Q007_1)
          'q007_2'     — DataFrame (Q007_2)
          'q007_3'     — DataFrame (Q007_3)
    """
    print("=" * 80)
    print("ПІДГОТОВКА ДАНИХ ФІНАНСОВОЇ ЗВІТНОСТІ (data_FS)")
    print("Приведення показників до річного виміру")
    print("=" * 80)

    # 1. Завантаження довідника
    print("\n1. Завантаження довідника...")
    ref = load_reference_data(dov_path)
    a3b_to_item = ref['mapping']
    balance_a3b = ref['balance_a3b']
    income_a3b = ref['income_a3b']
    cashflow_a3b = ref['cashflow_a3b']

    # 2. Завантаження вхідних даних
    print("\n2. Завантаження вхідних даних...")
    df_bx, df_vx = load_source_data(source_path)

    # Динамічне визначення фактичної назви колонки Q026
    # (вирішує проблему різних типів апострофів: U+0027 vs U+2019)
    col_q026_actual = _find_q026_column(df_bx.columns)
    if col_q026_actual != COL_Q026_BX:
        print(f"  ⚠ Q026: використовується фактична назва колонки з файлу")
    print(f"  Q026 колонка: «{col_q026_actual}»")

    # Визначення колонок A3B, що реально є у Source_db_3BX
    a3b_cols_in_bx = [c for c in df_bx.columns if c.startswith('A3B')]

    a3b_balance_cols = [c for c in a3b_cols_in_bx if c in balance_a3b]
    a3b_income_cols = [c for c in a3b_cols_in_bx if c in income_a3b]
    a3b_cashflow_cols = [c for c in a3b_cols_in_bx if c in cashflow_a3b]
    # Інші A3B (не класифіковані) — також балансові
    a3b_other_cols = [c for c in a3b_cols_in_bx
                      if c not in income_a3b and c not in cashflow_a3b and c not in balance_a3b]

    # Для аркуша FS_Annualized: баланс (+ інші) — без приведення; Звіт про фін. результати — з приведенням
    # ITEM_id: починаються з '1' та '2'
    a3b_annualized_balance = a3b_balance_cols + a3b_other_cols  # без приведення
    a3b_annualized_income = a3b_income_cols  # з приведенням

    # Зворотній маппінг: ITEM_id → A3B колонка (для наявних у джерелі колонок)
    item_to_a3b = {}
    for a3b_col in a3b_cols_in_bx:
        item_id = a3b_to_item.get(a3b_col)
        if item_id and item_id not in item_to_a3b:
            item_to_a3b[item_id] = a3b_col

    # Для аркушів Q007: баланс + Звіт про фін. результати + Рух грошових коштів (сирі дані)
    # ITEM_id: починаються з '1', '2' та '3'
    a3b_q007_all = a3b_balance_cols + a3b_other_cols + a3b_income_cols + a3b_cashflow_cols

    print(f"\n  Колонки A3B у Source_db_3BX: {len(a3b_cols_in_bx)}")
    print(f"    Баланс (prefix '1'): {len(a3b_balance_cols)}")
    print(f"    Звіт про фін. результати (prefix '2', для приведення): {len(a3b_income_cols)}")
    print(f"    Звіт про рух грошових коштів (prefix '3'): {len(a3b_cashflow_cols)}")
    print(f"    Інші (не класифіковані): {len(a3b_other_cols)}")

    # 3. Побудова словника F115 / F059 / K115_1 по ID_KEY з 3VX
    print("\n3. Обробка довідкової інформації з Source_db_3VX...")
    # [V2-5] Пошук колонок 3VX без урахування пробілів на краях (у джерелі — 'F110 ')
    vx_col = {str(c).strip(): c for c in df_vx.columns}

    def _vx(vx_row, name, default=''):
        col = vx_col.get(name)
        return vx_row.get(col, default) if col is not None else default

    vx_info = {}
    for _, vx_row in df_vx.iterrows():
        key = vx_row['ID_KEY']
        vx_info[key] = {
            'F115': _parse_f115(_vx(vx_row, 'F115'), default=1),
            'F059': _vx(vx_row, 'F059', 2),
            'K115_1': _vx(vx_row, 'K115_1', ''),
            'F110': _vx(vx_row, 'F110', ''),
            'FMC': _norm_code(_vx(vx_row, 'FMC', '')),                       # [V2-4]
            'comp_name': _vx(vx_row, 'Q001 Найменування боржника', ''),
            'Q026_vx': _vx(vx_row,
                "Q026 Належність до групи пов'язаних контрагентів", ''),
        }

    f115_counts = pd.Series([v['F115'] for v in vx_info.values()]).value_counts()
    print(f"  Розподіл методів приведення (F115):")
    f115_names = {1: 'Не проводиться', 2: 'Ковзна річна сума', 3: '4/квартал'}
    for code, cnt in f115_counts.items():
        print(f"    {code} ({f115_names.get(code, '?')}): {cnt}")

    # 4. Розподіл рядків 3BX по типах Q007
    print("\n4. Розподіл рядків 3BX за типами дат Q007...")

    # [V2-1] Один рядок на ID_KEY для кожного типу дати; дублікати на одну дату об'єднуються
    print(f"  Рядки з Q007_1 (основний період): {int(df_bx[COL_Q007_1].notna().sum())}")
    print(f"  Рядки з Q007_2 (річні дані): {int(df_bx[COL_Q007_2].notna().sum())}")
    print(f"  Рядки з Q007_3 (аналогічний період попереднього року): {int(df_bx[COL_Q007_3].notna().sum())}")

    rows_q1, dup_q1, old_q1 = _collapse_period_rows(df_bx, COL_Q007_1, a3b_cols_in_bx)
    rows_q2, dup_q2, old_q2 = _collapse_period_rows(df_bx, COL_Q007_2, a3b_cols_in_bx)
    rows_q3, dup_q3, old_q3 = _collapse_period_rows(df_bx, COL_Q007_3, a3b_cols_in_bx)
    print(f"  [V2] Об'єднано рядків-дублікатів на одну дату: "
          f"Q007_1 — {dup_q1}, Q007_2 — {dup_q2}, Q007_3 — {dup_q3}")
    print(f"  [V2] Відкинуто рядків з ранішою датою: "
          f"Q007_1 — {old_q1}, Q007_2 — {old_q2}, Q007_3 — {old_q3}")
    print(f"  [V2] Боржників з Q007_1: {len(rows_q1)}; з Q007_2: {len(rows_q2)}; з Q007_3: {len(rows_q3)}")

    # Індексація Q007_2 та Q007_3 рядків по ID_KEY для швидкого пошуку
    dict_q2 = {r['ID_KEY']: r for _, r in rows_q2.iterrows()}
    dict_q3 = {r['ID_KEY']: r for _, r in rows_q3.iterrows()}

    # 5. Формування вихідних таблиць
    print("\n5. Формування таблиць...")
    annualized_rows = []    # FS_Annualized (Баланс + Звіт про фін. результати приведений)
    q007_1_rows = []        # Q007_1 сирі дані (Баланс + Звіт про фін. результати + Рух грош. коштів)
    q007_2_rows = []        # Q007_2 сирі дані
    q007_3_rows = []        # Q007_3 сирі дані

    total = len(rows_q1)
    skipped = 0
    skipped_fmc = {}          # [V2-4] FMC → кількість виключених боржників
    processed_methods = {1: 0, 2: 0, 3: 0}
    method_overrides = {}     # [V2-6] (F115 звітований → метод застосований) → кількість

    for idx, (_, bx_row) in enumerate(rows_q1.iterrows()):
        if (idx + 1) % 500 == 0 or idx == 0:
            print(f"  Оброблено: {idx + 1}/{total}")

        id_key = bx_row['ID_KEY']
        info = vx_info.get(id_key)
        if info is None:
            skipped += 1
            continue

        # [V2-4] Лише боржники, оцінювані за моделлю інтегрального показника
        if APPLY_FMC_FILTER and info['FMC'] not in FMC_SCORING:
            fmc = info['FMC'] or '(порожньо)'
            skipped_fmc[fmc] = skipped_fmc.get(fmc, 0) + 1
            continue

        f115 = info['F115']
        f059 = info['F059']
        bs_id, bs_name = map_f059_to_big_small(f059, info['F110'])   # [V2-2]
        is_big = bs_name == 'Big'
        q007_1_date = bx_row[COL_Q007_1]
        quarter_num = determine_quarter(q007_1_date)
        annual_flag = is_annual_date(q007_1_date)

        # --- Базові метадані (спільні для всіх аркушів) ---
        # [V2] Нумерація без пропусків (виключені боржники не залишають «дірок» у №)
        meta = _build_meta_row(len(annualized_rows), bx_row, info, q007_1_date, bs_id, bs_name,
                               quarter_num, f115, col_q026_actual)

        # === АРКУШ Q007_1: сирі дані за останній звітний період ===
        q1_row = meta.copy()
        q1_row['Q007_1_date'] = q007_1_date if pd.notna(q007_1_date) else ''
        q1_raw = _extract_raw_row(bx_row, a3b_q007_all, a3b_to_item)
        q1_row.update(q1_raw)
        q007_1_rows.append(q1_row)

        # === АРКУШ Q007_2: річні дані для розрахунку інтегрального показника ===
        row_q2 = dict_q2.get(id_key)
        q2_row = meta.copy()
        if row_q2 is not None:
            q2_date = row_q2[COL_Q007_2]
            q2_row['Q007_2_date'] = q2_date if pd.notna(q2_date) else ''
            q2_raw = _extract_raw_row(row_q2, a3b_q007_all, a3b_to_item)
            q2_row.update(q2_raw)
        else:
            q2_row['Q007_2_date'] = ''
            # Заповнюємо нулями всі ITEM_id колонки (щоб структура була однакова)
            for a3b_col in a3b_q007_all:
                item_id = a3b_to_item.get(a3b_col)
                if item_id is not None:
                    q2_row[item_id] = 0.0
        q007_2_rows.append(q2_row)

        # === АРКУШ Q007_3: дані за аналогічний період попереднього року ===
        row_q3 = dict_q3.get(id_key)
        q3_row = meta.copy()
        if row_q3 is not None:
            q3_date = row_q3[COL_Q007_3]
            q3_row['Q007_3_date'] = q3_date if pd.notna(q3_date) else ''
            q3_raw = _extract_raw_row(row_q3, a3b_q007_all, a3b_to_item)
            q3_row.update(q3_raw)
        else:
            q3_row['Q007_3_date'] = ''
            for a3b_col in a3b_q007_all:
                item_id = a3b_to_item.get(a3b_col)
                if item_id is not None:
                    q3_row[item_id] = 0.0
        q007_3_rows.append(q3_row)

        # === АРКУШ FS_Annualized: приведені до річного виміру ===
        ann_row = meta.copy()

        # Балансові показники (з рядка Q007_1, без приведення)
        for a3b_col in a3b_annualized_balance:
            item_id = a3b_to_item.get(a3b_col)
            if item_id is None:
                continue
            val = bx_row.get(a3b_col, 0)
            ann_row[item_id] = float(val) if pd.notna(val) else 0.0

        # [V2-6] Метод приведення за п.9–10 Додатка 7 Постанови 351:
        #   «під час розрахунку фінансових показників на підставі квартальної звітності здійснюється
        #    … приведення показників форми № 2 до річного виміру» (п.9) — отже для квартальної дати
        #   Q007_1 приведення обов'язкове незалежно від звітованого F115;
        #   ковзна річна сума (п.10), а якщо її неможливо розрахувати через відсутність даних за
        #   попередні періоди — формула × 4 / K. F115 = 3 — × 4 / K, як і раніше.
        #   (v1/v2 до зміни: F115 = 1 → без приведення; F115 = 2 без Q007_2/Q007_3 → без приведення.)
        if annual_flag:
            method = 1
        elif f115 == 3:
            method = 3
        elif row_q2 is not None and row_q3 is not None:
            method = 2
        else:
            method = 3
        if method != f115 and not annual_flag:          # лише квартальні дати (річна — без приведення, як і раніше)
            key_ovr = (f115, method)
            method_overrides[key_ovr] = method_overrides.get(key_ovr, 0) + 1

        # Показники Звіту про фін. результати (приведення до річного виміру)
        for a3b_col in a3b_annualized_income:
            item_id = a3b_to_item.get(a3b_col)
            if item_id is None:
                continue

            val_q1 = float(bx_row.get(a3b_col, 0)) if pd.notna(bx_row.get(a3b_col)) else 0.0

            if method == 1:
                # Метод 1: Не проводиться — дані вже річні (річна дата Q007_1)
                annual_val = val_q1

            elif method == 2:
                # Метод 2: Ковзна річна сума
                # annual = Q007_1 + (Q007_2 - Q007_3)
                if row_q2 is not None and row_q3 is not None:
                    val_q2 = float(row_q2.get(a3b_col, 0)) if pd.notna(row_q2.get(a3b_col)) else 0.0
                    val_q3 = float(row_q3.get(a3b_col, 0)) if pd.notna(row_q3.get(a3b_col)) else 0.0
                    annual_val = val_q1 + (val_q2 - val_q3)
                else:
                    # Якщо немає даних для ковзної суми — залишаємо як є
                    annual_val = val_q1

            elif method == 3:
                # Метод 3: 4 / квартал
                if quarter_num in (1, 2, 3):
                    annual_val = val_q1 * 4 / quarter_num
                else:
                    annual_val = val_q1  # Q4 / річна = без зміни

            else:
                annual_val = val_q1

            ann_row[item_id] = annual_val

        # ==================================================================
        # Корекція парних рахунків прибуток/збиток (Метод 2 — ковзна річна сума)
        # [V2-3] Для ВСІХ розмірів (у v1 — лише F059=1): окремі рахунки
        # розраховуються як нетто-величина з подальшим розподілом:
        #   net = Q007_1(profit-abs(loss)) + (Q007_2(profit-abs(loss)) - Q007_3(profit-abs(loss)))
        #   Якщо net >= 0 → profit = net, loss = 0
        #   Якщо net < 0  → profit = 0, loss = net (без abs())
        # Пара 2290/2295 розподіляється так лише для BigMed; для малих (форма 2-м/2-мс
        # має один рядок фінрезультату до оподаткування) 2290 = net зі знаком, 2295 = 0 —
        # саме 2290 використовують показники MK4, MK7, MK10.
        # ==================================================================
        if method == 2:                                   # [V2-6] було: f115 == 2 and not annual_flag
            paired_accounts = [
                ('2090', '2095'),  # Валовий прибуток / збиток
                ('2190', '2195'),  # Фін. результат від операційної діяльності
                ('2290', '2295'),  # Фін. результат до оподаткування
                ('2350', '2355'),  # Чистий фінансовий результат
            ]

            def _get_a3b_val(source_row, a3b_col):
                """Отримання значення A3B колонки з рядка джерела."""
                if source_row is None or a3b_col is None:
                    return 0.0
                val = source_row.get(a3b_col, 0)
                return float(val) if pd.notna(val) else 0.0

            for profit_code, loss_code in paired_accounts:
                a3b_profit = item_to_a3b.get(profit_code)
                a3b_loss = item_to_a3b.get(loss_code)

                # Якщо обидва рахунки відсутні у джерелі — пропускаємо
                if a3b_profit is None and a3b_loss is None:
                    continue

                # Нетто-значення за кожний період: profit - abs(loss)
                net_q1 = (_get_a3b_val(bx_row, a3b_profit) - abs(_get_a3b_val(bx_row, a3b_loss)))
                net_q2 = (_get_a3b_val(row_q2, a3b_profit) - abs(_get_a3b_val(row_q2, a3b_loss)))
                net_q3 = (_get_a3b_val(row_q3, a3b_profit) - abs(_get_a3b_val(row_q3, a3b_loss)))

                # Ковзна річна сума нетто-величини
                if row_q2 is not None and row_q3 is not None:
                    net_annual = net_q1 + (net_q2 - net_q3)
                else:
                    net_annual = net_q1

                # [V2-3] Малі: фінрезультат до оподаткування — один рядок 2290 зі знаком
                if profit_code == '2290' and not is_big:
                    ann_row[profit_code] = net_annual
                    ann_row[loss_code] = 0.0
                    continue

                # Розподіл: додатне → profit, від'ємне → loss
                if net_annual >= 0:
                    ann_row[profit_code] = net_annual
                    ann_row[loss_code] = 0.0
                else:
                    ann_row[profit_code] = 0.0
                    ann_row[loss_code] = net_annual   # abs()

        annualized_rows.append(ann_row)
        processed_methods[f115] = processed_methods.get(f115, 0) + 1

    print(f"\n  Підсумок обробки:")
    print(f"    Всього оброблено: {len(annualized_rows)}")
    print(f"    Пропущено (немає у 3VX): {skipped}")
    if APPLY_FMC_FILTER:                                                          # [V2-4]
        print(f"    Виключено за FMC (не модель інтегрального показника): "
              f"{sum(skipped_fmc.values())}  {dict(sorted(skipped_fmc.items()))}")
    if method_overrides:                                                         # [V2-6]
        names = {1: 'без приведення', 2: 'ковзна річна сума', 3: '×4/квартал'}
        for (rep_m, app_m), cnt in sorted(method_overrides.items()):
            print(f"    [V2-6] F115={rep_m} → застосовано «{names.get(app_m, app_m)}» "
                  f"(п.9–10, квартальна Q007_1): {cnt}")
    for method, cnt in sorted(processed_methods.items()):
        print(f"    Метод F115={method} ({f115_names.get(method, '?')}): {cnt}")

    # 6. Формування DataFrames
    print("\n6. Формування DataFrames...")

    # Метаколонки (спільні для всіх аркушів)
    meta_cols = [
        '№', 'ReportingDate_Quarter', 'NKB', 'Bank_name', 'comp_id', 'comp_name',
        'Big_Small_id', 'Big_Small_name',
        'kved_2n', 'kved_4n', 'Sector',
        'Year_Quarter', 'ReportingDate_FS',
        'Q026_НалежністьБоржникаДо_ГСК_ГПК',
        'F115_method', 'F110_form', 'Quarter',
    ]

    def sort_key(col_name):
        """Ключ сортування: числові спочатку, комбіновані після."""
        parts = col_name.split('_')
        try:
            return (int(parts[0]), int(parts[1]) if len(parts) > 1 else 0)
        except (ValueError, IndexError):
            return (99999, 0)

    def build_df(rows, extra_date_col=None, filter_prefixes=None):
        """
        Побудова DataFrame з правильним порядком колонок.

        Args:
            rows: список рядків (dict)
            extra_date_col: додаткова колонка дати (Q007_1_date, Q007_2_date, Q007_3_date)
            filter_prefixes: tuple/list префіксів ITEM_id для включення (наприклад, ('1','2','3'))
                             None — включити всі ITEM_id

        Returns:
            pd.DataFrame
        """
        df = pd.DataFrame(rows)

        # Визначення додаткової колонки дати
        date_cols = [extra_date_col] if extra_date_col and extra_date_col in df.columns else []

        # ITEM_id колонки
        all_item_cols = [c for c in df.columns if c not in meta_cols and c not in date_cols]

        if filter_prefixes is not None:
            item_cols = [c for c in all_item_cols
                         if any(c.startswith(p) for p in filter_prefixes)]
        else:
            item_cols = all_item_cols

        item_cols.sort(key=sort_key)

        # Порядок: мета → дата Q007 → ITEM_id
        final_cols = ([c for c in meta_cols if c in df.columns]
                      + date_cols
                      + item_cols)
        df = df.reindex(columns=final_cols, fill_value=0.0)

        # Заповнення NaN нулями для числових колонок
        for col in item_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        return df

    # FS_Annualized: ITEM_id починається з '1' та '2'
    df_annualized = build_df(annualized_rows, filter_prefixes=('1', '2'))

    # Q007_1: ITEM_id починається з '1', '2' та '3'
    df_q007_1 = build_df(q007_1_rows, extra_date_col='Q007_1_date',
                         filter_prefixes=('1', '2', '3'))

    # Q007_2: ITEM_id починається з '1', '2' та '3'
    df_q007_2 = build_df(q007_2_rows, extra_date_col='Q007_2_date',
                         filter_prefixes=('1', '2', '3'))

    # Q007_3: ITEM_id починається з '1', '2' та '3'
    df_q007_3 = build_df(q007_3_rows, extra_date_col='Q007_3_date',
                         filter_prefixes=('1', '2', '3'))

    # Перевірка ідентичності переліку боржників
    assert len(df_annualized) == len(df_q007_1) == len(df_q007_2) == len(df_q007_3), \
        (f"Різна кількість боржників: Annualized={len(df_annualized)}, "
         f"Q007_1={len(df_q007_1)}, Q007_2={len(df_q007_2)}, Q007_3={len(df_q007_3)}")

    print(f"\n  Фінальні таблиці:")
    print(f"    FS_Annualized: {df_annualized.shape[0]} рядків × {df_annualized.shape[1]} колонок")
    print(f"    Q007_1:        {df_q007_1.shape[0]} рядків × {df_q007_1.shape[1]} колонок")
    print(f"    Q007_2:        {df_q007_2.shape[0]} рядків × {df_q007_2.shape[1]} колонок")
    print(f"    Q007_3:        {df_q007_3.shape[0]} рядків × {df_q007_3.shape[1]} колонок")

    return {
        'annualized': df_annualized,
        'q007_1': df_q007_1,
        'q007_2': df_q007_2,
        'q007_3': df_q007_3,
    }


def save_results(dfs: dict, output_path: str):
    """
    Збереження результатів у Excel з кількома аркушами.

    Args:
        dfs: словник з DataFrames (annualized, q007_1, q007_2, q007_3)
        output_path: шлях до вихідного файлу
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 1. FS_Annualized — приведені до річного виміру
        dfs['annualized'].to_excel(writer, sheet_name='FS_Annualized', index=False)

        # 2. Q007_1 — дані за останній звітний період
        dfs['q007_1'].to_excel(
            writer, sheet_name='Q007_1_За останній ЗП', index=False)

        # 3. Q007_2 — річні дані для розрахунку інтегрального показника
        dfs['q007_2'].to_excel(
            writer, sheet_name='Q007_2_річна ФЗ розрах ІнтПок', index=False)

        # 4. Q007_3 — дані за аналогічний період попереднього року
        dfs['q007_3'].to_excel(
            writer, sheet_name='Q007_3_аналогічн ЗП попер року', index=False)

        # 5. Info — метадані
        info_df = pd.DataFrame([{
            'Дата створення': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'Кількість боржників': len(dfs['annualized']),
            'Кількість колонок FS_Annualized': len(dfs['annualized'].columns),
            'Кількість колонок Q007': len(dfs['q007_1'].columns),
            'Версія': 'prepare_data_FS_v2',
            'Опис': ('Фінансова звітність боржників (v2: об\'єднання дублікатів 3BX, розмір за F110/F059, '
                     'згортання прибуток/збиток для всіх розмірів, фільтр FMC 11–25, '
                     'обов\'язкове приведення квартальної звітності за п.9–10). '
                     'FS_Annualized — показники Звіту про фін. результати '
                     'приведені до річного виміру (Баланс + Звіт про фін. результати). '
                     'Q007_1/Q007_2/Q007_3 — сирі дані з Source_db_3BX '
                     '(Баланс + Звіт про фін. результати + Рух грошових коштів).'),
        }])
        info_df.to_excel(writer, sheet_name='Info', index=False)

    print(f"\n  Файл збережено: {output_path}")


# ============================================================================
# ЗАПУСК
# ============================================================================
if __name__ == "__main__":
    print("=" * 80)
    print("ПІДГОТОВКА ДАНИХ ФІНАНСОВОЇ ЗВІТНОСТІ БОРЖНИКІВ")
    print("Приведення показників до річного виміру для Z-моделі скорингу")
    print("=" * 80)

    # Визначення шляхів
    source_path = SOURCE_FILE
    if source_path is None:
        source_path = find_source_file(BASE_DIR)

    # Генерація назви вихідного файлу з ReportingDate_Quarter та датою розрахунку
    reporting_date = extract_reporting_date_from_source(source_path)
    output_file = get_output_path(reporting_date)

    print(f"\nФайл джерела:    {source_path}")
    print(f"Файл довідника:  {DOV_FILE}")
    print(f"Вихідний файл:   {output_file}")

    # Перевірка наявності файлів
    if not os.path.exists(source_path):
        print(f"\n Файл не знайдено: {source_path}")
        sys.exit(1)
    if not os.path.exists(DOV_FILE):
        print(f"\n Файл не знайдено: {DOV_FILE}")
        sys.exit(1)

    try:
        # Основний розрахунок
        dfs = prepare_data_fs(source_path, DOV_FILE)

        # Збереження
        save_results(dfs, output_file)

        # Статистика
        df_result = dfs['annualized']
        print("\n" + "=" * 80)
        print("СТАТИСТИКА")
        print("=" * 80)
        print(f"Всього боржників:  {len(df_result)}")
        print(f"\nРозподіл по банках (NKB):")
        print(df_result['NKB'].value_counts().to_string())
        print(f"\nРозподіл по розміру:")
        print(df_result['Big_Small_name'].value_counts().to_string())
        print(f"\nРозподіл по секторах (Sector):")
        print(df_result['Sector'].value_counts().to_string())
        print(f"\nРозподіл по методу приведення (F115):")
        print(df_result['F115_method'].value_counts().to_string())

        # Статистика Q007_2 / Q007_3 наявності
        df_q2 = dfs['q007_2']
        df_q3 = dfs['q007_3']
        q2_has_data = (df_q2['Q007_2_date'] != '').sum()
        q3_has_data = (df_q3['Q007_3_date'] != '').sum()
        print(f"\nНаявність даних Q007_2 (річна ФЗ): {q2_has_data}/{len(df_q2)}")
        print(f"Наявність даних Q007_3 (аналогічний ЗП попер. року): {q3_has_data}/{len(df_q3)}")

        print("\n Готово! Файл створено: " + os.path.basename(output_file))
        print("   Аркуші: FS_Annualized, Q007_1, Q007_2, Q007_3, Info")

    except FileNotFoundError as e:
        print(f"\n ПОМИЛКА: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n ПОМИЛКА: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)