# -*- coding: utf-8 -*-
"""
Система скорингу боржників (Z-модель)
Постанова НБУ №351 - Розрахунок інтегрального показника та визначення класу боржника

"""

import pandas as pd
import numpy as np
from datetime import datetime
import importlib
import sys
import os

# ============================================================================
# ВЕРСІЯ 2 (debtor_scoring_v2.py) — див. debtor_scoring_v2_changes.md:
#   [V2-A1] чистий борг без обрізання до 0 (K1/MK1 2025–2026, K11/MK11 2019) — модулі *_v2.py
#   [V2-A2] MK17 (2019) «не враховується» для форм 2-мс (F110 = 4), п.8 — calculate_coefficients_2019_v2.py
#   [V2-A3] межа діапазону WOE: верхня межа НЕ входить у діапазон («Ki max <»)
#   [V2-B5] дохід 2220 за модулем — модулі *_v2.py
#   [V2-B6] відсутній коефіцієнт → «не враховується» (None), а не 0
#   [V2-B7] контроль: сектор не визначено / модель не знайдено → колонка «Примітки» та аркуш «Контроль_даних»;
#           резервне визначення сектора за K110_1 з 3VX
#   [V2-B8] Additional_Data: пошук колонок 3VX без пробілів і за кодом (F110, F113, F118, K110);
#           попередження про дублікати ID_KEY у 3VX
# ЗМІНИ 2026-09 (звірка з Постановою 351): пороги класів M/T у data_Const_2025 приведено
# до Додатка 7 (табл. 4–5); маркери перейменовано: MARKER_FIRST_RANGE / MARKER_LAST_RANGE
# (при знаменнику 0 береться X першого/останнього діапазону за п.6/п.7, а не min/max WOE).
# Деталі: NewCulc202609/Зміни_розрахунку_скорингу_202609.md
# ============================================================================
# ============================================================================
# НАЛАШТУВАННЯ ВЕРСІЇ СКОРИНГУ
# ============================================================================
# Доступні версії: '2019', '2025', '2026'
SCORING_VERSION = '2026'

# ============================================================================
# ШЛЯХИ ДО ФАЙЛІВ
# ============================================================================

# Автовизначення базової директорії
# Директорія, де знаходиться скрипт
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Створити outputs якщо не існує
os.makedirs(os.path.join(BASE_DIR, "outputs"), exist_ok=True)

# Файл фінансової звітності боржників
PATH_FS = os.path.join(BASE_DIR, "source", "data_FS_2025-07-01.xlsx")

# Файл джерела Source_3BX_3VX (для додаткових параметрів аркуша Additional_Data)
# Якщо None — аркуш Additional_Data не створюється
PATH_SOURCE_VX = os.path.join(BASE_DIR, "source", "Source_3BX_3VX_2025-07-01.xlsx")

# Файл source_f613 (для аркуша Статистика_2 — аналітика по сумі боргу)
# Якщо None або файл відсутній — аркуш Статистика_2 створюється з повідомленням про відсутність даних
PATH_SOURCE_F613 = os.path.join(BASE_DIR, "source", "source_f613_2025-07-01.xlsx")

# ============================================================================
# КОЛОНКИ ДЛЯ АРКУША Additional_Data (з Source_db_3VX)
# Якщо колонка відсутня у джерелі — залишається порожньою
# ============================================================================
VX_ADDITIONAL_COLUMNS = [
    'F003', 'F059', 'F063', 'F073', 'F074', 'F075', 'F076', 'F082',
    'F110', 'F111', 'F112', 'F113', 'F114', 'F115', 'F116', 'F118',
    'FMC', 'K031', 'K040', 'K061', 'K074', 'K110',
    'K115_1_Ind', 'K115_2_Group', 'K190', 'K021',
    'S190', 'S080_1_Class_Z_model', 'S080_2_factorGroup',
    'S080_3_basedOnAdditionalFactors', 'S080_4_hightRisk',
    'Q036_numberOfPersonnel', 'PD_fact', 'Score_fact', 'Q007_4 Дата реєстрації',
    'Q007_5 Дата визнання дефолту', 'Q007_6 Дата припинення визнання дефолту',
]

# Маппінг: назва колонки у виводі → можливі назви у Source_db_3VX
# (для колонок, що мають іншу назву у джерелі)
VX_COLUMN_ALIASES = {
    'S080_1_Class_Z_model': [
        'S080_1 Код визначеного класу боржника/групи на підставі оцінки фінансового стану (результат Z-моделі)',
        'S080_1_Class_Z_model'],
    'S080_2_factorGroup': [
        'S080_2 Код скорегованого класу боржника з урахуванням фактору належності боржника до групи',
        'S080_2_factorGroup'],
    'S080_3_basedOnAdditionalFactors': [
        'S080_3 Код скорегованого класу боржника на основі додаткових характеристик емітента цінних паперів',
        'S080_3_basedOnAdditionalFactors'],
    'S080_4_hightRisk': [
        'S080_4 Код скоригованого класу боржника з урахуванням фактору належності боржника до групи, факторів високого кредитного ризику та наявності дефолту',
        'S080_4_hightRisk'],
    'K115_1_Ind': ['K115_1_Ind', 'K115_1'],
    'K110': ['K110', 'K110_1'],                                    # [V2-B8]
    'K115_2_Group': ['K115_2_Group', 'K115_2'],
    'Q036_numberOfPersonnel': ['Q036_numberOfPersonnel', 'Q036 Кількість персоналу'],
    'PD_fact': ['PD_fact', 'Коефіцієнт ймовірності дефолту'],
    'Score_fact': ['Score_fact', 'Інтегральний показник'],
}

# Вихідний файл з результатами (версія та дата додаються автоматично)
# Формат: scoring_results_Ver{VERSION}_{ReportingDate_Quarter}__CulcVer_{YYYY-MM-DD}.xlsx
# Приклад: scoring_results_Ver2025_2025-07-01__CulcVer_2026-02-18.xlsx
def extract_reporting_date_quarter(fs_path: str) -> str:
    """Витягнення ReportingDate_Quarter з назви файлу фінансової звітності."""
    import re
    basename = os.path.basename(fs_path)
    match = re.search(r'data_FS_(\d{4}-\d{2}-\d{2})', basename)
    if match:
        return match.group(1)
    return 'unknown'


def get_output_path(version: str, reporting_date_quarter: str = 'unknown') -> str:
    """Генерація шляху до вихідного файлу."""
    date_str = datetime.now().strftime('%Y-%m-%d')
    filename = f"scoring_results_Ver{version}_{reporting_date_quarter}__CulcVer_{date_str}.xlsx"
    return os.path.join(BASE_DIR, "outputs", filename)

# ============================================================================
# КОНФІГУРАЦІЯ ВЕРСІЙ
# ============================================================================
VERSION_CONFIG = {
    '2026': {
        'description': 'Z-модель з 25 коефіцієнтами (Постанова НБУ №351, редакція 2026)',
        'const_file': os.path.join(BASE_DIR, "const_coef", "data_Const_2026.xlsx"),
        'coef_module': os.path.join(BASE_DIR, "const_coef", "calculate_coefficients_2026_v2.py"),
        'num_coefficients': 25,
    },
    '2025': {
        'description': 'Z-модель з 23 коефіцієнтами (Постанова НБУ №351, редакція 2025)',
        'const_file': os.path.join(BASE_DIR, "const_coef", "data_Const_2025.xlsx"),
        'coef_module': os.path.join(BASE_DIR, "const_coef", "calculate_coefficients_2025_v2.py"),
        'num_coefficients': 23,
    },
    '2019': {
        'description': 'Z-модель з 17 коефіцієнтами (Постанова НБУ №351, редакція 2019)',
        'const_file': os.path.join(BASE_DIR, "const_coef", "data_Const_2019.xlsx"),
        'coef_module': os.path.join(BASE_DIR, "const_coef", "calculate_coefficients_2019_v2.py"),
        'num_coefficients': 17,
    },
}

# ============================================================================
# КОНСТАНТИ (спільні для всіх версій)
# ============================================================================
MARKER_FIRST_RANGE = -9999999999  # Група 1 (п.6): при знаменнику 0 береться X ПЕРШОГО діапазону (найменше значення показника)
MARKER_LAST_RANGE = 9999999999   # Група 2 (п.7): при знаменнику 0 береться X ОСТАННЬОГО діапазону (найбільше значення показника)
MARKER_ZERO = None            # Група 0: коефіцієнт не використовується (повертає WOE=0)

# ============================================================================
# ШЛЯХИ ДО МОДУЛІВ РОЗРАХУНКУ (папка const_coef)
# ============================================================================
CONST_COEF_DIR = os.path.join(BASE_DIR, "const_coef")

# Модуль розрахунку фінансових показників для аркуша Result
FINANCIAL_RATIO_MODULE = os.path.join(CONST_COEF_DIR, "calculate_FinancialRatio.py")

# Модуль формування статистики Статистика_1
STATISTICS_1_MODULE = os.path.join(CONST_COEF_DIR, "calculate_statistics_1.py")

# Модуль формування статистики Статистика_2  + з урахуванням файлів 6хх (ф613)
STATISTICS_2_MODULE = os.path.join(CONST_COEF_DIR, "calculate_statistics_2.py")


def _load_module_from_file(module_path: str, module_name: str = None):
    """
    Динамічне завантаження Python-модуля з файлу.

    Args:
        module_path: повний шлях до .py файлу
        module_name: ім'я модуля (за замовчуванням — ім'я файлу без .py)

    Returns:
        module: завантажений модуль
    """
    if module_name is None:
        module_name = os.path.basename(module_path).replace('.py', '')
    import importlib.util
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


class DebtorScoring:
    """
    Клас для розрахунку скорингу боржників за Z-моделлю.
    Підтримує різні версії моделі скорингу через модульну архітектуру.
    """
    
    def __init__(self, fs_path: str, version: str = '2019', source_vx_path: str = None, source_f613_path: str = None):
        """
        Ініціалізація системи скорингу.
        
        Args:
            fs_path: шлях до файлу фінансової звітності
            version: версія моделі скорингу ('2019', '2025' або '2026')
            source_vx_path: шлях до Source_3BX_3VX (для Additional_Data), None — без цього аркуша
            source_f613_path: шлях до source_f613 (для Статистика_2), None — без аналітики по боргу
        """
        self.version = version
        self.fs_path = fs_path
        self.source_vx_path = source_vx_path
        self.source_f613_path = source_f613_path
        
        # Перевірка версії
        if version not in VERSION_CONFIG:
            raise ValueError(f"Невідома версія скорингу: {version}. "
                           f"Доступні версії: {list(VERSION_CONFIG.keys())}")
        
        self.config = VERSION_CONFIG[version]
        self.const_path = self.config['const_file']
        self.num_coefficients = self.config['num_coefficients']
        
        # Завантаження модулів розрахунку
        self._load_coefficient_module()
        self._load_auxiliary_modules()
        
        # Завантаження даних
        self.load_data()
        
        self.results = None
        self.results_main = None
        self.coefficients_df = None
        self.z_components_b_df = None
        self.z_components_bx_df = None
        self.result_df = None
        self.additional_data_df = None
    
    def _load_coefficient_module(self):
        """Динамічне завантаження модуля розрахунку коефіцієнтів."""
        module_path = self.config['coef_module']
        
        if os.path.exists(module_path) and module_path.endswith('.py'):
            self.coef_module = _load_module_from_file(module_path)
            print(f"  - Завантажено модуль коефіцієнтів з файлу: {module_path}")
        else:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            if current_dir not in sys.path:
                sys.path.insert(0, current_dir)
            try:
                import importlib
                self.coef_module = importlib.import_module(module_path)
                print(f"  - Завантажено модуль коефіцієнтів: {module_path}")
            except ImportError as e:
                raise ImportError(f"Не вдалося завантажити модуль {module_path}: {e}")
    
    def _load_auxiliary_modules(self):
        """Завантаження допоміжних модулів (фін. показники, статистика)."""
        # Модуль фінансових показників
        if os.path.exists(FINANCIAL_RATIO_MODULE):
            self.fin_ratio_module = _load_module_from_file(FINANCIAL_RATIO_MODULE)
            print(f"  - Завантажено модуль фінансових показників: {os.path.basename(FINANCIAL_RATIO_MODULE)}")
        else:
            raise FileNotFoundError(
                f"Модуль фінансових показників не знайдено: {FINANCIAL_RATIO_MODULE}")

        # Модуль статистики 1
        if os.path.exists(STATISTICS_1_MODULE):
            self.stats1_module = _load_module_from_file(STATISTICS_1_MODULE)
            print(f"  - Завантажено модуль статистики 1: {os.path.basename(STATISTICS_1_MODULE)}")
        else:
            raise FileNotFoundError(
                f"Модуль статистики 1 не знайдено: {STATISTICS_1_MODULE}")

        # Модуль статистики 2
        if os.path.exists(STATISTICS_2_MODULE):
            self.stats2_module = _load_module_from_file(STATISTICS_2_MODULE)
            print(f"  - Завантажено модуль статистики 2: {os.path.basename(STATISTICS_2_MODULE)}")
        else:
            raise FileNotFoundError(
                f"Модуль статистики 2 не знайдено: {STATISTICS_2_MODULE}")
    
    def load_data(self):
        """Завантаження всіх вхідних даних."""
        print(f"\nЗавантаження даних (версія {self.version})...")
        
        # Фінансова звітність
        self.df_fs = pd.read_excel(self.fs_path)
        self.df_fs = self.df_fs.fillna(0.0)
        
        numeric_cols = [col for col in self.df_fs.columns 
                       if isinstance(col, (int, float)) or 
                       (isinstance(col, str) and col.isdigit())]
        for col in numeric_cols:
            self.df_fs[col] = pd.to_numeric(self.df_fs[col], errors='coerce').fillna(0.0)
        
        print(f"  - Завантажено {len(self.df_fs)} записів фінансової звітності")
        
        if not os.path.exists(self.const_path):
            raise FileNotFoundError(f"Файл констант не знайдено: {self.const_path}")
        
        # Константи
        self.df_bi = pd.read_excel(self.const_path, sheet_name='tab_state_BI')
        self.df_bi = self.df_bi.fillna(0)
        print(f"  - Завантажено {len(self.df_bi)} записів коефіцієнтів BI")
        
        self.df_z = pd.read_excel(self.const_path, sheet_name='tab_state_Z')
        print(f"  - Завантажено {len(self.df_z)} записів діапазонів Z")
        
        self.df_woe = pd.read_excel(self.const_path, sheet_name='tab_state_WOE')
        if 'Группа' in self.df_woe.columns:
            self.df_woe = self.df_woe.rename(columns={'Группа': 'Група'})
        self.df_woe['Група'] = self.df_woe['Група'].str.upper()
        print(f"  - Завантажено {len(self.df_woe)} записів WOE")
        
        self.df_pd = pd.read_excel(self.const_path, sheet_name='tab_state_PD')
        print(f"  - Завантажено {len(self.df_pd)} записів PD")
        
        try:
            self.df_coef_groups = pd.read_excel(self.const_path, sheet_name='tab_coef_groups')
            self.coef_groups = dict(zip(
                self.df_coef_groups['Coef_name'], 
                self.df_coef_groups['Group']
            ))
            print(f"  - Завантажено {len(self.df_coef_groups)} записів груп коефіцієнтів")
        except Exception as e:
            print(f"  ! Попередження: не вдалося завантажити групи коефіцієнтів: {e}")
            self.coef_groups = {}
        
        # Завантаження даних Source_db_3VX для аркуша Additional_Data
        self.vx_data = {}
        self.vx_duplicate_keys = set()   # [V2-B8]
        if self.source_vx_path and os.path.exists(self.source_vx_path):
            try:
                df_vx_source = pd.read_excel(self.source_vx_path, sheet_name='Source_db_3VX')
                vx_available_cols = set(df_vx_source.columns)
                # [V2-B8] пошук без пробілів на краях та за кодом на початку назви
                #         ('F110 ' → F110; "F113 Код наявності ..." → F113)
                vx_stripped = {str(c).strip(): c for c in df_vx_source.columns}
                
                vx_col_map = {}
                for out_col in VX_ADDITIONAL_COLUMNS:
                    aliases = VX_COLUMN_ALIASES.get(out_col, [out_col])
                    found = False
                    for alias in aliases:
                        if alias in vx_available_cols:
                            vx_col_map[out_col] = alias
                            found = True
                            break
                        if alias.strip() in vx_stripped:
                            vx_col_map[out_col] = vx_stripped[alias.strip()]
                            found = True
                            break
                    if not found:
                        by_code = [c for k, c in vx_stripped.items()
                                   if k.startswith(out_col + ' ')]
                        vx_col_map[out_col] = by_code[0] if by_code else None
                
                found_cols = [k for k, v in vx_col_map.items() if v is not None]
                missing_cols = [k for k, v in vx_col_map.items() if v is None]
                print(f"  - Source_db_3VX: {len(df_vx_source)} записів")
                print(f"    Знайдено колонок для Additional_Data: {len(found_cols)}/{len(VX_ADDITIONAL_COLUMNS)}")
                if missing_cols:
                    print(f"    Відсутні колонки (будуть порожні): {', '.join(missing_cols)}")
                
                # [V2-B8] попередження про дублікати ID_KEY у 3VX (залишається останній запис, як у v1)
                dup_keys = df_vx_source['ID_KEY'].astype(str)
                dup_keys = sorted(set(dup_keys[dup_keys.duplicated()]))
                if dup_keys:
                    print(f"    ! Дублікати ID_KEY у Source_db_3VX: {len(dup_keys)} "
                          f"({', '.join(dup_keys[:10])}{' …' if len(dup_keys) > 10 else ''}) — використано останній запис")
                self.vx_duplicate_keys = set(dup_keys)
                
                for _, vx_row in df_vx_source.iterrows():
                    id_key = str(vx_row.get('ID_KEY', ''))
                    row_data = {}
                    for out_col, src_col in vx_col_map.items():
                        if src_col is not None:
                            val = vx_row.get(src_col, '')
                            row_data[out_col] = val if pd.notna(val) else ''
                        else:
                            row_data[out_col] = ''
                    self.vx_data[id_key] = row_data
                    
            except Exception as e:
                print(f"  ! Попередження: не вдалося завантажити Source_db_3VX: {e}")
                self.vx_data = {}
        else:
            if self.source_vx_path:
                print(f"  ! Файл Source_3BX_3VX не знайдено: {self.source_vx_path}")
                print(f"    Аркуш Additional_Data не буде створено.")
            
        # Завантаження даних source_f613 для аркуша Статистика_2
        self.f613_data = None
        if self.source_f613_path and os.path.exists(self.source_f613_path):
            try:
                self.f613_data = pd.read_excel(self.source_f613_path)
                print(f"  - source_f613: {len(self.f613_data)} записів")
                has_cols = all(c in self.f613_data.columns for c in ['ID_KEY', 'UAH_FX', 'Debt'])
                print(f"    Колонки: ID_KEY, UAH_FX, Debt — {'знайдено' if has_cols else 'ВІДСУТНІ!'}")
            except Exception as e:
                print(f"  ! Попередження: не вдалося завантажити source_f613: {e}")
                self.f613_data = None
        else:
            if self.source_f613_path:
                print(f"  ! Файл source_f613 не знайдено: {self.source_f613_path}")
            print(f"    Аркуш Статистика_2 буде створено з повідомленням про відсутність даних.")

        print("Дані успішно завантажено!")
    
    @staticmethod
    def _cluster_by_kved(kved: int) -> str:
        """Галузева група за 2-значним КВЕД (секції A; B, C, F; G; K, L, M, N; інші)."""
        if kved <= 3:
            return 'A'
        elif kved <= 33 or (41 <= kved <= 43):
            return 'M'
        elif 45 <= kved <= 47:
            return 'T'
        elif 64 <= kved <= 82:
            return 'F'
        return 'O'

    def get_cluster_with_note(self, row) -> tuple:
        """
        [V2-B7] Галузева група + примітка, якщо сектор визначено не з поля Sector.
        Порядок: Sector (A/M/T/F/O) → kved_2n → K110_1 з 3VX (напр. 'N0620' → 06) → 'O'.
        """
        sector = row.get('Sector', None)
        if pd.notna(sector) and str(sector).upper() in ['A', 'M', 'T', 'F', 'O']:
            return str(sector).upper(), ''
        try:
            kved = round(float(row.get('kved_2n', 0) or 0))
        except (TypeError, ValueError):
            kved = 0
        if kved > 0:
            return self._cluster_by_kved(kved), f'Sector «{sector}» → група за kved_2n={kved}'
        id_key = f"{row.get('NKB', '')}_{row.get('comp_id', '')}"
        k110 = str(self.vx_data.get(id_key, {}).get('K110', '') or '')
        digits = ''.join(ch for ch in k110 if ch.isdigit())[:2]
        if digits and int(digits) > 0:
            return self._cluster_by_kved(int(digits)), f'Sector «{sector}» → група за K110_1={k110}'
        return 'O', f'Sector «{sector}» не визначено, КВЕД відсутній → група O за замовчуванням'

    def get_cluster(self, row) -> str:
        """Визначення галузевої групи."""
        sector = row.get('Sector', None)
        if pd.notna(sector) and str(sector).upper() in ['A', 'M', 'T', 'F', 'O']:
            return str(sector).upper()
        
        kved = row.get('kved_2n', 0)
        try:
            kved = round(float(kved)) if pd.notna(kved) else 0
        except:
            return 'O'
            
        if kved == 0:
            return 'O'
        elif kved <= 3:
            return 'A'
        elif kved <= 33 or (41 <= kved <= 43):
            return 'M'
        elif 45 <= kved <= 47:
            return 'T'
        elif 64 <= kved <= 82:
            return 'F'
        else:
            return 'O'
    
    def get_report_type(self, big_small: str) -> int:
        """Визначення типу звітності."""
        if big_small == 'Big':
            return 1
        else:
            return 2
    
    def get_woe_value(self, koef: float, zvit: int, cluster: str, x_name: str) -> float:
        """Отримання WOE-значення для коефіцієнта."""
        if koef is None or koef is MARKER_ZERO:
            return 0.0
        
        subset = self.df_woe[
            (self.df_woe['Звітність'] == zvit) & 
            (self.df_woe['Група'] == cluster.upper()) &
            (self.df_woe['Xname'] == x_name)
        ].sort_values('min')
        
        if len(subset) == 0:
            return 0.0
        
        if koef == MARKER_FIRST_RANGE:
            return float(subset.iloc[0]['value'])
        
        if koef == MARKER_LAST_RANGE:
            return float(subset.iloc[-1]['value'])
        
        # [V2-A3] Колонка 'min' містить ВЕРХНЮ межу діапазону, яка за Постановою до нього не входить
        #         («Максимальне значення діапазону (Ki max <)»); у v1 було koef <= межа.
        for _, row in subset.iterrows():
            if koef < row['min']:
                return float(row['value'])
        
        return float(subset.iloc[-1]['value'])
    
    def calculate_z_score(self, x_values: dict, zvit: int, cluster: str) -> tuple:
        """Розрахунок інтегрального показника Z та його складових."""
        row = self.df_bi[
            (self.df_bi['Форма звіту const'] == zvit) & 
            (self.df_bi['Фін клас const'] == cluster.upper())
        ]
        
        if len(row) == 0:
            # [V2-B7] модель для (форма, група) відсутня — позначається у «Примітки»
            return 0.0, {'_note': f'модель не знайдено для форми {zvit}, групи {cluster} → Z = 0'}
        
        row = row.iloc[0]
        b0 = float(row['b0'])
        z = b0
        
        z_components = {'b0': b0}
        
        for i in range(1, self.num_coefficients + 1):
            b_val = float(row[f'b{i}']) if pd.notna(row.get(f'b{i}', np.nan)) else 0.0
            x_val = float(x_values.get(f'X{i}', 0))
            component = b_val * x_val
            z += component
            z_components[f'b{i}'] = b_val
            z_components[f'b{i}*X{i}'] = round(component, 6)
        
        return z, z_components
    
    def get_financial_class(self, z_score: float, zvit: int, cluster: str) -> int:
        """Визначення фінансового класу за Z-score."""
        subset = self.df_z[
            (self.df_z['Форма звіту Z'] == zvit) & 
            (self.df_z['Фін клас Z'] == cluster.upper())
        ].sort_values('Значення Z')
        
        for _, row in subset.iterrows():
            if z_score < row['Значення Z']:
                return int(row['Фін клас П'])
        
        return 9
    
    def get_pd(self, fin_class: int, cluster: str) -> float:
        """Отримання PD за класом."""
        row = self.df_pd[
            (self.df_pd['Група'] == cluster.upper()) & 
            (self.df_pd['Клас боржника - юридичної особи'] == fin_class)
        ]
        
        if len(row) == 0:
            return 1.0
        
        row = row.iloc[0]
        return (float(row['Більше або дорівнює']) + float(row['Менше'])) / 2
    
    def process_all(self) -> pd.DataFrame:
        """Обробка всіх боржників."""
        results = []
        coefficients_list = []
        z_components_b_list = []
        z_components_bx_list = []
        result_list = []
        additional_data_list = []
        total = len(self.df_fs)
        
        print(f"\nОбробка {total} боржників (версія {self.version})...")
        
        for idx, row in self.df_fs.iterrows():
            if (idx + 1) % 100 == 0 or idx == 0:
                print(f"  Оброблено: {idx + 1}/{total}")
            
            zvit = self.get_report_type(row.get('Big_Small_name', 'Small'))
            cluster, cluster_note = self.get_cluster_with_note(row)       # [V2-B7]
            notes = [cluster_note] if cluster_note else []
            
            base_cols = {
                'NKB': row.get('NKB', ''),
                'Bank_name': row.get('Bank_name', ''),
                'comp_id': row.get('comp_id', ''),
                'comp_name': row.get('comp_name', ''),
                'Big_Small_id': row.get('Big_Small_id', ''),
                'Big_Small_name': row.get('Big_Small_name', ''),
                'kved_2n': row.get('kved_2n', ''),
                'kved_4n': row.get('kved_4n', ''),
                'Sector': row.get('Sector', ''),
                'Year_Quarter': row.get('Year_Quarter', ''),
                'ReportingDate_FS': row.get('ReportingDate_FS', ''),
                'Q026_НалежністьБоржникаДо_ГСК_ГПК': row.get('Q026_НалежністьБоржникаДо_ГСК_ГПК', ''),
            }
            
            # Розрахунок коефіцієнтів через модуль версії
            coeffs = self.coef_module.calculate_coefficients(row, zvit, self.coef_groups)
            
            coef_row = base_cols.copy()
            coef_row['Звітність'] = zvit
            coef_row['Галузева група'] = cluster
            for i in range(1, self.num_coefficients + 1):
                coef_row[f'K{i}'] = coeffs.get(f'K{i}', MARKER_ZERO)      # [V2-B6]
            coefficients_list.append(coef_row)
            
            # Перетворення у WOE
            x_values = {}
            for i in range(1, self.num_coefficients + 1):
                x_values[f'X{i}'] = self.get_woe_value(
                    coeffs.get(f'K{i}', MARKER_ZERO), zvit, cluster, f'X{i}'   # [V2-B6]
                )
            
            # Розрахунок Z та складових
            z_score, z_components = self.calculate_z_score(x_values, zvit, cluster)
            if '_note' in z_components:                                 # [V2-B7]
                notes.append(z_components.pop('_note'))
            id_key_row = f"{row.get('NKB', '')}_{row.get('comp_id', '')}"
            if id_key_row in self.vx_duplicate_keys:                    # [V2-B8]
                notes.append('дублікат ID_KEY у Source_db_3VX')
            fin_class = self.get_financial_class(z_score, zvit, cluster)
            pd_value = self.get_pd(fin_class, cluster)
            
            z_b_row = base_cols.copy()
            z_b_row['Звітність'] = zvit
            z_b_row['Галузева група'] = cluster
            z_b_row['b0'] = z_components.get('b0', 0)
            for i in range(1, self.num_coefficients + 1):
                z_b_row[f'b{i}'] = z_components.get(f'b{i}', 0)
            z_components_b_list.append(z_b_row)
            
            z_bx_row = base_cols.copy()
            z_bx_row['Звітність'] = zvit
            z_bx_row['Галузева група'] = cluster
            z_bx_row['b0'] = z_components.get('b0', 0)
            for i in range(1, self.num_coefficients + 1):
                z_bx_row[f'b{i}*X{i}'] = z_components.get(f'b{i}*X{i}', 0)
            z_bx_row['SCORE'] = round(z_score, 6)
            z_components_bx_list.append(z_bx_row)
            
            b0_value = z_components.get('b0', 0.0)
            
            result = base_cols.copy()
            result['Звітність'] = zvit
            result['Галузева група'] = cluster
            for i in range(1, self.num_coefficients + 1):
                result[f'X{i}'] = round(x_values[f'X{i}'], 6) if x_values[f'X{i}'] != 0 else 0
            result['b0'] = round(b0_value, 6)
            result['SCORE'] = round(z_score, 6)
            result['Фінансовий клас'] = fin_class
            result['PD'] = round(pd_value, 6)
            result['Примітки'] = '; '.join(notes)                        # [V2-B7]
            results.append(result)
            
            # Розрахунок фінансових показників через зовнішній модуль
            result_row = self.fin_ratio_module.calculate_financial_ratios(
                row, zvit, idx, z_score, fin_class, pd_value
            )
            result_list.append(result_row)
            
            # Формування рядка для аркуша Additional_Data
            if self.vx_data:
                id_key = f"{row.get('NKB', '')}_{row.get('comp_id', '')}"
                add_row = base_cols.copy()
                add_row['ID_KEY'] = id_key
                add_row['Звітність'] = zvit
                add_row['Галузева група'] = cluster
                vx_row_data = self.vx_data.get(id_key, {})
                for col_name in VX_ADDITIONAL_COLUMNS:
                    add_row[col_name] = vx_row_data.get(col_name, '')
                additional_data_list.append(add_row)
        
        self.results = pd.DataFrame(results)
        self.coefficients_df = pd.DataFrame(coefficients_list)
        self.z_components_b_df = pd.DataFrame(z_components_b_list)
        self.z_components_bx_df = pd.DataFrame(z_components_bx_list)
        self.result_df = pd.DataFrame(result_list)
        if additional_data_list:
            self.additional_data_df = pd.DataFrame(additional_data_list)
        
        n_notes = int((self.results['Примітки'] != '').sum()) if len(self.results) else 0   # [V2-B7]
        print(f"\nОбробку завершено!")
        print(f"  Рядків із примітками контролю: {n_notes}")
        print(f"  Всього оброблено: {len(self.results)} записів")
        
        print("\nРозподіл по фінансових класах:")
        class_counts = self.results['Фінансовий клас'].value_counts().sort_index()
        for cls, count in class_counts.items():
            print(f"  Клас {cls}: {count} ({count/len(self.results)*100:.1f}%)")
        
        return self.results
    
    def save_results(self, output_path: str = None):
        """
        Збереження результатів у Excel.
        
        Порядок аркушів: Result, Статистика_1, Статистика_2, Коефіцієнти,
        WOE_Результати, Z_Складові_b, Z_Складові_b_x_score, Інформація
        """
        if output_path is None:
            reporting_date_quarter = extract_reporting_date_quarter(self.fs_path)
            output_path = get_output_path(self.version, reporting_date_quarter)
        
        if self.results is not None:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 1. Аркуш Result
                if self.result_df is not None:
                    self.result_df.to_excel(writer, sheet_name='Result', index=False)
                
                # 2. Статистика_1 (розподіл по кількості боржників)
                stats = self.stats1_module.create_statistics_1(self.results)
                start_row = 0
                for stat_name, stat_df in stats.items():
                    header_df = pd.DataFrame([[stat_name]], columns=[''])
                    header_df.to_excel(writer, sheet_name='Статистика_1', startrow=start_row, index=False, header=False)
                    start_row += 1
                    stat_df.to_excel(writer, sheet_name='Статистика_1', startrow=start_row, index=False)
                    start_row += len(stat_df) + 3
                
                # 3. Статистика_2 (аналітика по сумі боргу з f613)
                stats_2 = self.stats2_module.create_statistics_2(
                    self.results, self.result_df, self.f613_data
                )
                if stats_2 is not None:
                    start_row_2 = 0
                    for stat_name, stat_df in stats_2.items():
                        header_df = pd.DataFrame([[stat_name]], columns=[''])
                        header_df.to_excel(writer, sheet_name='Статистика_2', startrow=start_row_2, index=False, header=False)
                        start_row_2 += 1
                        stat_df.to_excel(writer, sheet_name='Статистика_2', startrow=start_row_2, index=False)
                        start_row_2 += len(stat_df) + 3
                else:
                    no_data_df = pd.DataFrame([{
                        'Повідомлення': 'Дані source_f613 відсутні або не вдалося з\'єднати з результатами скорингу. '
                                        'Аркуш Статистика_2 потребує файл source_f613 з колонками ID_KEY, UAH_FX, Debt.'
                    }])
                    no_data_df.to_excel(writer, sheet_name='Статистика_2', index=False)
                
                # 4. Коефіцієнти
                if self.coefficients_df is not None:
                    self.coefficients_df.to_excel(writer, sheet_name='Коефіцієнти', index=False)
                
                # 5. WOE_Результати
                self.results.to_excel(writer, sheet_name='WOE_Результати', index=False)
                
                # 6. Z_Складові_b
                if self.z_components_b_df is not None:
                    self.z_components_b_df.to_excel(writer, sheet_name='Z_Складові_b', index=False)
                
                # 7. Z_Складові_b_x_score
                if self.z_components_bx_df is not None:
                    self.z_components_bx_df.to_excel(writer, sheet_name='Z_Складові_b_x_score', index=False)
                
                # 8. Additional_Data
                if self.additional_data_df is not None:
                    self.additional_data_df.to_excel(writer, sheet_name='Additional_Data', index=False)
                
                # [V2-B7] 8a. Контроль_даних — рядки з примітками
                ctrl = self.results[self.results['Примітки'] != '']
                ctrl_cols = ['NKB', 'Bank_name', 'comp_id', 'comp_name', 'Sector', 'Звітність',
                             'Галузева група', 'SCORE', 'Фінансовий клас', 'Примітки']
                (ctrl[[c for c in ctrl_cols if c in ctrl.columns]] if len(ctrl) else
                 pd.DataFrame([{'Примітки': 'Зауважень немає'}])).to_excel(
                    writer, sheet_name='Контроль_даних', index=False)
                
                # 9. Інформація
                version_info = pd.DataFrame([{
                    'Версія': self.version,
                    'Скрипт': 'debtor_scoring_v2',
                    'Опис': self.config['description'],
                    'Файл констант': self.const_path,
                    'Дата розрахунку': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'Всього записів': len(self.results),
                }])
                version_info.to_excel(writer, sheet_name='Інформація', index=False)
            
            print(f"\nРезультати збережено: {output_path}")


def print_available_versions():
    """Виведення інформації про доступні версії."""
    print("\nДоступні версії скорингу:")
    print("-" * 60)
    for ver, config in VERSION_CONFIG.items():
        print(f"  {ver}: {config['description']}")
        print(f"       Коефіцієнтів: {config['num_coefficients']}")
        print(f"       Файл констант: {config['const_file']}")
        print()


# ============================================================================
# ЗАПУСК
# ============================================================================
if __name__ == "__main__":
    print("=" * 80)
    print("СИСТЕМА СКОРИНГУ БОРЖНИКІВ (Z-МОДЕЛЬ)")
    print("Постанова НБУ №351")
    print("=" * 80)
    
    REPORTING_DATE_QUARTER = extract_reporting_date_quarter(PATH_FS)
    PATH_OUTPUT = get_output_path(SCORING_VERSION, REPORTING_DATE_QUARTER)
    
    print("=" * 80)
    print(f"ПОТОЧНА ВЕРСІЯ: {SCORING_VERSION}")
    print(f"Файл фінансової звітності: {PATH_FS}")
    print(f"Файл констант: {VERSION_CONFIG[SCORING_VERSION]['const_file']}")
    print(f"Файл Source_3BX_3VX: {PATH_SOURCE_VX}")
    print(f"Файл source_f613: {PATH_SOURCE_F613}")
    print(f"Вихідний файл: {PATH_OUTPUT}")
    print("=" * 80)
    
    try:
        scorer = DebtorScoring(PATH_FS, version=SCORING_VERSION, source_vx_path=PATH_SOURCE_VX, source_f613_path=PATH_SOURCE_F613)
        results = scorer.process_all()
        scorer.save_results(PATH_OUTPUT)
        
    except FileNotFoundError as e:
        print(f"\n ПОМИЛКА: {e}")
        print("Перевірте шляхи до файлів у налаштуваннях.")
    except Exception as e:
        print(f"\n ПОМИЛКА: {e}")
        import traceback
        traceback.print_exc()
