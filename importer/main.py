import pandas as pd
import pyodbc
import os
import sys
import warnings
import ctypes
from datetime import datetime
from copy import copy
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.worksheet.datavalidation import DataValidation
from dotenv import load_dotenv

# ==================== ЗАРЕЖДАНЕ НА .ENV ====================
env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
load_dotenv(dotenv_path=env_path)

# ==================== КОНФИГУРАЦИЯ ОТ .ENV ====================
CONFIG = {
    'server': os.getenv('DB_SERVER', '.'),
    'database': os.getenv('DB_DATABASE', ''),
    'table_name': os.getenv('DB_TABLE', 'Items'),
    'excel_file': os.getenv('EXCEL_FILE', None),
    'sheet_name': int(os.getenv('EXCEL_SHEET', '0')),
    'skiprows': int(os.getenv('EXCEL_SKIPROWS', '0')),
    'trusted_connection': os.getenv('DB_TRUSTED_CONNECTION', 'True').lower() == 'true',
    'login_timeout': int(os.getenv('DB_TIMEOUT', '15'))
}

EXPECTED_COLUMNS = ['Код', 'Стока', 'Мярка', 'Цена']

class ExcelSQLManager:
    def log(self, message):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{timestamp}] {message}")
    
    def check_odbc_driver(self):
        """Проверява дали е инсталиран необходимият ODBC драйвер"""
        drivers = pyodbc.drivers()
        required_driver = "ODBC Driver 17 for SQL Server"
        
        if required_driver not in drivers:
            print("\n" + "!"*60)
            print("ГРЕШКА: Не е инсталиран необходимият ODBC драйвер!")
            print("!"*60)
            print(f"\nОчакван: {required_driver}")
            print("\nИнсталирани драйвери на тази машина:")
            for i, driver in enumerate(drivers, 1):
                print(f"  {i}. {driver}")
            print("\nМоля инсталирайте: Microsoft ODBC Driver 17 for SQL Server")
            print("Линк за изтегляне:")
            print("https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server")
            print("\nСлед инсталацията рестартирайте програмата.")
            input("\nНатиснете Enter за изход...")
            return False
        
        self.log(f"✓ Намерен драйвер: {required_driver}")
        return True
    
    def get_available_databases(self):
        """Връща списък с наличните бази данни на сървъра"""
        try:
            # Свързваме се без да посочваме конкретна база (към master)
            conn_str = (f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                       f"SERVER={CONFIG['server']};"
                       f"Trusted_Connection=yes;"
                       f"Login Timeout={CONFIG['login_timeout']};")
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sys.databases WHERE state = 0 AND name NOT IN ('master', 'tempdb', 'model', 'msdb') ORDER BY name")
            databases = [row[0] for row in cursor.fetchall()]
            conn.close()
            return databases
        except Exception as e:
            self.log(f"Не може да се извлече списък с базите: {e}")
            return []
    
    def prompt_database_selection(self):
        """Показва меню за избор на база данни при грешка"""
        databases = self.get_available_databases()
        
        if not databases:
            self.log("✗ Не са намерени достъпни бази данни или липсва връзка със сървъра")
            return False
        
        print("\n" + "="*60)
        print("       НАЛИЧНИ БАЗИ ДАННИ НА СЪРВЪРА")
        print("="*60)
        for i, db in enumerate(databases, 1):
            marker = " <-- ТЕКУЩА" if db == CONFIG['database'] else ""
            print(f"{i:2}. {db}{marker}")
        print("="*60)
        print("0. Отказ (обратно към менюто)")
        print("-"*60)
        
        while True:
            choice = input(f"Изберете база данни (0-{len(databases)}): ").strip()
            if choice == '0':
                return False
            try:
                idx = int(choice) - 1
                if 0 <= idx < len(databases):
                    old_db = CONFIG['database']
                    CONFIG['database'] = databases[idx]
                    self.log(f"✓ Сменена база данни: {old_db} -> {CONFIG['database']}")
                    return True
                else:
                    print("Невалиден номер!")
            except ValueError:
                # Проверка дали е въведено име директно
                if choice in databases:
                    old_db = CONFIG['database']
                    CONFIG['database'] = choice
                    self.log(f"✓ Сменена база данни: {old_db} -> {CONFIG['database']}")
                    return True
                else:
                    print("Моля въведете валиден номер или име от списъка!")
    
    def ensure_database_selected(self):
        """Гарантира, че има избрана база данни преди операция"""
        if str(CONFIG.get('database', '')).strip():
            return True
        self.log("⚠ Името на базата данни е празно.")
        self.log("  Изберете база данни от списъка:")
        return self.prompt_database_selection()
    
    def check_table_exists(self, conn, table_name=None):
        """Проверява дали таблицата съществува в текущата база"""
        try:
            table_to_check = table_name or CONFIG['table_name']
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES 
                WHERE TABLE_NAME = ? AND TABLE_TYPE = 'BASE TABLE'
            """, (table_to_check,))
            exists = cursor.fetchone()[0] > 0
            cursor.close()
            return exists
        except:
            return False
    
    def handle_connection_error(self, error):
        """Обработва грешки при свързване и предлага избор на база при нужда"""
        error_msg = str(error).lower()
        error_str = str(error)
        
        # Проверка за грешки свързани с несъществуваща база или липса на права
        if any(x in error_msg for x in ["cannot open database", "4060", "login failed", "28000", "недостъпна"]):
            self.log(f"✗ Неуспешно свързване към база '{CONFIG['database']}'")
            self.log(f"  Грешка: {error_str}")
            print("\nВъзможни причини:")
            print("  - Базата данни не съществува")
            print("  - Нямате права за достъп")
            print("  - Грешно име на базата")
            
            if self.prompt_database_selection():
                return True  # Потребителят избра нова база, може да опитаме пак
            else:
                return False  # Отказ
        else:
            # Други грешки (мрежа, сървър и т.н.)
            self.log(f"✗ Грешка при свързване: {error_str}")
            if "network" in error_msg or "server" in error_msg:
                print("\nПроблем с връзката към сървъра.")
                print(f"Проверете дали SQL Server '{CONFIG['server']}' е достъпен.")
            return False
    
    def _with_tk_dialog(self, func):
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update_idletasks()
        try:
            return func(root)
        finally:
            root.destroy()
            self.bring_console_to_front()
    
    def bring_console_to_front(self):
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                ctypes.windll.user32.ShowWindow(hwnd, 1)
        except Exception:
            pass
    
    def transliterate(self, text):
        if pd.isna(text) or not str(text).strip():
            return ''
        trans_map = {
            'Щ': 'Sht', 'Ш': 'Sh', 'Ч': 'Ch', 'Ж': 'Zh', 'Ц': 'Ts', 'Ю': 'Yu', 'Я': 'Qa',
            'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'З': 'Z', 'И': 'I',
            'Й': 'Y', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O', 'П': 'P', 'Р': 'R',
            'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'H', 'Ъ': 'A', 'Ь': 'Y',
            'щ': 'sht', 'ш': 'sh', 'ч': 'ch', 'ж': 'zh', 'ц': 'ts', 'ю': 'yu', 'я': 'q',
            'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'з': 'z', 'и': 'i',
            'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r',
            'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ъ': 'a', 'ь': 'y',
        }
        return ''.join(trans_map.get(char, char) for char in str(text))
    
    def get_connection_string(self):
        driver = "ODBC Driver 17 for SQL Server"
        return (f"DRIVER={{{driver}}};"
                f"SERVER={CONFIG['server']};"
                f"DATABASE={CONFIG['database']};"
                f"Trusted_Connection=yes;"
                f"Login Timeout={CONFIG['login_timeout']};")
    
    def auto_adjust_column_width(self, worksheet):
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def format_header_bold(self, worksheet):
        for cell in worksheet[1]:
            new_font = copy(cell.font)
            new_font.bold = True
            cell.font = new_font
    
    def parse_id_value(self, value):
        if pd.isna(value):
            return None
        try:
            return int(float(value))
        except (ValueError, TypeError):
            pass
        str_val = str(value).strip()
        if ' - ' in str_val:
            try:
                return int(str_val.split(' - ')[0])
            except ValueError:
                pass
        return None
    
    def add_dropdown_validation(self, worksheet, column_letter, source_sheet, source_column, start_row, end_row, allow_blank=True):
        max_source_row = 1000
        formula = f"='{source_sheet}'!${source_column}$2:${source_column}${max_source_row}"
        dv = DataValidation(type="list", formula1=formula, allow_blank=allow_blank)
        dv.error = 'Моля изберете стойност от списъка'
        dv.errorTitle = 'Невалидна стойност'
        dv.prompt = f'Изберете от {source_sheet}'
        dv.promptTitle = 'Справочник'
        cell_range = f'{column_letter}{start_row}:{column_letter}{end_row}'
        dv.add(cell_range)
        worksheet.add_data_validation(dv)
    
    def connect_with_fallback(self):
        """Опитва се да се свърже, при неуспех предлага избор на база"""
        if not self.ensure_database_selected():
            return None
        
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                conn = pyodbc.connect(self.get_connection_string())
                # Проверка дали таблицата съществува
                if not self.check_table_exists(conn):
                    self.log(f"⚠ Таблицата '{CONFIG['table_name']}' не съществува в база '{CONFIG['database']}'!")
                    conn.close()
                    if not self.prompt_database_selection():
                        return None
                    continue  # Опитваме пак с нова база
                return conn
            except pyodbc.Error as e:
                if attempt < max_attempts - 1:
                    if self.handle_connection_error(e):
                        continue  # Потребителят избра нова база, опитваме пак
                    else:
                        return None
                else:
                    self.log("✗ Неуспешно свързване след няколко опита")
                    return None
            except Exception as e:
                self.log(f"✗ Неочаквана грешка: {e}")
                return None
    
    def export_items_to_excel(self):
        if not self.ensure_database_selected():
            self.log("Експортът е отменен: няма избрана база данни.")
            return
        
        initial_dir = os.path.dirname(CONFIG['excel_file']) if CONFIG['excel_file'] and os.path.exists(CONFIG['excel_file']) else os.getcwd()
        initial_name = "invoice_pro_items_export.xlsx"
        export_file = self._with_tk_dialog(lambda r: filedialog.asksaveasfilename(
            title="Запази Excel файл като",
            initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Excel файлове", "*.xlsx"), ("Всички файлове", "*.*")],
            parent=r
        ))
        if not export_file:
            self.log("Експортът е отменен от потребителя.")
            return
        
        self.log(f"=== ЕКСПОРТ ОТ SQL КЪМ EXCEL ===")
        self.log(f"Сървър: {CONFIG['server']}")
        self.log(f"База: {CONFIG['database']}")
        self.log(f"Таблица: {CONFIG['table_name']}")
        
        if os.path.exists(export_file):
            try:
                os.remove(export_file)
            except Exception as e:
                self._with_tk_dialog(lambda r: messagebox.showerror("Грешка", 
                    f"Файлът е отворен в друга програма.\nМоля затворете го.", parent=r))
                return
        
        # Свързване с fallback
        conn = self.connect_with_fallback()
        if not conn:
            return
        
        try:
            cursor = conn.cursor()
            
            # Проверка дали колоните съществуват
            try:
                cursor.execute(f"SELECT TOP 1 * FROM [dbo].[{CONFIG['table_name']}]")
                cursor.fetchone()
            except pyodbc.Error as e:
                self.log(f"✗ Грешка при достъп до таблица: {e}")
                return
            
            query_items = f"""
            SELECT [Code] as 'Код', [Name] as 'Стока', [Measure] as 'Мярка', 
                   [SalePrice] as 'Цена', [VatRateID] as 'ДДС ID',
                   [GroupID] as 'Група ID', [StatusID] as 'Статус ID', 
                   [VatTermID] as 'ДДС Срок ID'
            FROM [dbo].[{CONFIG['table_name']}]
            WHERE [Visible] = 1
            ORDER BY [Name]
            """
            
            query_vatrates = """SELECT [VatRateID] as 'ДДС ID', [Code] as 'Код',
                [Description] as 'Описание', [Rate] as 'Стойност', [TypeIdentifier] as 'Тип'
                FROM [dbo].[VatRates] ORDER BY [VatRateID]"""
            
            query_itemgroups = """SELECT [GroupID] as 'Група ID', [Code] as 'Код', [Name] as 'Име'
                FROM [dbo].[ItemGroups] ORDER BY [GroupID]"""
            
            query_status = """SELECT [StatusID] as 'Статус ID', [Name] as 'Име'
                FROM [dbo].[Status] ORDER BY [StatusID]"""
            
            query_vatterms = """SELECT [VatTermID] as 'ДДС Срок ID', [Description] as 'Описание',
                [TypeIdentifier] as 'Тип', [VatValue] as 'Стойност'
                FROM [dbo].[VatTerms] ORDER BY [VatTermID]"""

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df_items = pd.read_sql(query_items, conn)
                df_vatrates = pd.read_sql(query_vatrates, conn)
                df_itemgroups = pd.read_sql(query_itemgroups, conn)
                df_status = pd.read_sql(query_status, conn)
                df_vatterms = pd.read_sql(query_vatterms, conn)
            
            if df_items.empty:
                self.log("ℹ Няма видими записи в 'Items'. Ще бъде създаден празен sheet 'Items'.")
            
            # Обработка и запис в Excel (както в предишната версия)
            df_items['Код'] = df_items['Код'].astype(str).replace(['nan', 'None', 'null'], '')
            df_items['Стока'] = df_items['Стока'].astype(str)
            
            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_items.to_excel(writer, index=False, sheet_name='Items')
                ws_items = writer.sheets['Items']
                self.auto_adjust_column_width(ws_items)
                self.format_header_bold(ws_items)
                items_count = len(df_items)
                
                # Валидации и други шийтове...
                if not df_vatrates.empty:
                    df_vatrates['Display'] = df_vatrates['ДДС ID'].astype(str) + ' - ' + df_vatrates['Описание']
                    df_vatrates[['ДДС ID', 'Display', 'Описание', 'Стойност', 'Тип']].to_excel(writer, index=False, sheet_name='VatRates')
                    if items_count > 0:
                        self.add_dropdown_validation(ws_items, 'E', 'VatRates', 'B', 2, items_count + 1)
                
                if not df_itemgroups.empty:
                    df_itemgroups['Display'] = df_itemgroups['Група ID'].astype(str) + ' - ' + df_itemgroups['Име']
                    df_itemgroups[['Група ID', 'Display', 'Име']].to_excel(writer, index=False, sheet_name='ItemGroups')
                    if items_count > 0:
                        self.add_dropdown_validation(ws_items, 'F', 'ItemGroups', 'B', 2, items_count + 1)
                
                if not df_status.empty:
                    df_status['Display'] = df_status['Статус ID'].astype(str) + ' - ' + df_status['Име']
                    df_status[['Статус ID', 'Display', 'Име']].to_excel(writer, index=False, sheet_name='Status')
                    if items_count > 0:
                        self.add_dropdown_validation(ws_items, 'G', 'Status', 'B', 2, items_count + 1)
                
                if not df_vatterms.empty:
                    df_vatterms['Display'] = df_vatterms['ДДС Срок ID'].astype(str) + ' - ' + df_vatterms['Описание']
                    df_vatterms[['ДДС Срок ID', 'Display', 'Описание', 'Тип']].to_excel(writer, index=False, sheet_name='VatTerms')
                    if items_count > 0:
                        self.add_dropdown_validation(ws_items, 'H', 'VatTerms', 'B', 2, items_count + 1)
            
            self.log(f"✓ Експортирани {len(df_items)} записа")
            if self._with_tk_dialog(lambda r: messagebox.askyesno("Успех", 
                f"Експортирани са {len(df_items)} записа.\nДа се отвори ли файла?", parent=r)):
                os.startfile(export_file)
                
        except Exception as e:
            self.log(f"✗ Грешка при експорт: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if conn:
                conn.close()

    def export_partners_to_excel(self):
        if not self.ensure_database_selected():
            self.log("Експортът е отменен: няма избрана база данни.")
            return

        initial_dir = os.path.dirname(CONFIG['excel_file']) if CONFIG['excel_file'] and os.path.exists(CONFIG['excel_file']) else os.getcwd()
        initial_name = "invoice_pro_partners_export.xlsx"
        export_file = self._with_tk_dialog(lambda r: filedialog.asksaveasfilename(
            title="Запази Excel файл като",
            initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Excel файлове", "*.xlsx"), ("Всички файлове", "*.*")],
            parent=r
        ))
        if not export_file:
            self.log("Експортът е отменен от потребителя.")
            return

        self.log(f"=== ЕКСПОРТ НА PARTNERS ОТ SQL КЪМ EXCEL ===")
        self.log(f"Сървър: {CONFIG['server']}")
        self.log(f"База: {CONFIG['database']}")
        self.log("Таблица: Partners")

        if os.path.exists(export_file):
            try:
                os.remove(export_file)
            except Exception:
                self._with_tk_dialog(lambda r: messagebox.showerror("Грешка",
                    f"Файлът е отворен в друга програма.\nМоля затворете го.", parent=r))
                return

        conn = self.connect_with_fallback()
        if not conn:
            return

        try:
            if not self.check_table_exists(conn, 'Partners'):
                self.log("✗ Таблица 'Partners' не е намерена в избраната база.")
                self._with_tk_dialog(lambda r: messagebox.showerror(
                    "Грешка",
                    "Таблица 'Partners' не е намерена в избраната база.",
                    parent=r
                ))
                return

            query_partners = """
            SELECT
                [PartnerID] as 'PartnerID',
                [Name] as 'Име',
                [NameEnglish] as 'Име (EN)',
                [ContactName] as 'Лице за контакт',
                [ContactNameEnglish] as 'Лице за контакт (EN)',
                [EMail] as 'EMail',
                [Bulstat] as 'Булстат',
                [VatId] as 'ДДС Номер',
                [BankName] as 'Банка',
                [BankCode] as 'Банков код',
                [BankAccount] as 'Банкова сметка',
                [Priority] as 'Priority',
                [GroupID] as 'GroupID',
                [Visible] as 'Visible',
                [MainPartnerID] as 'MainPartnerID',
                [StatusID] as 'StatusID',
                [IsExported] as 'IsExported',
                [IsOSSPartner] as 'IsOSSPartner',
                [CountryID] as 'CountryID',
                [DocumentEndDatePeriod] as 'DocumentEndDatePeriod'
            FROM [dbo].[Partners]
            WHERE [Visible] = 1
            ORDER BY [Name]
            """

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df_partners = pd.read_sql(query_partners, conn)

            if df_partners.empty:
                self.log("ℹ Няма видими записи в 'Partners'. Ще бъде създаден празен sheet 'Партньори'.")

            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_partners.to_excel(writer, index=False, sheet_name='Партньори')
                ws_partners = writer.sheets['Партньори']
                self.auto_adjust_column_width(ws_partners)
                self.format_header_bold(ws_partners)

            self.log(f"✓ Експортирани {len(df_partners)} партньора")
            if self._with_tk_dialog(lambda r: messagebox.askyesno("Успех",
                f"Експортирани са {len(df_partners)} партньора.\nДа се отвори ли файла?", parent=r)):
                os.startfile(export_file)

        except Exception as e:
            self.log(f"✗ Грешка при експорт на Partners: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if conn:
                conn.close()

    def get_access_odbc_driver(self):
        """Връща наличен ODBC драйвер за Microsoft Access или None."""
        drivers = pyodbc.drivers()
        access_drivers = [
            "Microsoft Access Driver (*.mdb, *.accdb)",
            "Microsoft Access Driver (*.mdb)"
        ]
        for driver in access_drivers:
            if driver in drivers:
                return driver
        return None

    def export_warehouse_pro_partners_to_excel(self):
        default_mdb_file = r"C:\ProgramData\Microinvest\Warehouse Pro\Microinvest.mdb"
        mdb_file = input(
            f"Въведете път до .MDB файл на Warehouse Pro [{default_mdb_file}]: "
        ).strip().strip('"')
        if not mdb_file:
            mdb_file = default_mdb_file

        if not os.path.exists(mdb_file):
            self.log(f"✗ .MDB файлът не е намерен: {mdb_file}")
            return

        access_driver = self.get_access_odbc_driver()
        if not access_driver:
            self.log("✗ Не е намерен ODBC драйвер за Microsoft Access.")
            self.log("  Инсталирайте Microsoft Access Database Engine (x64).")
            return

        initial_dir = os.path.dirname(mdb_file) if os.path.exists(mdb_file) else os.getcwd()
        initial_name = "warehouse_pro_partners_export.xlsx"
        export_file = self._with_tk_dialog(lambda r: filedialog.asksaveasfilename(
            title="Запази Excel файл като",
            initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension=".xlsx",
            filetypes=[("Excel файлове", "*.xlsx"), ("Всички файлове", "*.*")],
            parent=r
        ))
        if not export_file:
            self.log("Експортът е отменен от потребителя.")
            return

        if os.path.exists(export_file):
            try:
                os.remove(export_file)
            except Exception:
                self._with_tk_dialog(lambda r: messagebox.showerror(
                    "Грешка",
                    "Файлът е отворен в друга програма.\nМоля затворете го.",
                    parent=r
                ))
                return

        password = "Microinvest6380"
        conn = None

        self.log("=== ЕКСПОРТ WAREHOUSE PRO PARTNERS -> EXCEL ===")
        self.log(f"MDB файл: {mdb_file}")
        self.log("Таблица: Partners")

        try:
            conn_str = (
                f"DRIVER={{{access_driver}}};"
                f"DBQ={mdb_file};"
                f"PWD={password};"
            )
            conn = pyodbc.connect(conn_str, timeout=CONFIG['login_timeout'])

            query_partners = "SELECT * FROM [Partners]"
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df_partners = pd.read_sql(query_partners, conn)

            if df_partners.empty:
                self.log("ℹ Таблица 'Partners' е празна.")

            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_partners.to_excel(writer, index=False, sheet_name='Partners')
                ws_partners = writer.sheets['Partners']
                self.auto_adjust_column_width(ws_partners)
                self.format_header_bold(ws_partners)

            self.log(f"✓ Експортирани {len(df_partners)} партньора")
            if self._with_tk_dialog(lambda r: messagebox.askyesno(
                "Успех",
                f"Експортирани са {len(df_partners)} партньора.\nДа се отвори ли файла?",
                parent=r
            )):
                os.startfile(export_file)

        except Exception as e:
            self.log(f"✗ Грешка при експорт от Warehouse Pro: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if conn:
                conn.close()

    def export_to_excel(self):
        """Backwards-compatible alias към експорт на Items."""
        self.export_items_to_excel()
    
    def prepare_import_data(self, df):
        self.log("Подготовка на данните...")
        df = df.dropna(subset=['Код', 'Стока'], how='all')
        df['Цена'] = df['Цена'].fillna(0)
        
        data = []
        skipped = 0
        
        for idx, row in df.iterrows():
            try:
                code = str(row['Код']).strip()
                name = str(row['Стока']).strip()
                
                if not code or code == 'nan' or not name or name == 'nan':
                    skipped += 1
                    continue
                
                measure = str(row['Мярка']).strip() if pd.notna(row['Мярка']) else 'бр.'
                price = float(row['Цена']) if pd.notna(row['Цена']) else 0.0
                
                vatrate_id = self.parse_id_value(row['ДДС ID']) if 'ДДС ID' in row else None
                group_id = self.parse_id_value(row['Група ID']) if 'Група ID' in row else None
                status_id = self.parse_id_value(row['Статус ID']) if 'Статус ID' in row else None
                vatterm_id = self.parse_id_value(row['ДДС Срок ID']) if 'ДДС Срок ID' in row else None
                
                # Default стойности
                if vatrate_id is None: vatrate_id = 1
                if group_id is None: group_id = 1
                if status_id is None: status_id = 3
                if vatterm_id is None: vatterm_id = 7
                
                data.append({
                    'Code': code, 'Name': name, 'Name2': self.transliterate(name),
                    'Measure': measure, 'Measure2': self.transliterate(measure),
                    'SalePrice': price, 'GroupID': group_id, 'VatRateID': vatrate_id,
                    'StatusID': status_id, 'VatTermID': vatterm_id, 'Visible': 1,
                    'FixedPrice': 0, 'EcoTax': 0, 'Priority': 0, 'IsService': 0,
                    'MainItemID': 0, 'Barcode': '', 'Permit': ''
                })
            except Exception as e:
                self.log(f"[ПРЕДУПРЕЖДЕНИЕ] Ред {idx + 1} пропуснат: {e}")
                skipped += 1
        
        if skipped > 0:
            self.log(f"Пропуснати {skipped} невалидни реда")
        return data
    
    def import_from_excel(self):
        if not self.ensure_database_selected():
            self.log("Импортът е отменен: няма избрана база данни.")
            return
        
        import_file = self._with_tk_dialog(lambda r: filedialog.askopenfilename(
            title="Изберете Excel файл за импорт",
            filetypes=[("Excel файлове", "*.xlsx *.xls"), ("Всички файлове", "*.*")],
            initialdir=os.getcwd(),
            parent=r
        ))
        if not import_file:
            self.log("Импортът е отменен от потребителя.")
            return

        self.log(f"✓ Избран файл за импорт: {import_file}")
        self.log(f"=== ИМПОРТ ОТ EXCEL КЪМ SQL ===")

        if not os.path.exists(import_file):
            self.log("✗ Файлът не съществува!")
            return
        
        try:
            try:
                df = pd.read_excel(import_file, sheet_name='Items', skiprows=CONFIG['skiprows'])
            except ValueError:
                df = pd.read_excel(import_file, sheet_name=CONFIG['sheet_name'], skiprows=CONFIG['skiprows'])
            
            if not all(col in df.columns for col in EXPECTED_COLUMNS):
                self.log("✗ Липсват задължителни колони!")
                return
            
            if df.empty:
                self.log("✗ Файлът е празен!")
                return
            
            print("\nПърви 3 реда:")
            print(df.head(3).to_string())
            
            if not self._with_tk_dialog(lambda r: messagebox.askyesno("Потвърждение", 
                f"Ще бъдат заменени записите в '{CONFIG['table_name']}' с {len(df)} нови.\nПотвърждавате ли?", parent=r)):
                return
            
            data = self.prepare_import_data(df)
            if not data:
                return
            
            # Свързване с fallback
            conn = self.connect_with_fallback()
            if not conn:
                return
            
            cursor = conn.cursor()
            
            try:
                # Изтриване на стари записи
                cursor.execute(f"""
                    DECLARE @Targets TABLE (ItemID INT);
                    INSERT INTO @Targets SELECT ItemID FROM [dbo].[{CONFIG['table_name']}] WHERE [Visible] = 1;
                    UPDATE [dbo].[{CONFIG['table_name']}] SET [Visible] = 0 WHERE ItemID IN (SELECT ItemID FROM @Targets);
                    DELETE FROM [dbo].[{CONFIG['table_name']}] WHERE ItemID IN (SELECT ItemID FROM @Targets)
                    AND ItemID NOT IN (SELECT ItemID FROM DocumentDetails WHERE ItemID IS NOT NULL)
                    AND ItemID NOT IN (SELECT ItemID FROM DocumentTemplateDetails WHERE ItemID IS NOT NULL);
                """)
                
                # Вмъкване
                for i, item in enumerate(data):
                    cursor.execute(f"""
                        INSERT INTO [dbo].[{CONFIG['table_name']}] (
                            Code, Name, Name2, Measure, Measure2, SalePrice, GroupID, VatRateID, 
                            StatusID, VatTermID, Visible, FixedPrice, EcoTax, Priority, IsService, 
                            MainItemID, Barcode, Permit
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, tuple(item.values()))
                    
                    if (i + 1) % 100 == 0:
                        self.log(f"  ... {i + 1}/{len(data)}")
                
                conn.commit()
                self.log(f"✓ Импортирани {len(data)} записа")
                self._with_tk_dialog(lambda r: messagebox.showinfo("Успех", 
                    f"Импортирани {len(data)} записа!", parent=r))
                
            except Exception as e:
                conn.rollback()
                self.log(f"✗ Грешка: {e}")
                raise
            finally:
                conn.close()
                
        except Exception as e:
            self.log(f"✗ Грешка при импорт: {e}")
    
    def show_menu(self):
        print("\n" + "="*60)
        print("       EXCEL ↔ SQL SERVER МЕНИДЖЪР")
        print("="*60)
        print(f"Сървър: {CONFIG['server']} | База: {CONFIG['database']}")
        print(f"Таблица: {CONFIG['table_name']}")
        print("-"*60)
        print("1. 📤 Експорт Invoice Pro Стоки + свързани таблици → Excel")
        print("2. 📤 Експорт Invoice Pro Партньори → Excel")
        print("3. 📤 Експорт Warehouse Pro партньори -> Excel")
        print("4. 📥 Импорт Excel → SQL")
        print("5. 🗃️ Смяна на база данни")
        print("6. 🚪 Изход")
        print("="*60)
    
    def run(self):
        self.log("Стартиране на Excel-SQL Manager...")
        
        # Проверка на ODBC драйвер
        if not self.check_odbc_driver():
            sys.exit(1)
        
        # Проверка на връзката при стартиране (silent)
        try:
            test_conn = pyodbc.connect(self.get_connection_string(), timeout=3)
            test_conn.close()
            self.log(f"✓ Успешна връзка с {CONFIG['database']}")
        except:
            self.log(f"⚠ Неуспешна първоначална връзка с {CONFIG['database']}")
            self.log("  Ще бъде предложен избор на база при първа операция")
        
        while True:
            self.show_menu()
            choice = input("Изберете (1-6): ").strip()
            
            if choice == '1':
                self.export_items_to_excel()
            elif choice == '2':
                self.export_partners_to_excel()
            elif choice == '3':
                self.export_warehouse_pro_partners_to_excel()
            elif choice == '4':
                self.import_from_excel()
            elif choice == '5':
                self.prompt_database_selection()
            elif choice == '6':
                self.log("Изход...")
                break
            else:
                print("Невалидна опция!")

def main():
    app = ExcelSQLManager()
    app.run()

if __name__ == "__main__":
    main()
