import pandas as pd
import pyodbc
import os
import sys
import warnings
import ctypes
import json
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
    'database': os.getenv('DB_DATABASE', 'InvoicePro_26020309341273'),
    'table_name': os.getenv('DB_TABLE', 'Items'),
    'excel_file': os.getenv('EXCEL_FILE', None),
    'sheet_name': int(os.getenv('EXCEL_SHEET', '0')),
    'skiprows': int(os.getenv('EXCEL_SKIPROWS', '0')),
    'trusted_connection': os.getenv('DB_TRUSTED_CONNECTION', 'True').lower() == 'true',
    'login_timeout': int(os.getenv('DB_TIMEOUT', '15'))
}

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app_config.json')
EXPECTED_COLUMNS = ['Код', 'Стока', 'Мярка', 'Цена']

class ExcelSQLManager:
    def __init__(self):
        self.selected_file = None
        self.load_settings()
        
        if not self.selected_file and CONFIG['excel_file'] and os.path.exists(CONFIG['excel_file']):
            self.selected_file = CONFIG['excel_file']
            self.log(f"Зареден файл от .env: {self.selected_file}")
    
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
    
    def check_table_exists(self, conn):
        """Проверява дали таблицата съществува в текущата база"""
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES 
                WHERE TABLE_NAME = ? AND TABLE_TYPE = 'BASE TABLE'
            """, (CONFIG['table_name'],))
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
    
    def load_settings(self):
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    last_file = settings.get('last_selected_file')
                    if last_file and os.path.exists(last_file):
                        self.selected_file = last_file
                        self.log(f"Зареден последен файл: {last_file}")
        except Exception as e:
            self.log(f"Не може да се заредят настройките: {e}")
    
    def save_settings(self):
        try:
            settings = {
                'last_selected_file': self.selected_file,
                'last_used': datetime.now().isoformat()
            }
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log(f"Не може да се запазят настройките: {e}")
    
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
    
    def select_file_dialog(self):
        self.log("Отваряне на диалог за избор на файл...")
        initial_dir = os.path.dirname(self.selected_file) if self.selected_file else os.getcwd()
        
        file_path = self._with_tk_dialog(lambda r: filedialog.askopenfilename(
            title="Изберете Excel файл",
            filetypes=[("Excel файлове", "*.xlsx *.xls"), ("Всички файлове", "*.*")],
            initialdir=initial_dir,
            parent=r
        ))
        
        if file_path:
            self.selected_file = file_path
            self.save_settings()
            self.log(f"✓ Избран файл: {file_path}")
            return True
        else:
            self.log("✗ Не е избран файл")
            return False
    
    def check_file_selected(self):
        if not self.selected_file:
            print("\n!!! Моля първо изберете файл (опция 1) !!!")
            return False
        return True
    
    def connect_with_fallback(self):
        """Опитва се да се свърже, при неуспех предлага избор на база"""
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
    
    def export_to_excel(self):
        if not self.check_file_selected():
            return
        
        base, ext = os.path.splitext(self.selected_file)
        export_file = f"{base}_exported{ext}"
        
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
                self.log("⚠ Няма данни за експортиране")
                self._with_tk_dialog(lambda r: messagebox.showwarning("Внимание", 
                    "Няма видими записи в таблицата!", parent=r))
                return
            
            # Обработка и запис в Excel (както в предишната версия)
            df_items['Код'] = df_items['Код'].astype(str).replace(['nan', 'None', 'null'], '')
            df_items['Стока'] = df_items['Стока'].astype(str)
            
            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_items.to_excel(writer, index=False, sheet_name='Items')
                ws_items = writer.sheets['Items']
                self.auto_adjust_column_width(ws_items)
                self.format_header_bold(ws_items)
                
                # Валидации и други шийтове...
                if not df_vatrates.empty:
                    df_vatrates['Display'] = df_vatrates['ДДС ID'].astype(str) + ' - ' + df_vatrates['Описание']
                    df_vatrates[['ДДС ID', 'Display', 'Описание', 'Стойност', 'Тип']].to_excel(writer, index=False, sheet_name='VatRates')
                    self.add_dropdown_validation(ws_items, 'E', 'VatRates', 'B', 2, len(df_items)+1)
                
                if not df_itemgroups.empty:
                    df_itemgroups['Display'] = df_itemgroups['Група ID'].astype(str) + ' - ' + df_itemgroups['Име']
                    df_itemgroups[['Група ID', 'Display', 'Име']].to_excel(writer, index=False, sheet_name='ItemGroups')
                    self.add_dropdown_validation(ws_items, 'F', 'ItemGroups', 'B', 2, len(df_items)+1)
                
                if not df_status.empty:
                    df_status['Display'] = df_status['Статус ID'].astype(str) + ' - ' + df_status['Име']
                    df_status[['Статус ID', 'Display', 'Име']].to_excel(writer, index=False, sheet_name='Status')
                    self.add_dropdown_validation(ws_items, 'G', 'Status', 'B', 2, len(df_items)+1)
                
                if not df_vatterms.empty:
                    df_vatterms['Display'] = df_vatterms['ДДС Срок ID'].astype(str) + ' - ' + df_vatterms['Описание']
                    df_vatterms[['ДДС Срок ID', 'Display', 'Описание', 'Тип']].to_excel(writer, index=False, sheet_name='VatTerms')
                    self.add_dropdown_validation(ws_items, 'H', 'VatTerms', 'B', 2, len(df_items)+1)
            
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
        if not self.check_file_selected():
            return
        
        self.log(f"=== ИМПОРТ ОТ EXCEL КЪМ SQL ===")
        
        if not os.path.exists(self.selected_file):
            self.log("✗ Файлът не съществува!")
            return
        
        try:
            try:
                df = pd.read_excel(self.selected_file, sheet_name='Items', skiprows=CONFIG['skiprows'])
            except ValueError:
                df = pd.read_excel(self.selected_file, sheet_name=CONFIG['sheet_name'], skiprows=CONFIG['skiprows'])
            
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
        if self.selected_file:
            path = self.selected_file if len(self.selected_file) < 50 else "..." + self.selected_file[-47:]
            print(f"Файл: {path}")
        else:
            print("Файл: [не е избран]")
        print("-"*60)
        print("1. 📂 Избор на файл")
        print("2. 📤 Експорт SQL → Excel")
        print("3. 📥 Импорт Excel → SQL")
        print("4. 🗃️  Смяна на база данни")
        print("5. 🚪 Изход")
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
            choice = input("Изберете (1-5): ").strip()
            
            if choice == '1':
                self.select_file_dialog()
            elif choice == '2':
                self.export_to_excel()
            elif choice == '3':
                self.import_from_excel()
            elif choice == '4':
                self.prompt_database_selection()
            elif choice == '5':
                self.save_settings()
                self.log("Изход...")
                break
            else:
                print("Невалидна опция!")

def main():
    app = ExcelSQLManager()
    app.run()

if __name__ == "__main__":
    main()