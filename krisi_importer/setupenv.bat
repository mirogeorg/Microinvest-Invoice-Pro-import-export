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

# ==================== КОНФИГУРАЦИЯ ====================
CONFIG = {
    'server': 'DESKTOP-90UGKRP',
    'database': 'InvoicePro_26020309341273',
    'table_name': 'Items',
    'excel_file': None,  # Ще бъде зададен чрез диалог
    'sheet_name': 0,
    'skiprows': 0,
    'trusted_connection': True,
    'login_timeout': 15
}

EXPECTED_COLUMNS = ['Код', 'Стока', 'Мярка', 'Цена']

class ExcelSQLManager:
    def __init__(self):
        self.selected_file = None
        self.root = tk.Tk()
        self.root.withdraw()  # Скриваме главния прозорец
        
    def log(self, message):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{timestamp}] {message}")
    
    def bring_console_to_front(self):
        """Връща фокуса към конзолния прозорец след Windows диалог"""
        try:
            # Взимаме хендъла на текущата конзола
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                # Връщаме фокуса
                ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass  # Ако не успеем, продължаваме без грешка
    
    def transliterate(self, text):
        """Транслитерация от български (кирилица) към латиница."""
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
        
        result = []
        text = str(text)
        for char in text:
            result.append(trans_map.get(char, char))
        return ''.join(result)
    
    def get_connection_string(self):
        driver = "ODBC Driver 17 for SQL Server"
        return (f"DRIVER={{{driver}}};"
                f"SERVER={CONFIG['server']};"
                f"DATABASE={CONFIG['database']};"
                f"Trusted_Connection=yes;"
                f"Login Timeout={CONFIG['login_timeout']};")
    
    def auto_adjust_column_width(self, worksheet):
        """Автоматично настройва ширината на колоните в шийта"""
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
        """Прави header-а на шийта удебелен"""
        for cell in worksheet[1]:
            new_font = copy(cell.font)
            new_font.bold = True
            cell.font = new_font
    
    def select_file_dialog(self):
        """Windows диалог за избор на Excel файл"""
        self.log("Отваряне на диалог за избор на файл...")
        
        file_path = filedialog.askopenfilename(
            title="Изберете Excel файл",
            filetypes=[
                ("Excel файлове", "*.xlsx *.xls"),
                ("Всички файлове", "*.*")
            ],
            initialdir=os.getcwd()
        )
        
        if file_path:
            self.selected_file = file_path
            self.log(f"✓ Избран файл: {file_path}")
            return True
        else:
            self.log("✗ Не е избран файл")
            return False
    
    def check_file_selected(self):
        """Проверява дали е избран файл"""
        if not self.selected_file:
            print("\n!!! Моля първо изберете файл (опция 1) !!!")
            return False
        return True
    
    def export_to_excel(self):
        """Експорт от SQL таблица към Excel файл с допълнителни шийтове за VatRates и ItemGroups"""
        if not self.check_file_selected():
            return
        
        # Формираме име за експортирания файл
        base, ext = os.path.splitext(self.selected_file)
        export_file = f"{base}_exported{ext}"
        
        self.log(f"=== ЕКСПОРТ ОТ SQL КЪМ EXCEL ===")
        self.log(f"Източник: {CONFIG['database']}.{CONFIG['table_name']}")
        self.log(f"Дестинация: {export_file}")
        
        # Проверка дали файлът съществува и може ли да бъде изтрит
        if os.path.exists(export_file):
            try:
                os.remove(export_file)
                self.log(f"Съществуващ файл изтрит: {export_file}")
            except Exception as e:
                self.log(f"✗ Не може да се изтрие съществуващия файл (вероятно е отворен в Excel): {e}")
                messagebox.showerror("Грешка", 
                    f"Файлът '{os.path.basename(export_file)}' е отворен в друга програма.\n"
                    f"Моля затворете го и опитайте отново.")
                return
        
        try:
            # Свързване с базата
            self.log("Свързване с SQL Server...")
            conn = pyodbc.connect(self.get_connection_string())
            
            # 1. Четем данните за Items с VatRateID и GroupID
            query_items = f"""
            SELECT [Code] as 'Код', 
                   [Name] as 'Стока', 
                   [Measure] as 'Мярка', 
                   [SalePrice] as 'Цена',
                   [VatRateID] as 'ДДС ID',
                   [GroupID] as 'Група ID'
            FROM [dbo].[{CONFIG['table_name']}]
            WHERE [Visible] = 1
            ORDER BY [Name]
            """
            
            # 2. Четем данните за VatRates
            query_vatrates = """
            SELECT [VatRateID] as 'ДДС ID',
                   [Code] as 'Код',
                   [Description] as 'Описание',
                   [Rate] as 'Стойност',
                   [TypeIdentifier] as 'Тип'
            FROM [dbo].[VatRates]
            ORDER BY [VatRateID]
            """
            
            # 3. Четем данните за ItemGroups
            query_itemgroups = """
            SELECT [GroupID] as 'Група ID',
                   [Code] as 'Код',
                   [Name] as 'Име'
            FROM [dbo].[ItemGroups]
            ORDER BY [GroupID]
            """
            
            # Потискаме warning-а за pandas и SQLAlchemy
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df_items = pd.read_sql(query_items, conn)
                df_vatrates = pd.read_sql(query_vatrates, conn)
                df_itemgroups = pd.read_sql(query_itemgroups, conn)
            
            conn.close()
            
            if df_items.empty:
                self.log("⚠ Няма данни за експортиране в Items")
                return
            
            # ВАЖНО: Конвертиране на ID колоните към текст, за да не се загубят водещи нули при Код
            df_items['Код'] = df_items['Код'].astype(str).replace(['nan', 'None', 'null'], '')
            df_items['Стока'] = df_items['Стока'].astype(str)
            # VatRateID и GroupID са числа, оставяме ги така или ги конвертираме към цели числа
            df_items['ДДС ID'] = df_items['ДДС ID'].fillna(0).astype(int)
            df_items['Група ID'] = df_items['Група ID'].fillna(0).astype(int)
            
            self.log(f"Подготвени {len(df_items)} записа за експорт от Items")
            self.log(f"Подготвени {len(df_vatrates)} записа за експорт от VatRates")
            self.log(f"Подготвени {len(df_itemgroups)} записа за експорт от ItemGroups")
            
            # Запис с openpyxl за форматиране на всички шийтове
            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                # === ШИЙТ 1: ITEMS ===
                df_items.to_excel(writer, index=False, sheet_name='Items')
                worksheet_items = writer.sheets['Items']
                
                # Авто-ширина на колоните
                self.auto_adjust_column_width(worksheet_items)
                
                # Форматиране на колоните в Items
                for row in range(2, worksheet_items.max_row + 1):
                    # Код (колона A) - текст
                    worksheet_items.cell(row=row, column=1).number_format = '@'
                    # Стока (колона B) - текст  
                    worksheet_items.cell(row=row, column=2).number_format = '@'
                    # Цена (колона D) - число с 2 знака след десетичната запетая
                    worksheet_items.cell(row=row, column=4).number_format = '0.00'
                    # ДДС ID (колона E) - цяло число
                    worksheet_items.cell(row=row, column=5).number_format = '0'
                    # Група ID (колона F) - цяло число
                    worksheet_items.cell(row=row, column=6).number_format = '0'
                
                # Форматиране на Header-а (удебелен)
                self.format_header_bold(worksheet_items)
                
                # === ШИЙТ 2: VATRATES ===
                if not df_vatrates.empty:
                    df_vatrates.to_excel(writer, index=False, sheet_name='VatRates')
                    worksheet_vat = writer.sheets['VatRates']
                    self.auto_adjust_column_width(worksheet_vat)
                    
                    # Форматиране: Код е текст, Стойност е число с 2 знака
                    for row in range(2, worksheet_vat.max_row + 1):
                        worksheet_vat.cell(row=row, column=2).number_format = '@'  # Код
                        worksheet_vat.cell(row=row, column=4).number_format = '0.00'  # Стойност
                    
                    self.format_header_bold(worksheet_vat)
                    self.log("✓ Шийт VatRates създаден успешно")
                else:
                    self.log("⚠ Няма данни в VatRates")
                
                # === ШИЙТ 3: ITEMGROUPS ===
                if not df_itemgroups.empty:
                    df_itemgroups.to_excel(writer, index=False, sheet_name='ItemGroups')
                    worksheet_groups = writer.sheets['ItemGroups']
                    self.auto_adjust_column_width(worksheet_groups)
                    
                    # Форматиране: Код и Име са текст
                    for row in range(2, worksheet_groups.max_row + 1):
                        worksheet_groups.cell(row=row, column=2).number_format = '@'  # Код
                        worksheet_groups.cell(row=row, column=3).number_format = '@'  # Име
                    
                    self.format_header_bold(worksheet_groups)
                    self.log("✓ Шийт ItemGroups създаден успешно")
                else:
                    self.log("⚠ Няма данни в ItemGroups")
            
            self.log(f"✓ Успешно експортирани {len(df_items)} записа от Items")
            self.log(f"  Формат: Код=TEXT, Стока=TEXT, Цена=0.00, ДДС ID=0, Група ID=0")
            self.log(f"  Допълнителни шийтове: VatRates ({len(df_vatrates)} реда), ItemGroups ({len(df_itemgroups)} реда)")
            self.log(f"  Файл: {os.path.abspath(export_file)}")
            
            if messagebox.askyesno("Експорт завършен", 
                                   f"Експортирани са:\n"
                                   f"• {len(df_items)} записа от Items\n"
                                   f"• {len(df_vatrates)} записа от VatRates\n"
                                   f"• {len(df_itemgroups)} записа от ItemGroups\n\n"
                                   f"Да се отвори ли файла?"):
                os.startfile(export_file)
                
        except Exception as e:
            self.log(f"✗ Грешка при експорт: {e}")
            import traceback
            traceback.print_exc()
    
    def prepare_import_data(self, df):
        """Подготовка на данните за импорт"""
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
                
                # Четем VatRateID и GroupID ако ги има в Excel, иначе default стойности
                vatrate_id = int(row['ДДС ID']) if 'ДДС ID' in row and pd.notna(row['ДДС ID']) else 1
                group_id = int(row['Група ID']) if 'Група ID' in row and pd.notna(row['Група ID']) else 1
                
                name2 = self.transliterate(name)
                measure2 = self.transliterate(measure)
                
                data.append({
                    'Code': code,
                    'Name': name,
                    'Name2': name2,
                    'Measure': measure,
                    'Measure2': measure2,
                    'SalePrice': price,
                    'GroupID': group_id,
                    'VatRateID': vatrate_id,
                    'StatusID': 3,
                    'VatTermID': 7,
                    'Visible': 1,
                    'FixedPrice': 0,
                    'EcoTax': 0,
                    'Priority': 0,
                    'IsService': 0,
                    'MainItemID': 0,
                    'Barcode': '',
                    'Permit': ''
                })
            except Exception as e:
                self.log(f"[ПРЕДУПРЕЖДЕНИЕ] Ред {idx + 1} пропуснат: {e}")
                skipped += 1
                continue
        
        if skipped > 0:
            self.log(f"Пропуснати {skipped} невалидни реда")
        
        self.log(f"Подготвени {len(data)} записа за импортиране")
        return data
    
    def execute_sql_import(self, cursor, data):
        """Изпълнение на SQL операциите за импорт"""
        self.log("Стартиране на SQL транзакция...")
        
        # 1. Скриване/изтриване на стари данни
        delete_script = f"""
        DECLARE @Targets TABLE (ItemID INT);
        INSERT INTO @Targets (ItemID)
        SELECT ItemID FROM [dbo].[{CONFIG['table_name']}] WHERE [Visible] = 1;

        UPDATE [dbo].[{CONFIG['table_name']}]
        SET [Visible] = 0
        WHERE ItemID IN (SELECT ItemID FROM @Targets);

        DELETE FROM [dbo].[{CONFIG['table_name']}]
        WHERE ItemID IN (SELECT ItemID FROM @Targets)
          AND ItemID NOT IN (SELECT DISTINCT ItemID FROM [dbo].[DocumentDetails] WHERE ItemID IS NOT NULL)
          AND ItemID NOT IN (SELECT DISTINCT ItemID FROM [dbo].[DocumentTemplateDetails] WHERE ItemID IS NOT NULL);
        """
        cursor.execute(delete_script)
        self.log("✓ Старите записи са скрити/изтрити")
        
        # 2. Вмъкване на нови данни
        self.log(f"Вмъкване на {len(data)} записа...")
        new_ids = []
        
        for i, item in enumerate(data):
            insert_sql = f"""
            INSERT INTO [dbo].[{CONFIG['table_name']}] (
                [Code], [Name], [Name2], [Measure], [Measure2], [SalePrice],
                [GroupID], [VatRateID], [StatusID], [VatTermID], [Visible],
                [FixedPrice], [EcoTax], [Priority], [IsService], [MainItemID],
                [Barcode], [Permit]
            )
            OUTPUT inserted.ItemID
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """
            
            cursor.execute(insert_sql, (
                item['Code'], item['Name'], item['Name2'], item['Measure'], 
                item['Measure2'], item['SalePrice'], item['GroupID'], 
                item['VatRateID'], item['StatusID'], item['VatTermID'], 
                item['Visible'], item['FixedPrice'], item['EcoTax'], 
                item['Priority'], item['IsService'], item['MainItemID'],
                item['Barcode'], item['Permit']
            ))
            
            new_id = cursor.fetchone()[0]
            new_ids.append(new_id)
            
            if (i + 1) % 100 == 0:
                self.log(f"  ... {i + 1}/{len(data)} записа")
        
        self.log(f"✓ Вмъкнати {len(new_ids)} записа")
        
        # 3. Обновяване на MainItemID
        if new_ids:
            ids_string = ','.join(str(id) for id in new_ids)
            update_sql = f"""
            UPDATE [dbo].[{CONFIG['table_name']}]
            SET [MainItemID] = [ItemID]
            WHERE [ItemID] IN ({ids_string});
            """
            cursor.execute(update_sql)
            self.log(f"✓ Обновени MainItemID за {cursor.rowcount} записа")
    
    def import_from_excel(self):
        """Импорт от избран Excel файл към SQL"""
        if not self.check_file_selected():
            return
        
        self.log(f"=== ИМПОРТ ОТ EXCEL КЪМ SQL SERVER ===")
        self.log(f"Файл: {self.selected_file}")
        
        # Проверка на структурата
        if not os.path.exists(self.selected_file):
            self.log(f"✗ Файлът не съществува: {self.selected_file}")
            return
        
        try:
            # Четем само шийт Items за импорт (ако съществува)
            try:
                df = pd.read_excel(self.selected_file, sheet_name='Items', skiprows=CONFIG['skiprows'])
                self.log(f"Прочетени {len(df)} реда от шийт 'Items'")
            except ValueError:
                # Ако няма шийт Items, четем първия шийт (за съвместимост със стари файлове)
                df = pd.read_excel(self.selected_file, sheet_name=CONFIG['sheet_name'], skiprows=CONFIG['skiprows'])
                self.log(f"Прочетени {len(df)} реда от първия шийт")
            
            actual_columns = list(df.columns)
            missing_columns = [col for col in EXPECTED_COLUMNS if col not in actual_columns]
            
            if missing_columns:
                self.log(f"✗ Липсват колони: {missing_columns}")
                self.log(f"  Очаквани: {EXPECTED_COLUMNS}")
                self.log(f"  Намерени: {actual_columns}")
                return
            
            if df.empty:
                self.log("✗ Excel файлът е празен!")
                return
            
            # Показваме примерни данни
            print("\nПърви 3 реда от файла:")
            print(df.head(3).to_string())
            print("-" * 60)
            
            # Потвърждение
            if not messagebox.askyesno("Потвърждение", 
                f"Ще бъдат изтрити всички видими записи в таблицата '{CONFIG['table_name']}'\n"
                f"и ще бъдат вмъкнати {len(df)} нови записа от избрания файл.\n\n"
                f"Потвърждавате ли?"):
                self.log("Импортът е отменен от потребителя")
                return
            
            # Подготовка на данните
            data = self.prepare_import_data(df)
            if not data:
                self.log("Няма валидни данни за импортиране!")
                return
            
            # Примери за транслитерация
            self.log("Примери за транслитерация:")
            for i, item in enumerate(data[:3]):
                self.log(f"  {i+1}. '{item['Name']}' -> '{item['Name2']}'")
            
            # SQL операции
            self.log("Свързване с SQL Server...")
            conn = pyodbc.connect(self.get_connection_string())
            cursor = conn.cursor()
            self.log("✓ Свързването е успешно")
            
            cursor.execute("BEGIN TRANSACTION;")
            
            try:
                self.execute_sql_import(cursor, data)
                cursor.execute("COMMIT TRANSACTION;")
                conn.commit()
                self.log("✓ Транзакцията е потвърдена")
                messagebox.showinfo("Успех", f"Успешно импортирани {len(data)} записа!")
                
            except Exception as e:
                self.log(f"✗ Грешка в SQL: {e}")
                cursor.execute("ROLLBACK TRANSACTION;")
                conn.rollback()
                self.log("✓ Транзакцията е отменена (ROLLBACK)")
                raise
            finally:
                conn.close()
                
            self.log("=== ИМПОРТЪТ ЗАВЪРШИ УСПЕШНО ===")
            
        except Exception as e:
            self.log(f"✗ Грешка при импорт: {e}")
            import traceback
            traceback.print_exc()
    
    def show_menu(self):
        """Показва главното меню"""
        print("\n" + "="*60)
        print("       EXCEL ↔ SQL SERVER МЕНИДЖЪР")
        print("="*60)
        
        if self.selected_file:
            print(f"Текущ файл: {os.path.basename(self.selected_file)}")
        else:
            print("Текущ файл: [не е избран]")
            
        print("-"*60)
        print("1. 📂 Избор на файл (Windows диалог)")
        print("2. 📤 Експорт от SQL към Excel (_exported)")
        print("3. 📥 Импорт от Excel към SQL (ИЗТРИВА стари данни!)")
        print("4. 🚪 Изход")
        print("="*60)
    
    def run(self):
        """Главен цикъл на програмата"""
        self.log("Стартиране на Excel-SQL Manager...")
        
        while True:
            self.show_menu()
            choice = input("Изберете опция (1-4): ").strip()
            
            if choice == '1':
                self.select_file_dialog()
                self.bring_console_to_front()
            elif choice == '2':
                self.export_to_excel()
                self.bring_console_to_front()
            elif choice == '3':
                self.import_from_excel()
                self.bring_console_to_front()
            elif choice == '4':
                self.log("Изход от програмата...")
                break
            else:
                print("Невалидна опция! Моля изберете 1-4.")

def main():
    app = ExcelSQLManager()
    app.run()

if __name__ == "__main__":
    main()