import os
import warnings

import pandas as pd
import pyodbc
from tkinter import filedialog, messagebox

try:
    from ..config import CONFIG
except ImportError:
    from config import CONFIG


class ExportMixin:
    def export_items_to_excel(self):
        if not self.ensure_database_selected():
            self.log('Експортът е отменен: няма избрана база данни.')
            return

        initial_dir = os.path.dirname(CONFIG['excel_file']) if CONFIG['excel_file'] and os.path.exists(CONFIG['excel_file']) else os.getcwd()
        initial_name = 'invoice_pro_items_export.xlsx'
        export_file = self._with_tk_dialog(
            lambda r: filedialog.asksaveasfilename(
                title='Запази Excel файл като',
                initialdir=initial_dir,
                initialfile=initial_name,
                defaultextension='.xlsx',
                filetypes=[('Excel файлове', '*.xlsx'), ('Всички файлове', '*.*')],
                parent=r,
            )
        )
        if not export_file:
            self.log('Експортът е отменен от потребителя.')
            return

        self.log('=== ЕКСПОРТ ОТ SQL КЪМ EXCEL ===')
        self.log(f"Сървър: {CONFIG['server']}")
        self.log(f"База: {CONFIG['database']}")
        self.log(f"Таблица: {CONFIG['table_name']}")

        if os.path.exists(export_file):
            try:
                os.remove(export_file)
            except Exception:
                self._with_tk_dialog(
                    lambda r: messagebox.showerror('Грешка', 'Файлът е отворен в друга програма.\nМоля затворете го.', parent=r)
                )
                return

        conn = self.connect_with_fallback()
        if not conn:
            return

        try:
            cursor = conn.cursor()
            try:
                cursor.execute(f"SELECT TOP 1 * FROM [dbo].[{CONFIG['table_name']}]")
                cursor.fetchone()
            except pyodbc.Error as e:
                self.log(f'✗ Грешка при достъп до таблица: {e}')
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
                warnings.simplefilter('ignore')
                df_items = pd.read_sql(query_items, conn)
                df_vatrates = pd.read_sql(query_vatrates, conn)
                df_itemgroups = pd.read_sql(query_itemgroups, conn)
                df_status = pd.read_sql(query_status, conn)
                df_vatterms = pd.read_sql(query_vatterms, conn)

            if df_items.empty:
                self.log("ℹ Няма видими записи в 'Items'. Ще бъде създаден празен sheet 'Items'.")

            df_items['Код'] = df_items['Код'].astype(str).replace(['nan', 'None', 'null'], '')
            df_items['Стока'] = df_items['Стока'].astype(str)

            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_items.to_excel(writer, index=False, sheet_name='Items')
                ws_items = writer.sheets['Items']
                for row in range(2, len(df_items) + 2):
                    ws_items[f'A{row}'].number_format = '@'
                    ws_items[f'C{row}'].number_format = '@'
                    ws_items[f'D{row}'].number_format = '0.00'
                self.auto_adjust_column_width(ws_items)
                self.format_header_bold(ws_items)
                items_count = len(df_items)

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
            if self._with_tk_dialog(
                lambda r: messagebox.askyesno('Успех', f"Експортирани са {len(df_items)} записа.\nДа се отвори ли файла?", parent=r)
            ):
                os.startfile(export_file)

        except Exception as e:
            self.log(f'✗ Грешка при експорт: {e}')
            import traceback

            traceback.print_exc()
        finally:
            if conn:
                conn.close()

    def export_partners_to_excel(self):
        if not self.ensure_database_selected():
            self.log('Експортът е отменен: няма избрана база данни.')
            return

        initial_dir = os.path.dirname(CONFIG['excel_file']) if CONFIG['excel_file'] and os.path.exists(CONFIG['excel_file']) else os.getcwd()
        initial_name = 'invoice_pro_partners_export.xlsx'
        export_file = self._with_tk_dialog(
            lambda r: filedialog.asksaveasfilename(
                title='Запази Excel файл като',
                initialdir=initial_dir,
                initialfile=initial_name,
                defaultextension='.xlsx',
                filetypes=[('Excel файлове', '*.xlsx'), ('Всички файлове', '*.*')],
                parent=r,
            )
        )
        if not export_file:
            self.log('Експортът е отменен от потребителя.')
            return

        self.log('=== ЕКСПОРТ НА PARTNERS ОТ SQL КЪМ EXCEL ===')
        self.log(f"Сървър: {CONFIG['server']}")
        self.log(f"База: {CONFIG['database']}")
        self.log('Таблица: Partners')

        if os.path.exists(export_file):
            try:
                os.remove(export_file)
            except Exception:
                self._with_tk_dialog(
                    lambda r: messagebox.showerror('Грешка', 'Файлът е отворен в друга програма.\nМоля затворете го.', parent=r)
                )
                return

        conn = self.connect_with_fallback()
        if not conn:
            return

        try:
            if not self.check_table_exists(conn, 'Partners'):
                self.log("✗ Таблица 'Partners' не е намерена в избраната база.")
                self._with_tk_dialog(
                    lambda r: messagebox.showerror('Грешка', "Таблица 'Partners' не е намерена в избраната база.", parent=r)
                )
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
                warnings.simplefilter('ignore')
                df_partners = pd.read_sql(query_partners, conn)

            if df_partners.empty:
                self.log("ℹ Няма видими записи в 'Partners'. Ще бъде създаден празен sheet 'Партньори'.")

            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_partners.to_excel(writer, index=False, sheet_name='Партньори')
                ws_partners = writer.sheets['Партньори']
                self.auto_adjust_column_width(ws_partners)
                self.format_header_bold(ws_partners)

            self.log(f"✓ Експортирани {len(df_partners)} партньора")
            if self._with_tk_dialog(
                lambda r: messagebox.askyesno('Успех', f"Експортирани са {len(df_partners)} партньора.\nДа се отвори ли файла?", parent=r)
            ):
                os.startfile(export_file)

        except Exception as e:
            self.log(f'✗ Грешка при експорт на Partners: {e}')
            import traceback

            traceback.print_exc()
        finally:
            if conn:
                conn.close()

    def export_warehouse_pro_partners_to_excel(self):
        default_mdb_file = r'C:\ProgramData\Microinvest\Warehouse Pro\Microinvest.mdb'
        mdb_file = input(f"Въведете път до .MDB файл на Warehouse Pro [{default_mdb_file}]: ").strip().strip('"')
        if not mdb_file:
            mdb_file = default_mdb_file

        if not os.path.exists(mdb_file):
            self.log(f'✗ .MDB файлът не е намерен: {mdb_file}')
            return

        access_driver = self.get_access_odbc_driver()
        if not access_driver:
            self.log('✗ Не е намерен ODBC драйвер за Microsoft Access.')
            self.log('  Инсталирайте Microsoft Access Database Engine (x64).')
            return

        initial_dir = os.path.dirname(mdb_file) if os.path.exists(mdb_file) else os.getcwd()
        initial_name = 'warehouse_pro_partners_export.xlsx'
        export_file = self._with_tk_dialog(
            lambda r: filedialog.asksaveasfilename(
                title='Запази Excel файл като',
                initialdir=initial_dir,
                initialfile=initial_name,
                defaultextension='.xlsx',
                filetypes=[('Excel файлове', '*.xlsx'), ('Всички файлове', '*.*')],
                parent=r,
            )
        )
        if not export_file:
            self.log('Експортът е отменен от потребителя.')
            return

        if os.path.exists(export_file):
            try:
                os.remove(export_file)
            except Exception:
                self._with_tk_dialog(
                    lambda r: messagebox.showerror('Грешка', 'Файлът е отворен в друга програма.\nМоля затворете го.', parent=r)
                )
                return

        password = 'Microinvest6380'
        conn = None

        self.log('=== ЕКСПОРТ WAREHOUSE PRO PARTNERS -> EXCEL ===')
        self.log(f'MDB файл: {mdb_file}')
        self.log('Таблица: Partners')

        try:
            conn_str = f"DRIVER={{{access_driver}}};DBQ={mdb_file};PWD={password};"
            conn = pyodbc.connect(conn_str, timeout=CONFIG['login_timeout'])

            query_partners = 'SELECT * FROM [Partners]'
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                df_partners = pd.read_sql(query_partners, conn)

            if df_partners.empty:
                self.log("ℹ Таблица 'Partners' е празна.")

            with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
                df_partners.to_excel(writer, index=False, sheet_name='Partners')
                ws_partners = writer.sheets['Partners']
                self.auto_adjust_column_width(ws_partners)
                self.format_header_bold(ws_partners)

            self.log(f"✓ Експортирани {len(df_partners)} партньора")
            if self._with_tk_dialog(
                lambda r: messagebox.askyesno('Успех', f"Експортирани са {len(df_partners)} партньора.\nДа се отвори ли файла?", parent=r)
            ):
                os.startfile(export_file)

        except Exception as e:
            self.log(f'✗ Грешка при експорт от Warehouse Pro: {e}')
            import traceback

            traceback.print_exc()
        finally:
            if conn:
                conn.close()

    def export_to_excel(self):
        self.export_items_to_excel()
