import os

import pandas as pd
from tkinter import filedialog, messagebox

try:
    from .config import CONFIG, EXPECTED_COLUMNS
    from .db import connect_with_fallback, ensure_database_selected
    from .utils import parse_id_value, transliterate, with_tk_dialog
except ImportError:
    from config import CONFIG, EXPECTED_COLUMNS
    from db import connect_with_fallback, ensure_database_selected
    from utils import parse_id_value, transliterate, with_tk_dialog


PARTNERS_NAME_COLUMNS = ['Име', 'Name', 'Company']


def build_items_import_payload(df, log):
    log('Подготовка на данните...')
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

            vatrate_id = parse_id_value(row['ДДС ID']) if 'ДДС ID' in row else None
            group_id = parse_id_value(row['Група ID']) if 'Група ID' in row else None
            status_id = parse_id_value(row['Статус ID']) if 'Статус ID' in row else None
            vatterm_id = parse_id_value(row['ДДС Срок ID']) if 'ДДС Срок ID' in row else None

            if vatrate_id is None:
                vatrate_id = 1
            if group_id is None:
                group_id = 1
            if status_id is None:
                status_id = 3
            if vatterm_id is None:
                vatterm_id = 7

            data.append(
                {
                    'Code': code,
                    'Name': name,
                    'Name2': transliterate(name),
                    'Measure': measure,
                    'Measure2': transliterate(measure),
                    'SalePrice': price,
                    'GroupID': group_id,
                    'VatRateID': vatrate_id,
                    'StatusID': status_id,
                    'VatTermID': vatterm_id,
                    'Visible': 1,
                    'FixedPrice': 0,
                    'EcoTax': 0,
                    'Priority': 0,
                    'IsService': 0,
                    'MainItemID': 0,
                    'Barcode': '',
                    'Permit': '',
                }
            )
        except Exception as e:
            log(f'[ПРЕДУПРЕЖДЕНИЕ] Ред {idx + 1} пропуснат: {e}')
            skipped += 1

    if skipped > 0:
        log(f'Пропуснати {skipped} невалидни реда')
    return data


def _to_int(value, default=0):
    if pd.isna(value):
        return default

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in ('true', 'yes', 'да'):
            return 1
        if normalized in ('false', 'no', 'не'):
            return 0

    parsed = parse_id_value(value)
    if parsed is not None:
        return parsed

    return default


def _to_clean_string(value, default=''):
    if pd.isna(value):
        return default
    text = str(value).strip()
    return text if text and text.lower() != 'nan' else default


def build_partners_import_payload(df, log):
    if not any(col in df.columns for col in PARTNERS_NAME_COLUMNS):
        log("✗ Липсва колона за име (очаква се 'Име', 'Name' или 'Company').")
        return []

    data = []
    skipped = 0

    for _, row in df.iterrows():
        name = _to_clean_string(_pick_first_existing_value(row, ['Име', 'Name', 'Company'], default=''))
        if not name:
            skipped += 1
            continue

        contact_name = _to_clean_string(_pick_first_existing_value(row, ['Лице за контакт', 'ContactName', 'MOL'], default=''))
        payload = {
            'Name': name,
            'NameEnglish': _to_clean_string(
                _pick_first_existing_value(row, ['Име (EN)', 'NameEnglish'], default=''),
                default=transliterate(name),
            ),
            'ContactName': contact_name,
            'ContactNameEnglish': _to_clean_string(
                _pick_first_existing_value(row, ['Лице за контакт (EN)', 'ContactNameEnglish'], default=''),
                default=transliterate(contact_name),
            ),
            'EMail': _to_clean_string(_pick_first_existing_value(row, ['EMail', 'Email'], default='')),
            'Bulstat': _to_clean_string(_pick_first_existing_value(row, ['Булстат', 'Bulstat'], default='')),
            'VatId': _to_clean_string(_pick_first_existing_value(row, ['ДДС Номер', 'VatId', 'TaxNo'], default='')),
            'BankName': _to_clean_string(_pick_first_existing_value(row, ['Банка', 'BankName'], default='')),
            'BankCode': _to_clean_string(_pick_first_existing_value(row, ['Банков код', 'BankCode'], default='')),
            'BankAccount': _to_clean_string(_pick_first_existing_value(row, ['Банкова сметка', 'BankAccount'], default='')),
            'Priority': _to_int(_pick_first_existing_value(row, ['Priority'], default=0), default=0),
            'GroupID': _to_int(_pick_first_existing_value(row, ['GroupID'], default=1), default=1),
            'StatusID': _to_int(_pick_first_existing_value(row, ['StatusID'], default=1), default=1),
            'CountryID': _to_int(_pick_first_existing_value(row, ['CountryID'], default=0), default=0),
        }
        data.append(payload)

    if skipped > 0:
        log(f'⚠ Пропуснати {skipped} невалидни реда.')
    return data


def import_items_excel(log, config=CONFIG):
    if not ensure_database_selected(config, log):
        log('Импортът е отменен: няма избрана база данни.')
        return

    import_file = with_tk_dialog(
        lambda r: filedialog.askopenfilename(
            title='Изберете Excel файл за импорт',
            filetypes=[('Excel файлове', '*.xlsx *.xls'), ('Всички файлове', '*.*')],
            initialdir=os.getcwd(),
            parent=r,
        )
    )
    if not import_file:
        log('Импортът е отменен от потребителя.')
        return

    log(f'✓ Избран файл за импорт: {import_file}')
    log('=== ИМПОРТ ОТ EXCEL КЪМ SQL ===')

    if not os.path.exists(import_file):
        log('✗ Файлът не съществува!')
        return

    try:
        try:
            df = pd.read_excel(import_file, sheet_name='Items', skiprows=config['skiprows'])
        except ValueError:
            df = pd.read_excel(import_file, sheet_name=config['sheet_name'], skiprows=config['skiprows'])

        if not all(col in df.columns for col in EXPECTED_COLUMNS):
            log('✗ Липсват задължителни колони!')
            return

        if df.empty:
            log('✗ Файлът е празен!')
            return

        print('\nПърви 3 реда:')
        print(df.head(3).to_string())

        if not with_tk_dialog(
            lambda r: messagebox.askyesno(
                'Потвърждение',
                f"Ще бъдат заменени записите в '{config['table_name']}' с {len(df)} нови.\nПотвърждавате ли?",
                parent=r,
            )
        ):
            return

        data = build_items_import_payload(df, log)
        if not data:
            return

        conn = connect_with_fallback(config, log)
        if not conn:
            return

        cursor = conn.cursor()

        try:
            cursor.execute(
                f"""
                DECLARE @Targets TABLE (ItemID INT);
                INSERT INTO @Targets SELECT ItemID FROM [dbo].[{config['table_name']}] WHERE [Visible] = 1;
                UPDATE [dbo].[{config['table_name']}] SET [Visible] = 0 WHERE ItemID IN (SELECT ItemID FROM @Targets);
                DELETE FROM [dbo].[{config['table_name']}] WHERE ItemID IN (SELECT ItemID FROM @Targets)
                AND ItemID NOT IN (SELECT ItemID FROM DocumentDetails WHERE ItemID IS NOT NULL)
                AND ItemID NOT IN (SELECT ItemID FROM DocumentTemplateDetails WHERE ItemID IS NOT NULL);
                """
            )

            for i, item in enumerate(data):
                cursor.execute(
                    f"""
                    INSERT INTO [dbo].[{config['table_name']}] (
                        Code, Name, Name2, Measure, Measure2, SalePrice, GroupID, VatRateID,
                        StatusID, VatTermID, Visible, FixedPrice, EcoTax, Priority, IsService,
                        MainItemID, Barcode, Permit
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    tuple(item.values()),
                )

                if (i + 1) % 100 == 0:
                    log(f'  ... {i + 1}/{len(data)}')

            conn.commit()
            log(f'✓ Импортирани {len(data)} записа')
            with_tk_dialog(lambda r: messagebox.showinfo('Успех', f'Импортирани {len(data)} записа!', parent=r))

        except Exception as e:
            conn.rollback()
            log(f'✗ Грешка: {e}')
            raise
        finally:
            conn.close()

    except Exception as e:
        log(f'✗ Грешка при импорт: {e}')


def import_partners_excel(log, config=CONFIG):
    if not ensure_database_selected(config, log):
        log('Импортът е отменен: няма избрана база данни.')
        return

    import_file = with_tk_dialog(
        lambda r: filedialog.askopenfilename(
            title='Изберете Excel файл за импорт на партньори',
            filetypes=[('Excel файлове', '*.xlsx *.xls'), ('Всички файлове', '*.*')],
            initialdir=os.getcwd(),
            parent=r,
        )
    )
    if not import_file:
        log('Импортът е отменен от потребителя.')
        return

    log(f'✓ Избран файл за импорт: {import_file}')
    log('=== ИМПОРТ EXCEL -> INVOICE PRO PARTNERS ===')

    if not os.path.exists(import_file):
        log('✗ Файлът не съществува!')
        return

    try:
        try:
            df = pd.read_excel(import_file, sheet_name='Партньори')
        except ValueError:
            try:
                df = pd.read_excel(import_file, sheet_name='Partners')
            except ValueError:
                df = pd.read_excel(import_file, sheet_name=0)
                log("ℹ Sheet 'Партньори'/'Partners' не е намерен. Използван е първият sheet.")

        if df.empty:
            log('✗ Файлът е празен!')
            return

        print('\nПърви 3 реда:')
        print(df.head(3).to_string())

        if not with_tk_dialog(
            lambda r: messagebox.askyesno(
                'Потвърждение',
                f'Ще бъдат заменени записите в Partners с {len(df)} нови.\nПотвърждавате ли?',
                parent=r,
            )
        ):
            return

        data = build_partners_import_payload(df, log)
        if not data:
            return

        conn = connect_with_fallback(config, log)
        if not conn:
            return

        cursor = conn.cursor()
        partner_id_is_identity = False

        try:
            cursor.execute(
                """
                SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_NAME = 'Partners' AND TABLE_TYPE = 'BASE TABLE'
                """
            )
            if cursor.fetchone()[0] == 0:
                log("✗ Таблица 'Partners' не е намерена в избраната база.")
                return

            cursor.execute(
                """
                UPDATE [dbo].[Partners] SET [Visible] = 0;
                BEGIN TRY
                    DELETE FROM [dbo].[Partners];
                END TRY
                BEGIN CATCH
                END CATCH;
                """
            )

            cursor.execute("SELECT ISNULL(MAX([PartnerID]), 0) FROM [dbo].[Partners]")
            max_partner_id = int(cursor.fetchone()[0] or 0)

            cursor.execute("SELECT COLUMNPROPERTY(OBJECT_ID('dbo.Partners'), 'PartnerID', 'IsIdentity')")
            row = cursor.fetchone()
            partner_id_is_identity = bool(row and row[0] == 1)

            if partner_id_is_identity:
                cursor.execute("SET IDENTITY_INSERT [dbo].[Partners] ON")

            insert_sql = """
                INSERT INTO [dbo].[Partners] (
                    PartnerID, Name, NameEnglish, ContactName, ContactNameEnglish, EMail, Bulstat,
                    VatId, BankName, BankCode, BankAccount, Priority, GroupID, Visible, MainPartnerID,
                    StatusID, IsExported, IsOSSPartner, CountryID, DocumentEndDatePeriod
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, 0, 0, ?, 0)
            """

            inserted = 0

            for i, partner in enumerate(data):
                partner_id = max_partner_id + i + 1
                insert_values = (
                    partner_id,
                    partner['Name'],
                    partner['NameEnglish'],
                    partner['ContactName'],
                    partner['ContactNameEnglish'],
                    partner['EMail'],
                    partner['Bulstat'],
                    partner['VatId'],
                    partner['BankName'],
                    partner['BankCode'],
                    partner['BankAccount'],
                    partner['Priority'],
                    partner['GroupID'],
                    partner_id,
                    partner['StatusID'],
                    partner['CountryID'],
                )

                cursor.execute(insert_sql, insert_values)
                inserted += 1

                if (i + 1) % 100 == 0:
                    log(f'  ... {i + 1}/{len(data)}')

            conn.commit()
            log(f'✓ Импортът приключи. Добавени: {inserted}')
            with_tk_dialog(
                lambda r: messagebox.showinfo(
                    'Успех',
                    f'Импортът приключи успешно.\nДобавени: {inserted}',
                    parent=r,
                )
            )
        except Exception as e:
            conn.rollback()
            log(f'✗ Грешка: {e}')
            raise
        finally:
            if partner_id_is_identity:
                try:
                    cursor.execute("SET IDENTITY_INSERT [dbo].[Partners] OFF")
                    conn.commit()
                except Exception:
                    pass
            conn.close()
    except Exception as e:
        log(f'✗ Грешка при импорт на партньори: {e}')


def _pick_first_existing_value(row, candidates, default=''):
    for col in candidates:
        if col in row and pd.notna(row[col]):
            value = row[col]
            if isinstance(value, str):
                value = value.strip()
            if value != '':
                return value
    return default


def convert_warehouse_partners_excel_for_invoice_pro(log, config=CONFIG):
    source_file = with_tk_dialog(
        lambda r: filedialog.askopenfilename(
            title='Изберете Excel файл от Warehouse Pro (sheet Partners)',
            filetypes=[('Excel файлове', '*.xlsx *.xls'), ('Всички файлове', '*.*')],
            initialdir=os.getcwd(),
            parent=r,
        )
    )
    if not source_file:
        log('Операцията е отменена от потребителя.')
        return

    initial_dir = os.path.dirname(source_file) if os.path.exists(source_file) else os.getcwd()
    initial_name = 'invoice_pro_partners_import_ready.xlsx'
    target_file = with_tk_dialog(
        lambda r: filedialog.asksaveasfilename(
            title='Запази готовия файл за импорт в Invoice Pro',
            initialdir=initial_dir,
            initialfile=initial_name,
            defaultextension='.xlsx',
            filetypes=[('Excel файлове', '*.xlsx'), ('Всички файлове', '*.*')],
            parent=r,
        )
    )
    if not target_file:
        log('Операцията е отменена от потребителя.')
        return

    log(f'✓ Избран входен файл: {source_file}')
    log(f'✓ Избран изходен файл: {target_file}')
    log('=== КОНВЕРТИРАНЕ WAREHOUSE PARTNERS -> INVOICE PRO ПАРТНЬОРИ ===')

    try:
        try:
            df_source = pd.read_excel(source_file, sheet_name='Partners')
        except ValueError:
            df_source = pd.read_excel(source_file, sheet_name=0)
            log("ℹ Sheet 'Partners' не е намерен. Използван е първият sheet.")

        if df_source.empty:
            log('✗ Входният файл е празен.')
            return

        output_rows = []
        generated_partner_ids = 0

        for idx, row in df_source.iterrows():
            partner_id_raw = _pick_first_existing_value(row, ['PartnerID', 'ID', 'MainPartnerID'], default=idx + 1)
            try:
                partner_id = int(float(partner_id_raw))
            except Exception:
                generated_partner_ids += 1
                partner_id = idx + 1

            output_rows.append(
                {
                    'PartnerID': partner_id,
                    'Име': _pick_first_existing_value(row, ['Company', 'Name'], default=''),
                    'Име (EN)': _pick_first_existing_value(row, ['NameEnglish'], default=''),
                    'Лице за контакт': _pick_first_existing_value(row, ['MOL', 'ContactName'], default=''),
                    'Лице за контакт (EN)': _pick_first_existing_value(row, ['ContactNameEnglish'], default=''),
                    'EMail': _pick_first_existing_value(row, ['EMail', 'Email'], default=''),
                    'Булстат': _pick_first_existing_value(row, ['Bulstat'], default=''),
                    'ДДС Номер': _pick_first_existing_value(row, ['TaxNo', 'VatId'], default=''),
                    'Банка': _pick_first_existing_value(row, ['BankName'], default=''),
                    'Банков код': _pick_first_existing_value(row, ['BankCode'], default=''),
                    'Банкова сметка': _pick_first_existing_value(row, ['BankAccount'], default=''),
                    'Priority': _pick_first_existing_value(row, ['Priority'], default=0),
                    'GroupID': _pick_first_existing_value(row, ['GroupID'], default=1),
                    'Visible': _pick_first_existing_value(row, ['Visible'], default=1),
                    'MainPartnerID': partner_id,
                    'StatusID': _pick_first_existing_value(row, ['StatusID'], default=1),
                    'IsExported': _pick_first_existing_value(row, ['IsExported'], default=0),
                    'IsOSSPartner': _pick_first_existing_value(row, ['IsOSSPartner'], default=0),
                    'CountryID': _pick_first_existing_value(row, ['CountryID'], default=0),
                    'DocumentEndDatePeriod': _pick_first_existing_value(row, ['DocumentEndDatePeriod'], default=0),
                }
            )

        df_output = pd.DataFrame(output_rows)

        with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
            df_output.to_excel(writer, index=False, sheet_name='Партньори')

        if generated_partner_ids > 0:
            log(f'⚠ За {generated_partner_ids} записа PartnerID беше генериран автоматично.')

        log(f'✓ Готов файл за импорт: {target_file}')
        log(f'✓ Конвертирани партньори: {len(df_output)}')
        if with_tk_dialog(
            lambda r: messagebox.askyesno(
                'Успех',
                f'Конвертирани са {len(df_output)} партньора.\nДа се отвори ли файлът?',
                parent=r,
            )
        ):
            os.startfile(target_file)
    except Exception as e:
        log(f'✗ Грешка при конвертиране на партньори: {e}')


# Backward-compatible aliases for legacy imports/calls.
prepare_import_data = build_items_import_payload
import_items_from_excel = import_items_excel
import_partners_from_excel = import_partners_excel
convert_warehouse_partners_to_invoice_pro_excel = convert_warehouse_partners_excel_for_invoice_pro
