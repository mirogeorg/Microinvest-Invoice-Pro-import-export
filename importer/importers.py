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


def prepare_import_data(df, log):
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


def import_items_from_excel(log, config=CONFIG):
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

        data = prepare_import_data(df, log)
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
