import sys
from datetime import datetime

import pyodbc

try:
    from .config import CONFIG
    from .db import check_odbc_driver, get_connection_string, prompt_database_selection
    from .export_service import export_items_excel, export_partners_excel, export_warehouse_partners_excel
    from .import_service import import_items_excel
except ImportError:
    from config import CONFIG
    from db import check_odbc_driver, get_connection_string, prompt_database_selection
    from export_service import export_items_excel, export_partners_excel, export_warehouse_partners_excel
    from import_service import import_items_excel


def log(message):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'[{timestamp}] {message}')


def show_menu(config=CONFIG):
    print('\n' + '=' * 60)
    print('       EXCEL ↔ SQL SERVER МЕНИДЖЪР')
    print('=' * 60)
    print(f"Сървър: {config['server']} | База: {config['database']}")
    print(f"Таблица: {config['table_name']}")
    print('-' * 60)
    print('1. 📤 Експорт Invoice Pro Стоки + свързани таблици → Excel')
    print('2. 📤 Експорт Invoice Pro Партньори → Excel')
    print('3. 📤 Експорт Warehouse Pro партньори -> Excel')
    print('4. 📥 Импорт Excel → Invoice Pro Items')
    print('5. 🗃️ Смяна на база данни')
    print('6. 🚪 Изход')
    print('=' * 60)


def run_app(config=CONFIG):
    log('Стартиране на Excel-SQL Manager...')

    if not check_odbc_driver(log):
        sys.exit(1)

    try:
        test_conn = pyodbc.connect(get_connection_string(config), timeout=3)
        test_conn.close()
        log(f"✓ Успешна връзка с {config['database']}")
    except Exception:
        log(f"⚠ Неуспешна първоначална връзка с {config['database']}")
        log('  Ще бъде предложен избор на база при първа операция')

    while True:
        show_menu(config)
        choice = input('Изберете (1-6): ').strip()

        if choice == '1':
            export_items_excel(log, config)
        elif choice == '2':
            export_partners_excel(log, config)
        elif choice == '3':
            export_warehouse_partners_excel(log, config)
        elif choice == '4':
            import_items_excel(log, config)
        elif choice == '5':
            prompt_database_selection(config, log)
        elif choice == '6':
            log('Изход...')
            break
        else:
            print('Невалидна опция!')
