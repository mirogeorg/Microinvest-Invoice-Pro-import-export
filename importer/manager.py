import sys
from datetime import datetime

import pyodbc

try:
    from .config import CONFIG
    from .mixins import CommonMixin, DatabaseMixin, ExportMixin, ImportMixin
except ImportError:
    from config import CONFIG
    from mixins import CommonMixin, DatabaseMixin, ExportMixin, ImportMixin


class ExcelSQLManager(DatabaseMixin, CommonMixin, ExportMixin, ImportMixin):
    def log(self, message):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f'[{timestamp}] {message}')

    def show_menu(self):
        print('\n' + '=' * 60)
        print('       EXCEL ↔ SQL SERVER МЕНИДЖЪР')
        print('=' * 60)
        print(f"Сървър: {CONFIG['server']} | База: {CONFIG['database']}")
        print(f"Таблица: {CONFIG['table_name']}")
        print('-' * 60)
        print('1. 📤 Експорт Invoice Pro Стоки + свързани таблици → Excel')
        print('2. 📤 Експорт Invoice Pro Партньори → Excel')
        print('3. 📤 Експорт Warehouse Pro партньори -> Excel')
        print('4. 📥 Импорт Excel → Invoice Pro Items')
        print('5. 🗃️ Смяна на база данни')
        print('6. 🚪 Изход')
        print('=' * 60)

    def run(self):
        self.log('Стартиране на Excel-SQL Manager...')

        if not self.check_odbc_driver():
            sys.exit(1)

        try:
            test_conn = pyodbc.connect(self.get_connection_string(), timeout=3)
            test_conn.close()
            self.log(f"✓ Успешна връзка с {CONFIG['database']}")
        except Exception:
            self.log(f"⚠ Неуспешна първоначална връзка с {CONFIG['database']}")
            self.log('  Ще бъде предложен избор на база при първа операция')

        while True:
            self.show_menu()
            choice = input('Изберете (1-6): ').strip()

            if choice == '1':
                self.export_items_to_excel()
            elif choice == '2':
                self.export_partners_to_excel()
            elif choice == '3':
                self.export_warehouse_pro_partners_to_excel()
            elif choice == '4':
                self.import_items_from_excel()
            elif choice == '5':
                self.prompt_database_selection()
            elif choice == '6':
                self.log('Изход...')
                break
            else:
                print('Невалидна опция!')
