import pyodbc

try:
    from .config import CONFIG
except ImportError:
    from config import CONFIG


def check_odbc_driver(log):
    drivers = pyodbc.drivers()
    required_driver = 'ODBC Driver 17 for SQL Server'

    if required_driver not in drivers:
        print('\n' + '!' * 60)
        print('ГРЕШКА: Не е инсталиран необходимият ODBC драйвер!')
        print('!' * 60)
        print(f'\nОчакван: {required_driver}')
        print('\nИнсталирани драйвери на тази машина:')
        for i, driver in enumerate(drivers, 1):
            print(f'  {i}. {driver}')
        print('\nМоля инсталирайте: Microsoft ODBC Driver 17 for SQL Server')
        print('Линк за изтегляне:')
        print('https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server')
        print('\nСлед инсталацията рестартирайте програмата.')
        input('\nНатиснете Enter за изход...')
        return False

    log(f'✓ Намерен драйвер: {required_driver}')
    return True


def get_available_databases(config, log):
    try:
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={config['server']};"
            'Trusted_Connection=yes;'
            f"Login Timeout={config['login_timeout']};"
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.databases WHERE state = 0 AND name NOT IN ('master', 'tempdb', 'model', 'msdb') ORDER BY name")
        databases = [row[0] for row in cursor.fetchall()]
        conn.close()
        return databases
    except Exception as e:
        log(f'Не може да се извлече списък с базите: {e}')
        return []


def prompt_database_selection(config, log):
    databases = get_available_databases(config, log)

    if not databases:
        log('✗ Не са намерени достъпни бази данни или липсва връзка със сървъра')
        return False

    print('\n' + '=' * 60)
    print('       НАЛИЧНИ БАЗИ ДАННИ НА СЪРВЪРА')
    print('=' * 60)
    for i, db_name in enumerate(databases, 1):
        marker = ' <-- ТЕКУЩА' if db_name == config['database'] else ''
        print(f'{i:2}. {db_name}{marker}')
    print('=' * 60)
    print('0. Отказ (обратно към менюто)')
    print('-' * 60)

    while True:
        choice = input(f"Изберете база данни (0-{len(databases)}): ").strip()
        if choice == '0':
            return False
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(databases):
                old_db = config['database']
                config['database'] = databases[idx]
                log(f"✓ Сменена база данни: {old_db} -> {config['database']}")
                return True
            print('Невалиден номер!')
        except ValueError:
            if choice in databases:
                old_db = config['database']
                config['database'] = choice
                log(f"✓ Сменена база данни: {old_db} -> {config['database']}")
                return True
            print('Моля въведете валиден номер или име от списъка!')


def ensure_database_selected(config, log):
    if str(config.get('database', '')).strip():
        return True
    log('⚠ Името на базата данни е празно.')
    log('  Изберете база данни от списъка:')
    return prompt_database_selection(config, log)


def check_table_exists(conn, config, table_name=None):
    try:
        table_to_check = table_name or config['table_name']
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_NAME = ? AND TABLE_TYPE = 'BASE TABLE'
            """,
            (table_to_check,),
        )
        exists = cursor.fetchone()[0] > 0
        cursor.close()
        return exists
    except Exception:
        return False


def handle_connection_error(error, config, log):
    error_msg = str(error).lower()
    error_str = str(error)

    if any(x in error_msg for x in ['cannot open database', '4060', 'login failed', '28000', 'недостъпна']):
        log(f"✗ Неуспешно свързване към база '{config['database']}'")
        log(f'  Грешка: {error_str}')
        print('\nВъзможни причини:')
        print('  - Базата данни не съществува')
        print('  - Нямате права за достъп')
        print('  - Грешно име на базата')
        return prompt_database_selection(config, log)

    log(f'✗ Грешка при свързване: {error_str}')
    if 'network' in error_msg or 'server' in error_msg:
        print('\nПроблем с връзката към сървъра.')
        print(f"Проверете дали SQL Server '{config['server']}' е достъпен.")
    return False


def get_connection_string(config):
    driver = 'ODBC Driver 17 for SQL Server'
    return (
        f"DRIVER={{{driver}}};"
        f"SERVER={config['server']};"
        f"DATABASE={config['database']};"
        'Trusted_Connection=yes;'
        f"Login Timeout={config['login_timeout']};"
    )


def connect_with_fallback(config, log):
    if not ensure_database_selected(config, log):
        return None

    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            conn = pyodbc.connect(get_connection_string(config))
            if not check_table_exists(conn, config):
                log(f"⚠ Таблицата '{config['table_name']}' не съществува в база '{config['database']}'!")
                conn.close()
                if not prompt_database_selection(config, log):
                    return None
                continue
            return conn
        except pyodbc.Error as e:
            if attempt < max_attempts - 1:
                if handle_connection_error(e, config, log):
                    continue
                return None
            log('✗ Неуспешно свързване след няколко опита')
            return None
        except Exception as e:
            log(f'✗ Неочаквана грешка: {e}')
            return None
