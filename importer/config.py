import os

from dotenv import load_dotenv

ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
load_dotenv(dotenv_path=ENV_PATH)


def _to_bool(value, default=True):
    if value is None:
        return default
    return str(value).strip().lower() == 'true'


CONFIG = {
    'server': os.getenv('DB_SERVER', '.'),
    'database': os.getenv('DB_DATABASE', ''),
    'table_name': os.getenv('DB_TABLE', 'Items'),
    'excel_file': os.getenv('EXCEL_FILE', None),
    'sheet_name': int(os.getenv('EXCEL_SHEET', '0')),
    'skiprows': int(os.getenv('EXCEL_SKIPROWS', '0')),
    'trusted_connection': _to_bool(os.getenv('DB_TRUSTED_CONNECTION', 'True'), default=True),
    'login_timeout': int(os.getenv('DB_TIMEOUT', '15')),
}

EXPECTED_COLUMNS = ['Код', 'Стока', 'Мярка', 'Цена']
