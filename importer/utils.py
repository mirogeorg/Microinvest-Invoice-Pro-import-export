import ctypes
from copy import copy

import pandas as pd
import pyodbc
import tkinter as tk
from openpyxl.worksheet.datavalidation import DataValidation


def with_tk_dialog(func):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.update_idletasks()
    try:
        return func(root)
    finally:
        root.destroy()
        bring_console_to_front()


def bring_console_to_front():
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            ctypes.windll.user32.ShowWindow(hwnd, 1)
    except Exception:
        pass


def transliterate(text):
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


def auto_adjust_column_width(worksheet):
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width


def format_header_bold(worksheet):
    for cell in worksheet[1]:
        new_font = copy(cell.font)
        new_font.bold = True
        cell.font = new_font


def parse_id_value(value):
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


def add_dropdown_validation(worksheet, column_letter, source_sheet, source_column, start_row, end_row, allow_blank=True):
    max_source_row = 1000
    formula = f"='{source_sheet}'!${source_column}$2:${source_column}${max_source_row}"
    dv = DataValidation(type='list', formula1=formula, allow_blank=allow_blank)
    dv.error = 'Моля изберете стойност от списъка'
    dv.errorTitle = 'Невалидна стойност'
    dv.prompt = f'Изберете от {source_sheet}'
    dv.promptTitle = 'Справочник'
    cell_range = f'{column_letter}{start_row}:{column_letter}{end_row}'
    dv.add(cell_range)
    worksheet.add_data_validation(dv)


def get_access_odbc_driver():
    drivers = pyodbc.drivers()
    access_drivers = [
        'Microsoft Access Driver (*.mdb, *.accdb)',
        'Microsoft Access Driver (*.mdb)',
    ]
    for driver in access_drivers:
        if driver in drivers:
            return driver
    return None
