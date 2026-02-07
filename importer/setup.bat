@echo off
chcp 65001 >nul
call setupenv.bat
title InvoicePro Excel Importer Setup
echo ===========================================
echo  InvoicePro Excel to SQL Importer - Setup
echo ===========================================
echo.

:: Проверка за администраторски права
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] Скриптът трябва да се стартира като Администратор!
    echo     Моля, кликнете с десен бутон и изберете "Run as administrator"
    pause
    exit /b 1
)


set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%venv"
set "PYTHON_CMD=python"

echo [1/5] Проверка за Python...
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo     Python не е намерен. Опит за инсталация чрез WinGet...
    
    where winget >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ГРЕШКА] WinGet не е наличен! Моля, инсталирайте Python ръчно от python.org
        pause
        exit /b 1
    )
    
    echo     Инсталиране на Python 3.12...
    winget install --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
    
    if %errorlevel% neq 0 (
        echo [ГРЕШКА] Неуспешна инсталация на Python.
        pause
        exit /b 1
    )
	winget upgrade Python.Python.3.12
    
    :: Обновяване на PATH
    refreshenv
    set "PATH=%PATH%;C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312;C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python311\Scripts"
) else (
    echo     Python е наличен.
    python --version
)

echo.
echo [2/5] Създаване на Virtual Environment...
if exist "%VENV_DIR%" (
    echo     Съществуващ venv е намерен. Използване на съществуващия.
) else (
    python -m venv "%VENV_DIR%"
    if %errorlevel% neq 0 (
        echo [ГРЕШКА] Неуспешно създаване на virtual environment.
        pause
        exit /b 1
    )
    echo     Virtual Environment създаден успешно.
)

echo.
echo [3/5] Активиране на Virtual Environment...
call "%VENV_DIR%\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo [ГРЕШКА] Неуспешно активиране на venv.
    pause
    exit /b 1
)

echo.
echo [4/5] Инсталация на зависимости...
if not exist "%SCRIPT_DIR%requirements.txt" (
    echo [ГРЕШКА] Липсва requirements.txt файл!
    pause
    exit /b 1
)

pip install --upgrade pip
pip install -r "%SCRIPT_DIR%requirements.txt"
if %errorlevel% neq 0 (
    echo [ГРЕШКА] Неуспешна инсталация на пакети.
    pause
    exit /b 1
)
echo     Зависимостите са инсталирани успешно.
pause