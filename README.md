# Microinvest Invoice Pro Import/Export

## Защо е създаден инструментът

Нуждата от този инструмент идва от преминаването към евро, при което трябва масово да се актуализират много цени. През стандартния интерфейс на Microinvest Invoice Pro това е бавно и неудобно за големи обеми промени.

## Как работи

- Инструментът се стартира на машината, където е базата данни на Invoice Pro (обичайно MS SQL Express 2017).
- Може да експортира таблица `Items` (Номенклатура на стоки) в Excel и след това да импортира обратно актуализирани данни.
- При експорт добавя и sheet `Партньори` с данни от таблица `Partners` (видимите записи).
- При импорт съществуващите използвани номенклатури първо се маркират като невидими `Invisible = True`, след което се зареждат новите записи.
- По този начин съществуващите документи не се засягат и продължават да използват старите наименования и цени.

## Project structure

```text
Microinvest-Invoice-Pro-import-export/
|- README.md
|- requirements.txt
|- importer/
|  |- .env.example
|  |- config.py
|  |- manager.py
|  |- main.py
|  |- mixins/
|  |  |- common_mixin.py
|  |  |- database_mixin.py
|  |  |- export_mixin.py
|  |  |- import_mixin.py
|  |- app_config.json (created/updated after run)
|- docs/
```

## Setup path 1: Simple (winget)

1. Install Git (if not installed):
   ```bat
   winget install Git.Git
   ```
2. Clone the repository and enter it:
   ```bat
   git clone https://github.com/mirogeorg/Microinvest-Invoice-Pro-import-export
   cd Microinvest-Invoice-Pro-import-export
   ```
3. Check Python:
   ```bat
   where python
   python --version
   ```
4. If Python is missing, install Python 3.12:
   ```bat
   winget install --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
   winget upgrade Python.Python.3.12
   ```
5. Create and activate a virtual environment from the repository root:
   ```bat
   python -m venv .venv
   .venv\Scripts\activate.bat
   ```
6. Install dependencies from root `requirements.txt`:
   ```bat
   pip install --upgrade pip
   pip install -r requirements.txt
   ```
7. Run the app:
   ```bat
   python importer\main.py
   ```

## Setup path 2: Advanced (pyenv)

Use this if you work with multiple Python versions and want per-project version control.

1. Install required tools:
   ```bat
   winget install Git.Git
   winget install pyenv-win.pyenv-win
   ```
2. Clone and enter the repository:
   ```bat
   git clone https://github.com/mirogeorg/Microinvest-Invoice-Pro-import-export
   cd Microinvest-Invoice-Pro-import-export
   ```
3. Install and pin Python for this project:
   ```bat
   pyenv install 3.12.8
   pyenv local 3.12.8
   python --version
   ```
4. Create and activate virtual environment:
   ```bat
   python -m venv .venv
   .venv\Scripts\activate.bat
   ```
5. Install dependencies and run:
   ```bat
   pip install --upgrade pip
   pip install -r requirements.txt
   python importer\main.py
   ```

If `pyenv` is not recognized, close and reopen the terminal.

## Notes

- The application entry point is `importer/main.py`.
- Rename `importer/.env.example` to `importer/.env` and adjust values if needed.
- Configuration is loaded from `importer/.env` if that file exists.
- SQL Server access requires `ODBC Driver 17 for SQL Server` to be installed.
