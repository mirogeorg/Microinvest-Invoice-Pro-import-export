# Krisi InvoicePro Importer

## Windows setup (human-friendly)

These steps follow what `krisi_importer/setup.bat` does, but written as manual instructions.

1. Open **Command Prompt** or **PowerShell** as **Administrator**.
2. Go to the project folder:
   ```bat
   cd D:\Dropbox\DropMiro\PRG\python\Krisi_invoice_pro_importer\krisi_importer
   ```
3. Check Python:
   ```bat
   where python
   python --version
   ```
4. If Python is missing, install Python 3.12 with WinGet:
   ```bat
   winget install --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
   winget upgrade Python.Python.3.12
   ```
   Then close and reopen the terminal (or refresh PATH) before continuing.
5. Create virtual environment (only once):
   ```bat
   python -m venv venv
   ```
6. Activate virtual environment:
   ```bat
   venv\Scripts\activate.bat
   ```
7. Install dependencies:
   ```bat
   pip install --upgrade pip
   pip install -r ..\requirements.txt
   ```

## Notes

- `setup.bat` checks for `krisi_importer\requirements.txt`, but in this repository the file is at the root (`requirements.txt`). That is why the manual command above uses `..\requirements.txt`.
- If `venv` already exists, you can skip step 5.
