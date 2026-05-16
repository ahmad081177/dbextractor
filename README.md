# Access DB Documenter

Access DB Documenter is a Windows tool that opens a Microsoft Access database and creates a Word document with screenshots of:

- each table in Design View
- each table in Datasheet View
- the Relationships diagram

The output is designed for simple database documentation and is suitable for students.

## What It Does

For each table, the tool adds:

- table name
- a description placeholder
- Design View screenshot
- Datasheet View screenshot

At the end, it also adds a Relationships section.

If Access cannot open the Relationships screen automatically, the app asks the user to open it manually in Access, then continues.

## Requirements

### To run with Python

- Windows
- Microsoft Access installed
- Python 3.x installed

### To run the EXE

- Windows
- Microsoft Access installed
- Python is not needed

## Run with Python

Install dependencies:

```powershell
pip install -r requirements.txt
```

Run the GUI:

```powershell
python documenter.py
```

Or run from the command line with a database path:

```powershell
python documenter.py path\to\your\database.accdb
```

## Build the EXE

Run:

```powershell
build.bat
```

This will:

1. install or upgrade required packages
2. build a standalone executable with PyInstaller

Output file:

```text
dist\AccessDBDocumenter.exe
```

## Download the EXE

GitHub Releases page:

https://github.com/ahmad081177/dbextractor/releases/

You can upload the built EXE there and students can download it from that page.

## How Students Should Use It

Students do not need command-line parameters.

They should:

1. double-click `AccessDBDocumenter.exe`
2. choose the `.accdb` file
3. click `Generate Documentation`
4. if prompted, open `Database Tools > Relationships` in Access and click `OK`

The generated `.docx` file is saved next to the database file.

## Notes

- Table screenshots are automatic.
- Relationships may need one manual step, depending on the Access installation.
- If the output `.docx` file is already open, the tool saves a new copy with a timestamp.

## Project Files

- `documenter.py` — main application
- `requirements.txt` — Python dependencies
- `build.bat` — EXE build script

## Dependencies

- `pywin32`
- `python-docx`
- `Pillow`
- `pyinstaller`
