# SmartPayer PDF Matcher

A Python automation tool that scans a `Generated_Letters` directory for `.docx` files missing a matching `.pdf`, then converts them automatically using Microsoft Word COM via `pywin32`.

## How It Works

Every letter is expected to exist as a pair:

| File | Purpose |
|---|---|
| `<name>_with_tables.docx` | Source Word document |
| `<name>_letter.pdf` | PDF output for distribution |

The tool finds any `.docx` without a matching `.pdf` and converts them in bulk using the locally installed Microsoft Word application.

## Prerequisites

- Windows OS with **Microsoft Word** installed
- **Python 3.8+** on PATH
- **pywin32** library

Run `install_prerequisites.bat` to install pywin32 automatically.

## Folder Structure

```
SmartPayer PDF Matcher/
├── Generated_Letters/          ← Letter files live here
│   ├── Smartpayer April 2026 FOOBAR_with_tables.docx
│   ├── Smartpayer April 2026 FOOBAR_letter.pdf
│   └── ...
├── convert_missing_pdfs.py     ← Main conversion script
├── install_prerequisites.bat   ← One-time setup
├── run_converter.bat           ← Run the converter
└── README.md
```

## Usage

### 1. Install prerequisites (first time only)

```
install_prerequisites.bat
```

Run as Administrator if prompted.

### 2. Run the converter

```
run_converter.bat
```

Or directly:

```
python convert_missing_pdfs.py
```

### 3. Check the log

After each run a timestamped log is saved inside `Generated_Letters\`:

```
Generated_Letters\conversion_log_20260612_143022.txt
```

The log contains:
- All `.docx` files found missing a `.pdf`
- Conversion result for each file (success / failure)
- Summary counts

## Error Messages

| Message | Cause | Fix |
|---|---|---|
| `Generated_Letters folder not found!` | Script is not in the right directory | Move the scripts into the folder that contains `Generated_Letters\` |
| `pywin32 is not installed` | Missing dependency | Run `install_prerequisites.bat` |
| `Python is not found in PATH` | Python not installed or not on PATH | Install Python and check "Add to PATH" |
| `Failed to convert <file>` | Word could not open or export the file | Check the file is not open/locked in Word |

## Naming Convention

The script resolves pairs by stripping known suffixes from the filename stem:

- `_with_tables` → docx side
- `_letter` → pdf side

Both sides share the same base name. Temporary Word lock files (`~$...`) are automatically ignored.
