# Sharjeel Web Automation - Windows Desktop App

A Windows desktop application for automated transcript downloading.

## Features

- **Reset Login**: Clear saved session to re-login
- **Excel File Selection**: Pick your student IDs file
- **Column Name**: Specify which column contains student IDs
- **Output Directory**: Choose where to save PDFs
- **Run Automation**: One-click batch download

## Setup

```bash
# Install dependencies
pip install flet openpyxl playwright requests

# Install Playwright browsers
python -m playwright install chromium
```

## Run

```bash
python app.py
```

## Build Windows Executable

```bash
pip install flet
flet pack app.py --platform windows
```

Or for a standalone .exe:

```bash
flet build app.py --platform windows
```

This creates a distributable `.exe` file in the `build/windows` folder.

## Usage

1. **Reset Login**: Click to clear any saved session (forces fresh login)
2. **Excel File**: Click folder icon to browse or enter path manually
3. **Column**: Enter the Excel column header name (e.g., "Roll No", "ID")
4. **Output**: Click folder icon to select download destination
5. **Run Automation**: Click the big button to start!

## Requirements

- Windows 10/11
- Python 3.11+
- Chrome/Chromium browser (installed automatically by Playwright)
