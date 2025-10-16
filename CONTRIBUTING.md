# Contributing to managerio-check-print

Thanks for helping improve **Manager.io → Voucher Check Printer**!  
This doc explains how to set up your environment, run the project, and submit quality pull requests.

## Quick Start (Windows)

> This project targets Windows + Generic/Text printers (RAW ESC/P).

### Prereqs
- Python 3.10+ (`py -3 --version`)
- Git
- (Optional) A Windows Generic/Text printer queue for end-to-end tests
- (Optional) Make sure you can install wheels: `pip install --upgrade pip wheel`

### Setup
```bat
git clone https://github.com/pmantos/managerio-check-print.git
cd managerio-check-print

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt

Run the Script
py -3 "print_voucher_checks_text.py" ^
  --input "C:\path\to\print_voucher_checks.txt" ^
  --printer "EPSON XP-7100 (GENERIC TEXT)"

Useful flags:

--no-print (preview only; writes {stem}_print.txt)

--rename (renames input to {stem}_printed_YYYYMMDD_HH_MM.txt on success)

--encoding cp850 --charset 850 (switch code page)

--cal (prints an 80×54 calibration grid and exits)

See README.md for details on | as end-of-field markers and address line breaks.