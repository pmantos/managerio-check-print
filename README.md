# Manager.io → Voucher Check Printer (Generic/Text)

Render Manager.io “Printable Checks” custom report text into QuickBooks-style voucher
check stock (check on top + 2 vouchers). Sends raw ESC/P text to a Generic/Text
printer and/or writes a preview.

## Features
- 80×54 fixed grid (10 CPI, 6 LPI) aligned for QB voucher stock
- Safe ASCII output (CP437 default; CP850 optional)
- Smart parsing with heuristics **plus** optional `|` field delimiters for perfect results
- Windows RAW printing via `pywin32` (or preview only)

## Usage
py -3 "print_voucher_checks_text.py" ^
  --input "C:\path\to\print_voucher_checks.txt" ^
  --printer "EPSON XP-7100 (GENERIC TEXT)"

##Useful flags:
  --no-print (preview only)
  --rename (rename input to *_printed_YYYYMMDD_HH_MM.txt on success)
  --encoding cp850 --charset 850 (switch code page)
  --cal (print calibration grid and exit)

## Install

```bash
# optional: create a venv
python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
