# -*- coding: utf-8 -*-
"""
print_voucher_checks_text.py

Author: Peter Mantos (Mantos I.T. Consulting, Inc.) with assistance from ChatGPT
Date: 15-Oct-2025

Purpose
-------
Manager.io does not print checks directly. This utility takes a Manager.io
**Custom Report** (“Printable Checks”) that you "print" to a text file via the
Windows **Generic/Text** printer, then renders the data into a fixed-grid
layout for QuickBooks-style check stock (check on top + 2 vouchers) and:
  • saves a preview text file, and/or
  • sends raw text to a Generic/Text (ESC/P) printer.

Envelope / stock
----------------
The page grid is 80×54 characters using 10 CPI and 6 LPI.

Tuned for Intuit/QuickBooks “Check on Top” stock (with two lower vouchers).
Address window alignment verified against Redi-Seal #24539 (narrower top window).
Also shows through #24529 (taller window), but #24539 leaves more vertical margin.

Encoding / glyphs
-----------------
Raw text to the printer is sent in legacy **code page 437** by default, and the
ESC/P printer “character table” is forced to 437 (ESC t 0). Smart quotes,
en-/em-dashes, and diacritics are sanitized so names like “Grasso’s” render as
ASCII “Grasso's” on simple text printers. You can switch to CP850 with
`--charset 850 --encoding cp850` if preferred.

About `py -3` (optional)
------------------------
On Windows, `py` is the Python Launcher.
  • `py -3 script.py` asks the launcher to run **Python 3** (helpful when multiple
    Python versions exist).
  • Optional: you can pin a minor version, e.g. `py -3.12`.
On macOS/Linux use `python3 script.py`.

------------------------------------------------------------------------
Command line
------------------------------------------------------------------------
USAGE (Windows examples):
  py -3 print_voucher_checks_text.py [OPTIONS]
  py -3 print_voucher_checks_text.py --input "C:\\in\\finance\\companyname_checks.txt"
  py -3 print_voucher_checks_text.py --cal
  py -3 print_voucher_checks_text.py --rename
  py -3 print_voucher_checks_text.py --no-print
  py -3 print_voucher_checks_text.py --printer "EPSON XP-7100 (GENERIC TEXT)"

Parameters (all optional unless noted)
--------------------------------------
--input PATH           Input text from Manager.io custom report.
                       Default: print_voucher_checks.txt
                       • Output names are derived from this:
                         {stem}_print.txt   (preview; overwritten)
                         {stem}_debug.txt   (first lines of raw on parse fail; overwritten)
                         {stem}_printed_YYYYMMDD_HH_MM.txt (rename target on success)

--printer NAME         Windows printer queue to receive RAW text.
                       Default: EPSON XP-7100 (GENERIC TEXT)

--no-print             Do NOT send to printer (preview file only). Default: printing ON.

--rename               After a successful print, rename the input file to
                       {stem}_printed_YYYYMMDD_HH_MM.txt. Default: OFF.

--encoding NAME        Code page used for WritePrinter(). Default: cp437
                       Common options: cp437, cp850, cp1252

--charset N            ESC/P “character table” (ESC t n):
                         0=437, 2=850, 3=860, 4=863, 5=865
                       Default: 437

--cal                  Print an 80×54 calibration grid (ruler) and EXIT.
                       • No input file needed.
                       • No output or rename performed.

Notes
-----
• In Manager.IO "Checks" are "Payments" .  Each Payment has a PAYEE and a contact type.
  It is recommended to make a "Supplier" (enable suppliers using CUSTOM if needed) as
  Suppliers have addresses that can be accessed from the custom report and therefore
  captured (by printing to a text file) for use by this script to print checks.
  
• The companion Manager.io custom report should include: Date, Contact, Description,
  Credit, and Supplier → Address. Filter on the bank account and (optionally)
  Payment.Reference “contains TBP” to only pull items “To Be Printed”.

• **Recommended field delimiters (`|`)**
  As Manager.IO has no, export function for reports, such as to a CSV file, the printed
     (To generic text) runs fields together making it hard to distinquish for example,
      the PAYEE name from the Memo/Description, or even the amount.     
  For perfect parsing, end each field with a pipe `|` in your Manager report output:
  - **Contact (Supplier Name):** `Century Link|`
  - **Description (Memo):** `505-291-1047|`
  - **Supplier Address:** put `|` between *and after* every address line, e.g.  
    `P.O. Box 2961| Phoenix, AZ| 85062-2961|`
  Using `|` acts like an end-of-field marker and guarantees the payee, memo,
  and address lines are separated exactly as intended. (The parser still has
  fallbacks when `|` is omitted, but `|` yields the most reliable results.)

• The companion Manager.io custom report should include: Date, Contact, Description,
  Credit, and Supplier → Address. Filter on the bank account and (optionally)
  Payment.Reference “contains TBP” to only pull items “To Be Printed”.

• If parsing fails, a debug dump of the first lines of the raw text is written to
  {stem}_debug.txt.

• On success, the preview file is {stem}_print.txt and can be saved/emailed.


# -----------------------------------------------------------------------------
# HOW TO BUILD THE “PRINTABLE CHECKS” CUSTOM REPORT (TBP-ONLY) IN MANAGER.IO
# -----------------------------------------------------------------------------
# Goal: produce a text output that this script can parse into voucher checks.
#
# 1) Create the report
#    Reports → Custom Reports → New Custom Report
#      Name:  Printable Checks
#      (Optional) Description: Used to print checks on QB voucher stock
#
# 2) Time window & basis
#      From / To: as desired (TIP: Put in a VERY broad range as only payments with
           TBD in the reference field will be selected)
#      Accounting method: Accrual basis (Cash is fine too)
#
# 3) Columns to SELECT (in this order)
#      • Date
#      • Contact                      # becomes the Payee
#      • Description                  # becomes the Memo
#      • Credit                       # check amount as a positive number
#      • Supplier → Address           # supplier mailing address
#
# 4) WHERE filters (TBP-only, outbound checks from a specific bank)
#      • Bank account → Name → contains → BOA      # or exact name of your checking account
#      • AccountAmount → is less than → 0          # outflows; or use Credit > 0
#      • Payment → Reference → contains → TBP      # “To Be Printed” tag
#
# 5) (Optional) ORDER BY
#      • Date (ascending)
#
# 6) Print to a text file
#      Open the report → Print → choose a “Generic / Text Only” printer.
#      (if you used a path/filename.txt as your "port" , it will make a file there)
#      Save the printed text as something like:
#        C:\Users\PeterAdmin\Documents\MITC\Finance\Manager.IO\print_voucher_checks.txt
#
# 7) Strongly recommended delimiters (the “|” bar)
#      Manager.io prints report columns without true CSV delimiters. To make
#      parsing bullet-proof, end each field with a pipe “|”:
#        • Supplier Name (Contact): append “|” to the name in the Supplier record
#            Example:  Century Link|
#        • Payment Description (Memo): end with “|”
#            Example:  505-291-1047|
#        • Supplier Address: use “|” between lines (and an optional trailing “|”)
#            Example:  P.O. Box 2961| Phoenix, AZ| 85062-2961|
#      Notes:
#        - “|” acts as an end-of-field marker and ensures perfect payee/memo split
#          and address line breaks for window envelopes.
#        - This script still includes heuristics when bars are omitted (e.g.,
#          phone/date patterns for memo splitting; “City, ST ZIP” for 1-line
#          addresses), but “|” yields the most reliable results.
#
# 8) Example command to render/print
#      py -3 "C:\Users\PeterAdmin\Documents\ABQ Casita LLC\Finance\print_voucher_checks_text.py" ^
#          --input "C:\Users\PeterAdmin\Documents\MITC\Finance\Manager.IO\print_voucher_checks.txt"
#      (Add --no-print to preview only, or --printer "Your Generic/Text printer".)
# -----------------------------------------------------------------------------


"""

import os, re, sys, argparse
from decimal import Decimal, InvalidOperation
from datetime import datetime
import unicodedata

# --------------------------- Layout / printer grid ----------------------------
CPL = 80               # characters per line (10 CPI across ~8")
LINES_PER_PAGE = 54    # lines per page (6 LPI across ~9")
CPI = 10.0
LPI = 6.0

ESC  = "\x1b"
INIT = ESC + "@"
CPI10 = ESC + "P"      # 10 CPI
LPI6  = ESC + "2"      # 1/6" line spacing
FF   = "\x0c"
CRLF = "\r\n"

# line/column helpers (1-based to 0-based)
R = lambda n: max(0, int(n) - 1)
C = lambda n: max(0, int(n) - 1)

# Margins (inches)
MARGIN_LEFT_IN  = 0.50
MARGIN_TOP_IN   = 0.50
def col_in(x_in): return int(round((MARGIN_LEFT_IN + float(x_in)) * CPI))
def row_in(y_in): return int(round((MARGIN_TOP_IN  + float(y_in)) * LPI))

# ---- Absolute positions tuned for QB 1-check + 2-voucher stock + #24539 env
ROW_DATE,   COL_DATE       = R(6),  C(50)
ROW_PAYEE,  COL_PAYEE      = R(6),  C(9)
ROW_AMT_NUM, AMT_RIGHT_COL = R(7),  73
ROW_WORDS,  COL_WORDS      = R(9),  C(7)
ROW_ADDR_START, COL_ADDR   = R(12), C(6)
ROW_MEMO,   COL_MEMO       = R(17), C(7)

CHECK_H_IN   = 3.50
ROW_STUB1    = row_in(CHECK_H_IN)
ROW_STUB2    = row_in(CHECK_H_IN * 2)
COL_STUB_LABEL = col_in(0.20)
COL_STUB_VAL   = col_in(1.20)

# ------------------------------- Parsing rules -------------------------------
MAX_PAYEE_LINES = 2
AMOUNT_RE = re.compile(r"^\s*[-+]?\d{1,3}(?:,\d{3})*(?:\.\d{2})\s*$")
DATE_RE   = re.compile(r"^\s*\d{2}/\d{2}/\d{4}\s*$")

SMART_MAP = {
    "\u2019": "'", "\u2018": "'", "\u201c": '"', "\u201d": '"',
    "\u2013": "-", "\u2014": "-", "\u00a0": " ",
}

# Amount finder that also handles "jammed" cases like "41.79P.O. Box …"
AMT_TIGHT_RE = re.compile(
    r"""
    (?<!\d)
    (-?(?:\d{1,3}(?:,\d{3})*|\d+)\.\d{2})
    (?!\d)
    """,
    re.VERBOSE,
)

ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")

def asciiize(s: str) -> str:
    if not s: return s
    for k, v in SMART_MAP.items():
        s = s.replace(k, v)
    s = unicodedata.normalize("NFKD", s)
    return s.encode("ascii", "ignore").decode("ascii")

def sanitize(s: str) -> str:
    if not s: return s
    return s.translate(str.maketrans(SMART_MAP))

def split_fields_by_pipes(s: str):
    """Split on '|' trimming pieces; preserve trailing empty if line ended with '|'."""
    parts = [p.strip() for p in s.split("|")]
    if s.endswith("|"):
        return parts
    while parts and not parts[-1]:
        parts.pop()
    return parts

# ----------------------------- Page helpers ----------------------------------
def put(linebuf, row, col, text):
    if row < 0 or row >= len(linebuf): return
    text = asciiize(text or "")
    line = linebuf[row]
    if len(line) < col:
        line += " " * (col - len(line))
    new = list(line)
    for i, ch in enumerate(text):
        p = col + i
        if p >= CPL: break
        if p >= len(new):
            new.extend(" " * (p - len(new) + 1))
        new[p] = ch
    linebuf[row] = "".join(new)

def put_right(linebuf, row, right_col_1based, text):
    right_col = C(right_col_1based)
    start_col = max(0, right_col - len(text) + 1)
    put(linebuf, row, start_col, text)

def blank_page(): return [""] * LINES_PER_PAGE

def trim_trailing_blank_lines(page):
    while page and not page[-1].strip():
        page.pop()
    return page

# ----------------------------- Money to words --------------------------------
def money_words(val: Decimal) -> str:
    n = int(val)
    cents = int(round((val - n) * 100))  # robust against Decimal quirks
    small = ["zero","one","two","three","four","five","six","seven","eight","nine",
             "ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen",
             "seventeen","eighteen","nineteen"]
    tens = ["","ten","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"]
    def up_to_999(x):
        parts = []
        if x >= 100:
            parts += [small[x//100], "hundred"]; x %= 100
        if x >= 20:
            parts.append(tens[x//10])
            if x % 10: parts.append(small[x%10])
        elif x > 0:
            parts.append(small[x])
        elif not parts:
            parts.append("zero")
        return " ".join(parts)
    def chunk(x):
        out = []
        if x >= 1_000_000_000: out += [up_to_999(x//1_000_000_000),"billion"]; x%=1_000_000_000
        if x >= 1_000_000:     out += [up_to_999(x//1_000_000),"million"];   x%=1_000_000
        if x >= 1000:          out += [up_to_999(x//1000),"thousand"];       x%=1000
        if x > 0:              out += [up_to_999(x)]
        return " ".join(out) if out else "zero"
    words = chunk(n) + (" dollar" if n == 1 else " dollars")
    return (words + f" and {cents:02d}/100").capitalize()

# ------------------------------- Parsing -------------------------------------
def read_text(path):
    """
    Robust text read (cp1252/utf-8/latin-1/utf-16) with NBSP/control cleanup.
    """
    tried = ("cp1252", "utf-8-sig", "latin-1", "utf-16-le", "utf-16-be")
    for enc in tried:
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                raw = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        with open(path, "r", encoding="latin-1", errors="replace") as f:
            raw = f.read()
    raw = raw.replace("\ufeff", "").replace("\u00a0", " ")
    raw = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", raw)
    return [ln.rstrip("\r\n") for ln in raw.splitlines()]

def normalize_addr(addr_lines):
    """
    1) If '|' appears anywhere, split on it (hard breaks).
    2) Otherwise keep raw lines; if we still have a single long line, try to
       split before the state/ZIP: "... City, ST 12345" -> two lines.
    """
    parts = []
    for ln in addr_lines:
        ln = (ln or "").strip()
        if not ln:
            continue
        if "|" in ln:
            parts.extend([p.strip() for p in ln.split("|") if p.strip()])
        else:
            parts.append(ln)

    # Try a nice 2-line break if we only have one long line
    if len(parts) == 1:
        s = parts[0]
        mzip = ZIP_RE.search(s)
        if mzip:
            left = s[:mzip.start()].rstrip()
            right = s[mzip.start():].lstrip()
            k = left.rfind(",")
            if k != -1:
                line1 = left[:k].strip()
                line2 = (left[k+1:].strip() + " " + right).strip()
                parts = [line1, line2]
    return parts

def parse_payload_with_pipes(first_payload: str, follow_lines: list):
    """
    Parse one record when the user uses pipes as field terminators.
    Layout after the date (all on the same first line):
        Contact| Memo| Amount[may be jammed to addr head]  Addr... (with pipes)
    """
    tokens = split_fields_by_pipes(first_payload)

    payee = tokens[0].strip() if tokens else ""
    memo  = tokens[1].strip() if len(tokens) >= 2 else ""

    remainder = ""
    if len(tokens) >= 3:
        remainder = "|".join(tokens[2:]).strip()

    # Find the amount in remainder; if not there, peek into the next nonblank line.
    amount = None
    addr_head = ""
    m = AMT_TIGHT_RE.search(remainder or "")
    if not m:
        j = 0
        while j < len(follow_lines) and not follow_lines[j].strip():
            j += 1
        if j < len(follow_lines):
            m = AMT_TIGHT_RE.search(follow_lines[j])
            if m:
                remainder = follow_lines[j]
                # consume that line from the caller's list
                del follow_lines[: j + 1]

    if not m:
        return None  # let caller fall back

    try:
        amount = Decimal(m.group(1).replace(",", ""))
    except InvalidOperation:
        return None

    addr_head = remainder[m.end():].strip()

    # Collect address lines: start with head, then following lines until blank/new record
    addr_lines = []
    if addr_head:
        addr_lines.append(addr_head)
    for ln in follow_lines:
        s = ln.strip()
        if not s:
            break
        if DATE_RE.match(s) or AMT_TIGHT_RE.search(s):
            break
        addr_lines.append(s)

    addr = normalize_addr(addr_lines)
    return {"payee": payee, "memo": memo, "amount": amount, "addr": addr}

def parse_text_report(lines):
    """
    Manager text export where:
      - Date is at the start of the first line of each record (e.g. '08/07/2025Acme …').
      - If the user uses pipes '|', the first line is 'Contact| Memo| Amount …' and the
        address lines (also pipe-delimited) follow. This path is preferred.
      - If there are no pipes, fall back to heuristics that infer payee/memo/amount/address.
    """
    # find the header row
    hdr = None
    for i, ln in enumerate(lines):
        s = re.sub(r"\s+", "", ln.lower())
        if s.startswith("datecontactdescriptioncredit"):
            hdr = i
            break
    if hdr is None:
        return []

    i, n = hdr + 1, len(lines)
    date_at_start = re.compile(r"^\s*(\d{2}/\d{2}/\d{4})(.*)$")
    amt_at_start  = re.compile(r"^\s*(-?\d{1,3}(?:,\d{3})*(?:\.\d{2}))(.*)$")
    out = []

    while i < n:
        # Find the next line that starts with a date
        m = None
        while i < n and not (m := date_at_start.match(lines[i])):
            if "Printable Checks - For the period" in lines[i] or "custom-report-view" in lines[i]:
                return out
            i += 1
        if i >= n:
            break

        date = m.group(1).strip()
        first_payload = (m.group(2) or "").strip()
        i += 1  # move to the line after the first payload

        # ---- PIPE-AWARE FAST PATH ------------------------------------------
        if "|" in first_payload:
            follow = lines[i:]  # mutable view for the helper to consume
            parsed = parse_payload_with_pipes(first_payload, follow)
            if parsed:
                # compute how many lines were consumed by the helper
                consumed = len(lines[i:]) - len(follow)
                i += consumed
                out.append({
                    "date":  date,
                    "payee": parsed["payee"],
                    "memo":  parsed["memo"],
                    "addr":  parsed["addr"],
                    "amount": parsed["amount"],
                })
                continue
        # ---- END PIPE-AWARE FAST PATH --------------------------------------

        # ---- Legacy multi-line fallback ------------------------------------
        pre = []
        # First fragment might include a memo-like date text; treat as part of 'pre'
        if first_payload:
            pre.append(first_payload)

        # Accumulate lines until we hit the amount line
        while i < n:
            s = lines[i].rstrip()
            if not s:
                i += 1; continue
            if amt_at_start.match(s): break
            if "Printable Checks - For the period" in s or "custom-report-view" in s: break
            pre.append(s.strip()); i += 1

        if i >= n:
            break

        mamt = amt_at_start.match(lines[i].strip())
        if not mamt:
            break
        amt_str, amt_tail = mamt.group(1), mamt.group(2)
        i += 1
        try:
            amount = Decimal(amt_str.replace(",", ""))
        except InvalidOperation:
            break

        addr_lines = []
        if amt_tail.strip(): addr_lines.append(amt_tail.strip())
        while i < n:
            s = lines[i].strip()
            if not s:
                i += 1; break
            if date_at_start.match(s) or amt_at_start.match(s) \
               or "Printable Checks - For the period" in s or "custom-report-view" in s:
                break
            addr_lines.append(s); i += 1

        # Heuristic split of payee vs memo
        if len(pre) >= 4:
            payee = " ".join(pre[:2]); memo = " ".join(pre[2:])
        elif len(pre) == 3:
            payee = pre[0];            memo = " ".join(pre[1:])
        elif len(pre) == 2:
            payee = pre[0];            memo = pre[1]
        elif len(pre) == 1:
            payee = pre[0];            memo = ""
        else:
            payee = "";                memo = ""

        addr = normalize_addr(addr_lines)
        out.append({"date": date, "payee": payee, "memo": memo, "addr": addr, "amount": amount})
        # ---- End legacy fallback -------------------------------------------

    return out

def parse_blocks(lines):
    """
    Fallback parser if header not found; looks for DATE lines and groups until amount.
    """
    i = 0
    while i < len(lines) and not DATE_RE.match(lines[i]): i += 1
    blocks = []
    while i < len(lines):
        if not DATE_RE.match(lines[i]): i += 1; continue
        date = lines[i].strip(); i += 1
        chunk = []
        while i < len(lines) and not DATE_RE.match(lines[i]):
            if lines[i].strip(): chunk.append(lines[i].rstrip())
            i += 1
        if not chunk: continue
        amt_idx = None
        for idx in range(len(chunk)-1, -1, -1):
            if AMOUNT_RE.match(chunk[idx]): amt_idx = idx; break
        if amt_idx is None: continue
        try:
            amount = Decimal(chunk[amt_idx].replace(",", "").strip())
        except InvalidOperation:
            continue
        addr_raw = [ln for ln in chunk[amt_idx+1:] if "|" in ln]
        address = normalize_addr(addr_raw)
        body = chunk[:amt_idx]
        payee = body[:MAX_PAYEE_LINES]; desc = body[len(payee):]
        payee_txt = " ".join([x.strip() for x in payee if x.strip()])
        desc_txt  = " ".join([x.strip() for x in desc  if x.strip()])
        blocks.append({"date": date, "payee": payee_txt, "memo": desc_txt,
                       "amount": amount, "addr": address})
    return blocks

# ------------------------------- Rendering -----------------------------------
def render_check_page(rec):
    page = blank_page()
    payee = sanitize(rec["payee"])
    memo  = sanitize(rec["memo"])
    addr  = [sanitize(a) for a in rec["addr"]]

    put(page, ROW_DATE,  COL_DATE,  rec["date"])
    put_right(page, ROW_AMT_NUM, AMT_RIGHT_COL, f"{rec['amount']:.2f}")
    put(page, ROW_PAYEE, COL_PAYEE, payee)
    put(page, ROW_WORDS, COL_WORDS, money_words(rec["amount"]))

    r = ROW_ADDR_START
    for ln in addr[:4]:
        put(page, r, COL_ADDR, ln); r += 1

    put(page, ROW_MEMO, COL_MEMO, memo)

    # Stub 1
    put(page, ROW_STUB1,   COL_STUB_LABEL, "Payee:");  put(page, ROW_STUB1,   COL_STUB_VAL,   payee)
    put(page, ROW_STUB1+1, COL_STUB_LABEL, "Date:");   put(page, ROW_STUB1+1, COL_STUB_VAL,   rec["date"])
    put(page, ROW_STUB1+2, COL_STUB_LABEL, "Amount:"); put(page, ROW_STUB1+2, COL_STUB_VAL,   f"${rec['amount']:.2f}")
    put(page, ROW_STUB1+3, COL_STUB_LABEL, "Memo:");   put(page, ROW_STUB1+3, COL_STUB_VAL,   memo)

    # Stub 2
    put(page, ROW_STUB2,   COL_STUB_LABEL, "Payee:");  put(page, ROW_STUB2,   COL_STUB_VAL,   payee)
    put(page, ROW_STUB2+1, COL_STUB_LABEL, "Date:");   put(page, ROW_STUB2+1, COL_STUB_VAL,   rec["date"])
    put(page, ROW_STUB2+2, COL_STUB_LABEL, "Amount:"); put(page, ROW_STUB2+2, COL_STUB_VAL,   f"${rec['amount']:.2f}")
    put(page, ROW_STUB2+3, COL_STUB_LABEL, "Memo:");   put(page, ROW_STUB2+3, COL_STUB_VAL,   memo)

    page = [ln[:CPL].ljust(CPL) for ln in page]
    return page

def write_text(path, pages):
    with open(path, "w", encoding="utf-8") as f:
        for p, page in enumerate(pages):
            while page and not page[-1].strip():
                page.pop()
            for ln in page:
                f.write(ln + "\n")
            if p != len(pages) - 1:
                f.write("\f")

def send_raw_to_printer(printer_name, text_data, encoding, charset_num):
    try:
        import win32print
    except Exception:
        raise SystemExit("pywin32 is required for printing (pip install pywin32)")

    # ESC t n (character table)
    charset_cmd = ESC + "t" + bytes([charset_num]).decode("latin-1")
    payload = INIT + CPI10 + LPI6 + charset_cmd + text_data

    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(hPrinter, 1, ("Checks", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, payload.encode(encoding, errors="replace"))
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

def build_calibration_page():
    header = "".join(str((i % 10)) for i in range(1, CPL + 1))
    lines = [header]
    for r in range(2, LINES_PER_PAGE + 1):
        ln = f"{r:02d}" + header[2:]
        lines.append(ln[:CPL])
    return CRLF.join(lines).rstrip() + FF

# --------------------------------- CLI ---------------------------------------
def derive_paths(input_path):
    """
    From 'C:\\in\\finance\\companyname_checks.txt' produce:
      print -> 'C:\\in\\finance\\companyname_checks_print.txt'
      debug -> 'C:\\in\\finance\\companyname_checks_debug.txt'
      printed name -> 'C:\\in\\finance\\companyname_checks_printed_YYYYMMDD_HH_MM.txt'
    """
    d, base = os.path.split(input_path)
    stem, _ = os.path.splitext(base)
    out_print = os.path.join(d or ".", f"{stem}_print.txt")
    out_debug = os.path.join(d or ".", f"{stem}_debug.txt")
    ts = datetime.now().strftime("%Y%m%d_%H_%M")
    renamed = os.path.join(d or ".", f"{stem}_printed_{ts}.txt")
    return out_print, out_debug, renamed

def main():
    ap = argparse.ArgumentParser(description="Render & print Manager.io 'Printable Checks' text to Generic/Text printers.")
    ap.add_argument("positional_input", nargs="?", help="Positional alternative to --input")
    ap.add_argument("--input", default="print_voucher_checks.txt", help="Input text file from Manager.io (default: %(default)s)")
    ap.add_argument("--printer", default="EPSON XP-7100 (GENERIC TEXT)", help="Windows printer name (default: %(default)s)")
    ap.add_argument("--no-print", action="store_true", help="Skip printing; only write preview text.")
    ap.add_argument("--rename", action="store_true", help="Rename input to *_printed_YYYYMMDD_HH_MM.txt on successful print.")
    ap.add_argument("--encoding", default="cp437", help="RAW printer encoding/code page (default: %(default)s)")
    ap.add_argument("--charset", type=int, default=437, help="ESC/P character table: 437, 850, 860, 863, 865 (default: %(default)s)")
    ap.add_argument("--cal", action="store_true", help="Print calibration grid (no input/output) and exit.")
    args = ap.parse_args()

    charset_map = {437: 0, 850: 2, 860: 3, 863: 4, 865: 5}

    # Calibration mode: print and exit (no files)
    if args.cal:
        try:
            send_raw_to_printer(
                args.printer,
                build_calibration_page(),
                args.encoding,
                charset_map.get(args.charset, 0)
            )
            print("Calibration page sent.")
        except Exception as e:
            print(f"Calibration error: {e}")
        return

    # Prefer --input, else positional, else default
    if args.positional_input and (not args.input or args.input == ap.get_default("input")):
        args.input = args.positional_input

    out_print, out_debug, renamed = derive_paths(args.input)

    if not os.path.isfile(args.input):
        print(f"Input not found: {args.input}")
        return

    # Read & parse
    try:
        raw_lines = read_text(args.input)
    except Exception as e:
        print(f"Failed to read input text: {e}")
        return

    records = parse_text_report(raw_lines) or parse_blocks(raw_lines)
    if not records:
        try:
            with open(out_debug, "w", encoding="utf-8") as df:
                for i, ln in enumerate(raw_lines[:400]):
                    df.write(f"{i:04d}: {ln!r}\n")
            print(f"No checks found. Wrote a debug dump to: {out_debug}")
        except Exception as e:
            print(f"No checks found and failed to write debug: {e}")
        return

    # Render all pages
    pages = [trim_trailing_blank_lines(render_check_page(rec)) for rec in records]
    page_strs = [CRLF.join(p).rstrip() for p in pages]
    text_blob = FF.join(page_strs) + FF  # page breaks; printer init added in send_raw_to_printer

    # Save preview (overwrite)
    try:
        write_text(out_print, pages)
        print(f"Wrote preview: {out_print}")
    except Exception as e:
        print(f"Failed to write preview: {e}")

    # Print?
    printed_ok = False
    if not args.no_print:
        try:
            send_raw_to_printer(args.printer, text_blob, args.encoding, charset_map.get(args.charset, 0))
            printed_ok = True
            print("Printed successfully.")
        except Exception as e:
            print(f"Printer error: {e}")
    else:
        print("Skipped printing (--no-print).")

    # Rename input if requested AND printed
    if args.rename and printed_ok:
        try:
            target = renamed
            n = 1
            while os.path.exists(target):
                base, ext = os.path.splitext(renamed)
                target = f"{base}_{n}{ext}"
                n += 1
            os.replace(args.input, target)
            print(f"Renamed input to: {target}")
        except Exception as e:
            print(f"Warning: could not rename input file: {e}")

if __name__ == "__main__":
    main()
