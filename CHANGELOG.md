# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
### Added
- (placeholder) Add new items here.

### Changed
- (placeholder)

### Fixed
- (placeholder)

---

## [1.0.0] - 2025-10-16
### Added
- Initial release of **Manager.io → Voucher Check Printer**.
- Renders Manager.io **Printable Checks** to QuickBooks-style voucher stock (check on top + two stubs).
- Fixed 80×54 grid (10 CPI, 6 LPI) aligned for Redi-Seal #24539 window envelopes.
- RAW ESC/P printing to **Generic/Text** printers via `pywin32`, plus preview text output.
- Safe ASCII output with CP437 default; switchable to CP850 (`--charset 850 --encoding cp850`).
- Heuristic parser that:
  - Detects inline amount, phone numbers, and embedded dates.
  - Splits supplier address into lines; falls back to city/state/ZIP split when no delimiters are present.
- **`|` pipe delimiter support** for perfect parsing:
  - Use a trailing `|` after **Contact (Supplier Name)** and **Description (Memo)**.
  - Use `|` between **and at the end** of each **Supplier → Address** line.
- Calibration page (`--cal`) that prints an 80×54 ruler for alignment.
- CLI flags: `--input`, `--printer`, `--no-print`, `--rename`, `--encoding`, `--charset`.
- Robust text reader (handles cp1252/utf-8/latin-1/utf-16; cleans NBSP/control chars).
- Money-to-words rendering (“One hundred twenty-three dollars and 45/100”).
- Documentation:
  - README with setup, usage, and Manager.io report instructions.
  - In-script comments describing report fields and delimiter guidance.

### Known limitations
- Designed and tuned for Windows + Generic/Text (ESC/P) printers.
- Best results when using the `|` delimiters in Manager.io outputs.

---

[Unreleased]: https://github.com/pmantos/managerio-check-print/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/pmantos/managerio-check-print/releases/tag/v1.0.0
