# Contact Manager Pro

An all-inclusive, standalone, self-healing desktop application for deduplicating, intelligently cleaning, and formatting large-scale `.vcf` and `.csv` contact archives into 100% Google Contacts-compatible exports.

Built with Python and PySide6 (Qt) featuring a modern fintech dark mode GUI.

---

## Key Highlights

- **Single `.pyw` File**: Fully standalone with zero external script dependencies. Can be double-clicked or launched standalone anywhere on Windows.
- **Fully Self-Healing**:
  - Automatically verifies and installs missing runtime dependencies (`PySide6`, `openpyxl`) via silent pip bootstrap.
  - Automatically creates and manages `imports/` and `exports/` folder structures.
  - Recovers from encoding inconsistencies (`utf-8-sig`, `utf-8`, `cp1252`, `latin-1` with unicode error fallback).
  - Unhandled exception hook automatically logs full stack traces to `crash_log.txt` with native dialog alert.
- **Strict Phone-Centric Architecture**:
  - 1 unique contact record per valid telephone number.
  - Classifies Indian mobile (10-digits starting with 6, 7, 8, 9) and landlines.
  - Automatically formats numbers to standard international representation (`+91 XXXXX XXXXX`).
- **Intelligent Name Cleaner**:
  - Eliminates repetitive duplicated words (`Vyapar Vyapar` &rarr; `Vyapar`).
  - Removes trailing counters, numbers, and garbage suffixes (`Suraj Sharma 1` &rarr; `Suraj Sharma`).
  - Cleans mangled character repetitions (`Surajj` &rarr; `Suraj`, `Adityaaa` &rarr; `Aditya`).
  - Auto-identifies **Prefixes** (DR, MR, ADV, CA, VIP, etc.) and **Suffixes** (LLP, PVT LTD, KOTAK, BALAJI, ASALFA, etc.) in **ALL CAPS**.
  - Formats personal names in proper **Title Case** with inner corporate/bank tags preserved in ALL CAPS.
  - Learns dataset-specific acronyms dynamically using whole-corpus token frequency analysis.
- **Chronological Newest-Name Selection**:
  - Compares vCard `REV` timestamps, filename dates (`2019-2026`), and filesystem metadata.
  - Prioritizes newer, higher-quality names while preserving older names in Google Contact Notes.
- **Comprehensive Multi-Format Exports**:
  - **Google Contacts CSV**: Chunked to &le;3,000 records per file with 31 official Google column headers.
  - **Google vCard 3.0**: RFC-compliant vCard stream with `FN`, `N`, `TEL`, and `NOTE`.
  - **Excel XLSX**: Styled headers, auto-adjusted column widths, freeze panes, and filters.
  - **Clean CSV**: Streamlined 3-column format (`Name`, `Number`, `Notes`).
  - **Duplicates Audit CSV**: Detailed audit trail of all telephone numbers with multiple merged historical names.
  - **⚡ One-Click Export All**: Exports all 5 formats with a single click directly into `exports/`.

---

## Directory Organization

```
contacts/
├── contact_manager.pyw      # Single all-inclusive standalone application
├── test_suite.py            # Automated verification test suite
├── requirements.txt         # Package dependencies (PySide6, openpyxl)
├── run.bat                  # One-click Windows runner
├── README.md                # Project documentation
├── .gitignore               # Security & data leak protection
├── sample_data/             # Anonymized sample contacts for quick tests
│   ├── sample_contacts.vcf
│   └── sample_contacts.csv
├── imports/                 # Reorganized source .vcf & .csv files
└── exports/                 # Generated Google CSVs, VCF 3.0, and Excel reports
```

---

## Getting Started

### Prerequisites

- Python 3.10+ on Windows, macOS, or Linux.

### Quick Start

1. **Clone the repository**:
   ```bash
   git clone https://github.com/balajifinancemachine-lab/contact-manager.git
   cd contact-manager
   ```

2. **Run the Application**:
   - On Windows: Double-click `contact_manager.pyw` or run `run.bat`.
   - Via Terminal:
     ```bash
     python contact_manager.pyw
     ```
   *(If PySide6 is missing, the self-healing bootstrap will automatically detect and install it).*

3. **Scan Contacts**:
   - Place your raw `.vcf` or `.csv` files into `imports/` (or choose any folder via "Browse Folder").
   - Click **⚡ START DEEP INTELLIGENT SCAN**.
   - Review contacts in the live table and inspector drawer.
   - Click any export button or **⚡ Export All Formats to exports/**.

---

## Running Verification Tests

Execute the automated test suite to verify phone normalizers, deduplication, name cleaning, and assertions:

```bash
python test_suite.py
```

---

## License

MIT License. Designed and maintained for production reliability.
