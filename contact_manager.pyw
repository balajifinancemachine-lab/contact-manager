"""
=============================================================================
VCF & CSV Intelligent Contact Manager Pro - All-Inclusive Standalone Edition
=============================================================================
Features:
- 100% Self-Contained, Standalone & Self-Healing:
  * Automatic dependency detection and silent auto-installation (PySide6, openpyxl).
  * Auto-initialization of imports/ and exports/ directories.
  * Resilient multi-encoding reader (UTF-8, UTF-8-SIG, CP1252, Latin-1 with replacement).
  * Safe crash reporting with automated diagnostic crash_log.txt capture.
- Phone-Centric Architecture:
  * Strictly 1 unique phone number per contact record.
  * Classifies Indian mobile (10 digits starting with 6,7,8,9) and landline numbers.
  * Formats numbers cleanly into standard international format (+91 XXXXX XXXXX).
- Deep Intelligent Name Cleaner:
  * Eliminates repetitive words (e.g. 'Vyapar Vyapar' -> 'Vyapar').
  * Eliminates trailing counters and numbers (e.g. 'Suraj Sharma 1' -> 'Suraj Sharma').
  * Eliminates mangled character repetitions (e.g. 'Surajj' -> 'Suraj', 'Adityaaa' -> 'Aditya').
  * Auto-identifies Prefixes (Honorifics / Tags) and Suffixes (Entities / Degrees) in ALL CAPS.
  * Formats core personal names in proper Title Case with inner bank names/acronyms in ALL CAPS.
  * Dynamic corpus learning: analyzes entire dataset to detect user-specific business acronyms.
- Chronological Newest-Name Resolution:
  * Analyzes vCard REV / X-ABDATE metadata, filename timestamps (2019-2026), and mtime.
  * Intelligently selects the newest, highest-quality name while preserving historical names in Notes.
- Google Contacts & Multi-Format Exports:
  * Official Google Contacts CSV export (chunked to <=3,000 contacts for 100% smooth Google import).
  * Valid Google-compatible vCard 3.0 export.
  * Styled Excel (.xlsx) export with auto-column widths, header styling, and auto-filters.
  * Clean 3-Column CSV export (Name, Number, Notes).
  * Comprehensive Number Duplicates Audit Report.
  * ⚡ One-Click "Export All Formats" to exports/ directory.
- Modern Responsive Dark UI:
  * Polished dark mode layout with KPI statistics cards.
  * Live real-time search and filter pills.
  * Splitter view with contact table and live inspector drawer.
  * Inline quick name editing with instant update.
"""

import sys
import os
import re
import csv
import json
import shutil
import subprocess
import traceback
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime

# =============================================================================
# SELF-HEALING DEPENDENCY & ENVIRONMENT BOOTSTRAPPER
# =============================================================================

REQUIRED_PACKAGES = {
    "PySide6": "PySide6>=6.5.0",
    "openpyxl": "openpyxl>=3.1.0"
}

def show_native_message(title, message, is_error=False):
    """Displays a native OS message box without requiring third-party GUI packages."""
    if sys.platform == "win32":
        try:
            import ctypes
            style = 0x10 if is_error else 0x40  # MB_ICONERROR or MB_ICONINFORMATION
            ctypes.windll.user32.MessageBoxW(0, str(message), str(title), style)
            return
        except Exception:
            pass
    print(f"[{'ERROR' if is_error else 'INFO'}] {title}: {message}", file=sys.stderr if is_error else sys.stdout)

def auto_heal_dependencies():
    """
    Self-healing routine: checks for required libraries and auto-installs them
    if missing, ensuring the .pyw file can execute from a fresh clone without manual pip steps.
    """
    missing = []
    for pkg in REQUIRED_PACKAGES:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)

    if not missing:
        return True

    try:
        pip_pkgs = [REQUIRED_PACKAGES[m] for m in missing]
        show_native_message(
            "Contact Manager - Self-Healing",
            f"Required packages missing: {', '.join(missing)}.\nAuto-installing now. Please wait a moment...",
            is_error=False
        )
        cmd = [sys.executable, "-m", "pip", "install", "--upgrade"] + pip_pkgs
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if res.returncode == 0:
            show_native_message(
                "Contact Manager - Self-Healing",
                "Successfully installed required packages! Starting application...",
                is_error=False
            )
            return True
        else:
            show_native_message(
                "Contact Manager - Setup Failed",
                f"Failed to auto-install packages:\n{res.stderr}\nPlease run 'pip install -r requirements.txt'.",
                is_error=True
            )
            return False
    except Exception as e:
        show_native_message(
            "Contact Manager - Error",
            f"Auto-install encounter error: {e}",
            is_error=True
        )
        return False

# Run self-healing check before importing GUI components
auto_heal_dependencies()

from PySide6.QtCore import QThread, Signal, Qt, QSize
from PySide6.QtGui import QFont, QColor, QIcon, QPainter, QBrush, QPen, QLinearGradient
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QCheckBox,
    QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QTabWidget, QTextEdit, QGroupBox, QAbstractItemView, QSplitter, QFrame,
    QScrollArea
)

APP_NAME = "VCF Intelligent Contact Manager Pro"
VERSION = "v9.0.0 Standalone"

# =============================================================================
# SELF-HEALING CRASH LOGGER & DIRECTORY INITIALIZER
# =============================================================================

def self_healing_crash_hook(exc_type, exc_value, exc_tb):
    """Captures unhandled exceptions and logs them safely to crash_log.txt."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return

    tb_lines = traceback.format_exception(exc_type, exc_value, exc_tb)
    err_msg = "".join(tb_lines)
    log_file = Path("crash_log.txt")
    try:
        with open(log_file, "a", encoding="utf-8") as fp:
            fp.write(f"\n{'='*70}\nCRASH REPORT: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{'='*70}\n")
            fp.write(f"OS: {sys.platform} | Python: {sys.version}\n")
            fp.write(err_msg)
    except Exception:
        pass

    show_native_message(
        f"{APP_NAME} - Unexpected Error",
        f"An error occurred. Details have been logged to {log_file.resolve()}.\n\n{exc_type.__name__}: {exc_value}",
        is_error=True
    )

sys.excepthook = self_healing_crash_hook

def ensure_workspace_directories():
    """Creates default imports/ and exports/ directories if they don't exist."""
    imports_dir = Path("imports")
    exports_dir = Path("exports")
    imports_dir.mkdir(exist_ok=True)
    exports_dir.mkdir(exist_ok=True)
    return imports_dir, exports_dir

# =============================================================================
# MASTER VOCABULARIES & CAPITALIZATION DATASETS
# =============================================================================

# Master Bank Names & Financial Institutions (ALL CAPS)
BANK_NAMES = {
    "SBI", "STATE BANK", "HDFC", "ICICI", "KOTAK", "KMBL", "AXIS", "PNB", "BOB", "BOI",
    "IDFC", "CANARA", "UNION", "RBL", "YES", "YES BANK", "INDUSIND", "HSBC", "CITI",
    "CITIBANK", "DBS", "SCB", "STANDARD CHARTERED", "IOB", "UCO", "CUB", "NABARD", "RBI",
    "FEDERAL", "PAYTM", "BANK", "BBPS", "NPCI", "MAHABANK", "BANDHAN"
}

# Master Corporate, Acronym & Tag Vocabulary (ALL CAPS)
MASTER_ALL_CAPS_TAGS = {
    # Business Entities & Corporate Structures
    "LLP", "PVT", "LTD", "PRIVATE", "LIMITED", "INC", "LLC", "CORP", "CORPORATION", "CO", "COMPANY",
    "ENT", "ENTERPRISE", "ENTERPRISES", "BROS", "BROTHERS", "SONS",
    # Specific Identifiers & Tags
    "RTR", "ONO", "PP", "PPP", "HP", "RJ", "RC", "AP", "CP", "VP", "VIP",
    "JMD", "TCS", "CCTV", "BMC", "I2C", "AC", "3A", "DCP", "ACP", "RTO", "GST",
    # Prominent Regional & Corporate Tags
    "BALAJI", "ASALFA", "CHUNABHATTI", "KURLA", "GHATKOPAR", "SAKINAKA",
    "ANDHERI", "BANDRA", "MUMBAI", "DELHI", "PUNE", "THANE", "MALAD", "BORIVALI",
    "KALYAN", "DOMBIVLI", "DADAR", "CHEMBUR", "POWAI", "MULUND", "BHANDUP",
    # Telecom & Tech
    "JIO", "AIRTEL", "VI", "IDEA", "VODAFONE", "BSNL", "MTNL",
    "OLA", "UBER", "SWIGGY", "ZOMATO", "AMAZON", "FLIPKART", "MEESHO",
    # Degrees & Academic Tags
    "FY", "SY", "TY", "BMS", "BCOM", "BSC", "BA", "BBA", "MBA", "MCA", "MSC",
    "BTECH", "MTECH", "BE", "ME", "MBBS", "MD", "MS", "BDS", "BAMS", "BHMS",
    "LLB", "LLM", "CA", "CS", "CWA", "CMA", "CFA", "CFP", "FRM",
    "FYBMS", "SYBMS", "TYBMS", "FYBCOM", "SYBCOM", "TYBCOM", "FYBSC", "SYBSC", "TYBSC"
}

DYNAMIC_ALL_CAPS = set()

# Honorifics & Prefixes
PREFIXES = {
    "DR", "DOC", "DOCTOR", "ER", "ENG", "ENGINEER", "CA", "CS", "CWA", "CMA",
    "ADV", "ADVOCATE", "PROF", "PROFESSOR", "MR", "MRS", "MS", "MISS", "MASTER",
    "COL", "COLONEL", "CAPT", "CAPTAIN", "MAJ", "MAJOR", "LT", "LIEUTENANT",
    "PT", "PANDIT", "SHREE", "SHRI", "SRI", "SMT", "SHRIMATI", "SWAMI", "GURU",
    "JUSTICE", "ARCH", "ARCHITECT", "HON", "HONORABLE",
    "VIP", "NEW", "OLD", "TEMP", "TEST", "PP", "RC", "RJ", "HP", "AP", "OM", "SAI", "RTR", "LLP"
}

# Suffix tags
SUFFIXES = MASTER_ALL_CAPS_TAGS | BANK_NAMES | {
    "STORE", "STORES", "SHOP", "SHOPS", "MART", "BAZAR", "BAZAAR", "MARKET",
    "JEWELLER", "JEWELLERS", "JEWELERS", "JEWELLERY", "JEWELERY", "JWELLARS",
    "MOTOR", "MOTORS", "AUTO", "AUTOMOBILE", "AUTOMOBILES", "GARAGE", "CAR", "CARS",
    "TRAVEL", "TRAVELS", "TOUR", "TOURS", "TRANSPORT", "LOGISTICS", "CARRIER", "CARRIERS",
    "PHARMA", "PHARMACEUTICALS", "MEDICAL", "CHEMIST", "CLINIC", "HOSPITAL", "LAB", "LABS",
    "ACADEMY", "CLASSES", "COACHING", "TUTORIALS", "SCHOOL", "COLLEGE", "INSTITUTE",
    "STUDIO", "SALON", "PARLOUR", "BEAUTY", "SPA", "FITNESS", "GYM",
    "HOTEL", "RESTAURANT", "CAFE", "DHABA", "CATERERS", "CATERING", "SWEETS", "BAKERY",
    "HARDWARE", "ELECTRICALS", "ELECTRONICS", "TELECOM", "COMMUNICATION", "COMMUNICATIONS",
    "MOBILE", "EVENT", "EVENTS", "SERVICE", "SERVICES", "TRADER", "TRADERS", "TRADING",
    "TECH", "SOLUTIONS", "INFOTECH", "SYSTEMS", "NETWORK", "NETWORKS"
}

# Multi-word suffix phrases recognized as a unit
MULTI_WORD_SUFFIXES = [
    ("PVT", "LTD"),
    ("PRIVATE", "LIMITED"),
    ("SONS", "AND", "DAUGHTERS"),
    ("AND", "SONS"),
    ("AND", "BROS"),
    ("AND", "BROTHERS"),
    ("TRAVEL", "AND", "TOURS"),
    ("TOURS", "AND", "TRAVELS"),
    ("FY", "BMS", "RJ"),
    ("SY", "BMS", "RJ"),
    ("TY", "BMS", "RJ"),
    ("FY", "BMS"),
    ("SY", "BMS"),
    ("TY", "BMS"),
    ("FY", "BCOM"),
    ("SY", "BCOM"),
    ("TY", "BCOM"),
    ("FY", "BSC"),
    ("SY", "BSC"),
    ("TY", "BSC"),
    ("MOBILE", "SHOP"),
    ("CAR", "CARE"),
    ("PAN", "PATTI"),
    ("KOTAK", "BANK"),
    ("HDFC", "BANK"),
    ("ICICI", "BANK"),
    ("SBI", "BANK"),
    ("AXIS", "BANK")
]

VALID_DOUBLE_ENDINGS = {
    "app", "all", "ell", "ill", "oll", "ull", "off", "iff", "uff",
    "ass", "ess", "iss", "oss", "uss", "inn", "ann", "lee", "bee", "see",
    "fee", "too", "zoo", "mann", "shroff", "kinn"
}

COMMON_SHORT_NAMES = {
    "RAM", "RAJ", "JAIN", "SHAH", "ALI", "ROY", "DEV", "PAL", "SEN", "DAS",
    "DEY", "RAO", "GILL", "LAL", "SAI", "OM", "ANUJ", "AJAY", "ALOK", "AMIT",
    "AMAR", "ARUN", "ATUL", "BABU", "BALU", "BHAI", "DAYA", "DIPU", "HARI",
    "JEET", "KAPIL", "KARAN", "MANI", "MANOJ", "MONU", "NEEL", "NEHA", "NILU",
    "NITIN", "PAPU", "PIHU", "PUJA", "RAVI", "RENU", "RINKU", "RISHI", "RITU",
    "RIYA", "ROHAN", "ROHIT", "RUBY", "RUPA", "SANJU", "SEEMA", "SHIKHA",
    "SONU", "SUMIT", "SUNIL", "TARA", "TINA", "VARUN", "VIJAY", "VIKAS", "VINOD"
}

PLACEHOLDER_NAMES = {
    "unknown", "unknown contact", "unnamed", "no name", "none", "null",
    "nan", "contact", "person", "new contact", "test", "demo", "fb",
    "facebook", "room", "admin", "owner", "customer", "staff", "office",
    "user", "default", "call", "temp", "temporary", "my contacts", "all contacts"
}

GENERIC_BULK_PATTERNS = [
    r"^(?:ZONE\s*\d+[A-Z]?|\d+\s*ZONE\s*\d+[A-Z]?)$",
    r"^(?:JUSTDIAL|INDIAMART)\s*(?:BUYER|SELLER|USER)?$",
    r"^(?:BUYER|SELLER|CUSTOMER|CLIENT)\b",
    r"^(?:ROOM|FLAT|OFFICE|HOME)\b",
    r"^(?:UNKNOWN|UNNAMED|NO\s*NAME|TEMP|TEST)\b",
    r"^(?:RC\s*RJ|SY\s*BMS\s*RJ|FY\s*BMS\s*RJ|RCBCW|RCRJ|RCC\s*RJ)\s*\d*$"
]

# =============================================================================
# INTELLIGENT ALGORITHMS & HELPERS
# =============================================================================

def learn_corpus_acronyms(raw_names_list):
    """Analyzes the full dataset to dynamically discover dataset-specific acronyms."""
    global DYNAMIC_ALL_CAPS
    token_total = Counter()
    token_upper = Counter()

    for name in raw_names_list:
        words = re.findall(r"[A-Za-z0-9&]+", str(name or ""))
        for w in words:
            u = w.upper()
            if len(u) < 2 or len(u) > 6 or u.isdigit():
                continue
            token_total[u] += 1
            if w.isupper():
                token_upper[u] += 1

    discovered = set()
    for u, total in token_total.items():
        if u in COMMON_SHORT_NAMES or u in PLACEHOLDER_NAMES:
            continue
        has_vowels = any(c in "AEIOU" for c in u)
        if not has_vowels and u.isalpha() and 2 <= len(u) <= 5:
            discovered.add(u)
            continue
        upper_c = token_upper.get(u, 0)
        if total >= 2 and (upper_c / total) >= 0.70:
            discovered.add(u)

    DYNAMIC_ALL_CAPS.update(discovered)
    return DYNAMIC_ALL_CAPS

def should_be_all_caps(word, is_prefix=False, is_suffix=False):
    """Checks whether a token should be formatted in ALL CAPS."""
    u = word.upper()
    if u in COMMON_SHORT_NAMES:
        return False
    if u in BANK_NAMES:
        return True
    if u in MASTER_ALL_CAPS_TAGS or u in DYNAMIC_ALL_CAPS:
        return True
    has_vowels = any(c in "AEIOU" for c in u)
    if not has_vowels and u.isalpha() and 2 <= len(u) <= 5:
        return True
    if is_prefix and u in PREFIXES:
        return True
    if is_suffix and u in SUFFIXES:
        return True
    return False

def digits_only(value):
    val = re.sub(r"(?i)^tel:", "", str(value or "").strip())
    val = re.sub(r"(?i)(?:ext|extension|x)\s*[=:.-]?\s*\d+\s*$", "", val)
    return re.sub(r"\D", "", val)

def has_too_many_consecutive_zeros(digits):
    return bool(re.search(r"0{5,}", digits))

def classify_indian_number(value):
    """Strict Indian telephone normalizer (10 digits)."""
    d = digits_only(value)
    if not d or len(d) < 7:
        return None, None
    if has_too_many_consecutive_zeros(d):
        return None, None

    if d.startswith("0091") and len(d) >= 12:
        d = d[4:]
    elif d.startswith("91") and len(d) in (12, 13):
        d = d[2:]
    elif d.startswith("091") and len(d) in (12, 13):
        d = d[3:]

    if len(d) == 11 and d.startswith("0"):
        d = d[1:]

    if len(d) != 10:
        return None, None

    if d[0] in "6789":
        return d, "mobile"
    if d[0] in "2345":
        return d, "landline"

    return None, None

def format_indian_number(number):
    if not number:
        return ""
    d = digits_only(number)
    if len(d) == 10:
        if d[0] in "6789":
            return f"+91 {d[:5]} {d[5:]}"
        return f"+91 {d[:2]} {d[2:6]} {d[6:]}"
    return number

def remove_character_repetition(word):
    if not word or len(word) <= 2:
        return word
    word = re.sub(r"([^\W\d_])\1{2,}", r"\1", word, flags=re.UNICODE)
    if word.lower() in VALID_DOUBLE_ENDINGS:
        return word
    word = re.sub(r"(?i)([jydhwvxgq])\1+$", r"\1", word)
    word = re.sub(r"(?i)([b-df-hj-np-tv-z])aa$", r"\1a", word)
    word = re.sub(r"(?i)iia$", r"ia", word)
    return word

def strip_embedded_numbers(text):
    raw = str(text or "")
    raw = re.sub(r"(?i)^\+?91\d{10}\s*", " ", raw)
    raw = re.sub(r"^\d{10}\s*", " ", raw)
    raw = re.sub(r"\b\d{5,}\b", " ", raw)
    tag_pattern = r"(?:HP|PP|PPP|RTR|LLP|BALAJI|RJ|RC|JMD|KOTAK|BMS|ASALFA|ONO)"
    raw = re.sub(r"\b\d{1,2}\b(?=\s+(?:" + tag_pattern + r")\b)", " ", raw, flags=re.IGNORECASE)
    raw = re.sub(r"[\s\-_(]+(\d{1,3})[)]?\s*$", " ", raw)
    return re.sub(r"\s+", " ", raw).strip()

def deduplicate_word_tokens(tokens):
    cleaned = []
    seen = set()
    for tok in tokens:
        normalized_tok = tok.casefold()
        if not normalized_tok:
            continue
        if cleaned and cleaned[-1].casefold() == normalized_tok:
            continue
        if normalized_tok in seen and len(normalized_tok) > 2:
            continue
        cleaned.append(tok)
        seen.add(normalized_tok)
    return cleaned

def is_prefix_word(word):
    u = word.upper().rstrip(".:-")
    return u in PREFIXES

def is_suffix_word(word):
    u = word.upper().rstrip(".:-")
    return should_be_all_caps(u, is_suffix=True)

def smart_format_token(word, is_prefix=False, is_suffix=False):
    if not word:
        return ""
    if should_be_all_caps(word, is_prefix=is_prefix, is_suffix=is_suffix):
        return word.upper()
    if len(word) == 1 and word.isalpha():
        return word.upper()
    return word[0].upper() + word[1:].lower()

def parse_and_format_name(name_string):
    raw = str(name_string or "").strip()
    if not raw:
        return {"final_name": "Contact", "prefix": "", "core_name": "Contact", "suffix": "", "original": ""}

    s = strip_embedded_numbers(raw)
    words = re.findall(r"[^\W_]+|&", s, flags=re.UNICODE)
    if not words:
        fallback = raw if any(c.isalnum() for c in raw) else "Contact"
        return {"final_name": fallback, "prefix": "", "core_name": fallback, "suffix": "", "original": raw}

    cleaned_words = [remove_character_repetition(w) for w in words]
    cleaned_words = [w for w in cleaned_words if w.strip()]
    tokens = deduplicate_word_tokens(cleaned_words)
    if not tokens:
        return {"final_name": raw or "Contact", "prefix": "", "core_name": raw or "Contact", "suffix": "", "original": raw}

    prefix_tokens = []
    suffix_tokens = []
    remaining = list(tokens)

    n = len(remaining)
    matched_suffix_len = 0
    for mw in MULTI_WORD_SUFFIXES:
        mw_len = len(mw)
        if n >= mw_len:
            window = [remaining[n - mw_len + i].upper() for i in range(mw_len)]
            if window == list(mw):
                matched_suffix_len = mw_len
                break

    if matched_suffix_len > 0:
        suffix_tokens = [t.upper() for t in remaining[n - matched_suffix_len:]]
        remaining = remaining[:n - matched_suffix_len]

    while remaining and is_suffix_word(remaining[-1]) and len(remaining) > 1:
        suffix_tokens.insert(0, remaining.pop().upper())

    while remaining and is_prefix_word(remaining[0]) and len(remaining) > 1:
        prefix_tokens.append(remaining.pop(0).upper())

    if not remaining:
        if suffix_tokens:
            remaining.append(suffix_tokens.pop(0))
        elif prefix_tokens:
            remaining.append(prefix_tokens.pop())

    core_formatted = [smart_format_token(w) for w in remaining]

    prefix_str = " ".join(prefix_tokens)
    core_str = " ".join(core_formatted)
    suffix_str = " ".join(suffix_tokens)

    final_parts = []
    if prefix_str:
        final_parts.append(prefix_str)
    if core_str:
        final_parts.append(core_str)
    if suffix_str:
        final_parts.append(suffix_str)

    final_name = " ".join(final_parts).strip()
    if not final_name:
        final_name = raw or "Contact"

    return {
        "final_name": final_name,
        "prefix": prefix_str,
        "core_name": core_str,
        "suffix": suffix_str,
        "original": raw
    }

def is_valid_name(name):
    s = str(name or "").strip()
    if not s:
        return False
    alpha = sum(c.isalpha() for c in s)
    if alpha < 2:
        return False
    norm = re.sub(r"[^\w\s]", " ", s.casefold(), flags=re.UNICODE)
    norm = re.sub(r"\s+", " ", norm).strip()
    if norm in PLACEHOLDER_NAMES:
        return False
    return True

def extract_date_from_string(text):
    if not text:
        return None
    m = re.search(r"(\d{4})(\d{2})(\d{2})_?(\d{2})(\d{2})(\d{2})", text)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6)))
        except ValueError:
            pass
    m = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", text)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.search(r"20\d{2}[01]\d[0-3]\d", text)
    if m:
        s = m.group(0)
        try:
            return datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            pass
    return None

def is_generic_bulk_label(text):
    clean = re.sub(r"[\s\-_(]+\d+[)]?\s*$", "", str(text or "").strip(), flags=re.I)
    clean = re.sub(r"^\d+\s*", "", clean).strip().upper()
    for pat in GENERIC_BULK_PATTERNS:
        if re.search(pat, clean):
            return True
    return False

def calculate_name_quality(raw_name, is_newest=False):
    if not is_valid_name(raw_name):
        return -999.0

    if is_generic_bulk_label(raw_name):
        return -50.0

    parsed = parse_and_format_name(raw_name)
    core = parsed["core_name"]
    words = core.split()
    score = 0.0

    score += min(len(words), 4) * 15.0
    score += min(len(core), 30) * 1.5

    digits = sum(c.isdigit() for c in core)
    score -= digits * 20.0

    if len(core) <= 2:
        score -= 25.0

    if parsed["prefix"]:
        score += 8.0
    if parsed["suffix"]:
        score += 8.0

    if is_newest:
        score += 35.0

    return score

# =============================================================================
# VCF & CSV PARSERS
# =============================================================================

def unfold_vcf(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    result = []
    for line in lines:
        if line.startswith((" ", "\t")) and result:
            result[-1] += line[1:]
        else:
            result.append(line)
    return result

def split_property(line):
    pos = line.find(":")
    if pos < 0:
        return None, None
    return line[:pos], line[pos + 1:]

def split_semicolon(text):
    out, buf, quoted = [], [], False
    for ch in text:
        if ch == '"':
            quoted = not quoted
            buf.append(ch)
        elif ch == ";" and not quoted:
            out.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    out.append("".join(buf))
    return out

def decode_vcf_value(raw_value, params):
    enc = ",".join(params.get("ENCODING", [])).upper()
    if "QUOTED-PRINTABLE" in enc:
        try:
            import quopri
            return quopri.decodestring(raw_value).decode("utf-8", errors="replace").strip()
        except Exception:
            pass
    val = raw_value or ""
    val = val.replace("\\r", "\r").replace("\\n", "\n").replace("\\N", "\n")
    val = val.replace("\\,", ",").replace("\\;", ";").replace("\\\\", "\\")
    return val.strip()

def parse_property(line):
    left, raw = split_property(line)
    if left is None:
        return None
    parts = split_semicolon(left)
    head = parts[0]
    group = ""
    if "." in head:
        group, head = head.rsplit(".", 1)
    name = head.strip().upper()

    params = defaultdict(list)
    for part in parts[1:]:
        if "=" in part:
            k, v = part.split("=", 1)
            k = k.strip().upper()
            v = v.strip('"')
            for x in v.split(","):
                if x:
                    params[k].append(x.strip())
        elif part:
            params["TYPE"].append(part.strip().upper())

    return {
        "name": name,
        "group": group,
        "params": dict(params),
        "raw_value": raw,
        "value": decode_vcf_value(raw, params)
    }

def parse_vcf_file(path):
    raw = Path(path).read_bytes()
    text = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="replace")

    cards = []
    current = None
    for line in unfold_vcf(text):
        line = line.lstrip("\ufeff").strip()
        u = line.upper()
        if u == "BEGIN:VCARD":
            current = []
        elif u == "END:VCARD":
            if current is not None:
                cards.append(current)
            current = None
        elif current is not None and line:
            p = parse_property(line)
            if p:
                current.append(p)
    return cards

def parse_csv_contacts(path):
    cards = []
    encodings = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
    lines = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc) as fp:
                reader = csv.DictReader(fp)
                rows = list(reader)
                lines = rows
                break
        except Exception:
            continue

    if not lines:
        return []

    for row in lines:
        name = row.get("First Name", "") or row.get("Name", "") or row.get("FN", "")
        last_name = row.get("Last Name", "") or row.get("Family Name", "")
        if last_name and last_name not in name:
            name = f"{name} {last_name}".strip()

        phones = []
        for k, v in row.items():
            if not v or not v.strip():
                continue
            k_upper = k.upper()
            if "PHONE" in k_upper or "TEL" in k_upper or "MOBILE" in k_upper or "NUMBER" in k_upper:
                phones.append(v.strip())

        notes = row.get("Notes", "") or row.get("Note", "")

        card_props = []
        if name:
            card_props.append({"name": "FN", "value": name, "params": {}, "raw_value": name})
        for p in phones:
            card_props.append({"name": "TEL", "value": p, "params": {"TYPE": ["CELL"]}, "raw_value": p})
        if notes:
            card_props.append({"name": "NOTE", "value": notes, "params": {}, "raw_value": notes})

        if card_props:
            cards.append(card_props)

    return cards

def get_source_timestamp(file_path, card_props=None):
    if card_props:
        for p in card_props:
            if p["name"] in ("REV", "X-ABDATE"):
                dt = extract_date_from_string(p["value"])
                if dt:
                    return dt

    dt = extract_date_from_string(Path(file_path).name)
    if dt:
        return dt

    try:
        mtime = os.path.getmtime(file_path)
        return datetime.fromtimestamp(mtime)
    except Exception:
        return datetime(2020, 1, 1)

def extract_phone_records(cards, file_path):
    records = []
    discarded_no_name = 0
    discarded_no_phone = 0
    base_timestamp = get_source_timestamp(file_path)

    for card in cards:
        card_time = get_source_timestamp(file_path, card) or base_timestamp

        candidate_names = []
        for p in card:
            if p["name"] in ("FN", "NICKNAME") and is_valid_name(p["value"]):
                candidate_names.append(p["value"].strip())
            elif p["name"] == "N" and p["value"].strip():
                parts = p["value"].split(";")
                assembled = " ".join(x.strip() for x in parts if x.strip())
                if is_valid_name(assembled):
                    candidate_names.append(assembled)

        card_notes = ""
        for p in card:
            if p["name"] == "NOTE" and p["value"].strip():
                card_notes = p["value"].strip()
                for sub in card_notes.split("|"):
                    sub_clean = sub.replace("Associated names:", "").replace("Original name:", "").strip()
                    if is_valid_name(sub_clean) and sub_clean not in candidate_names:
                        candidate_names.append(sub_clean)

        found_phones = []
        for p in card:
            if p["name"] in ("TEL", "PHONE", "MOBILE", "LANDLINE") or "TEL" in p["name"]:
                norm, kind = classify_indian_number(p["value"])
                if norm and norm not in [x[0] for x in found_phones]:
                    found_phones.append((norm, kind))

        if not found_phones:
            discarded_no_phone += 1
            continue
        if not candidate_names:
            discarded_no_name += 1
            continue

        for number, kind in found_phones:
            for name in candidate_names:
                records.append({
                    "number": number,
                    "kind": kind,
                    "raw_name": name,
                    "source_file": Path(file_path).name,
                    "timestamp": card_time,
                    "notes": card_notes
                })

    return records, len(cards), discarded_no_name, discarded_no_phone

def merge_and_resolve_contacts(records_by_number):
    contacts = []

    for number in sorted(records_by_number):
        entries = records_by_number[number]
        entries.sort(key=lambda e: e["timestamp"], reverse=True)
        newest_date = entries[0]["timestamp"]
        kind = entries[0]["kind"]

        all_raw_names = []
        for e in entries:
            name = e["raw_name"].strip()
            if name and name not in all_raw_names:
                all_raw_names.append(name)

        evaluated = []
        for name in all_raw_names:
            is_from_newest = any(e["raw_name"] == name and e["timestamp"] == newest_date for e in entries)
            q_score = calculate_name_quality(name, is_newest=is_from_newest)
            parsed = parse_and_format_name(name)
            evaluated.append({
                "raw": name,
                "parsed": parsed,
                "score": q_score,
                "is_newest": is_from_newest
            })

        evaluated.sort(key=lambda x: x["score"], reverse=True)
        best_candidate = evaluated[0] if evaluated else {
            "parsed": parse_and_format_name(all_raw_names[0] if all_raw_names else "Contact"),
            "raw": all_raw_names[0] if all_raw_names else "Contact"
        }

        final_parsed = best_candidate["parsed"]
        final_name = final_parsed.get("final_name", "").strip()
        if not final_name:
            final_name = best_candidate.get("raw", "").strip() or "Contact"

        distinct_formatted_names = []
        for item in evaluated:
            fn = item["parsed"]["final_name"]
            if fn and fn not in distinct_formatted_names:
                distinct_formatted_names.append(fn)

        notes = ""
        if len(distinct_formatted_names) > 1:
            notes = "Associated names: " + " | ".join(distinct_formatted_names)

        history = []
        for e in entries:
            history.append({
                "raw_name": e["raw_name"],
                "source_file": e["source_file"],
                "date": e["timestamp"].strftime("%Y-%m-%d"),
                "is_newest": (e["timestamp"] == newest_date)
            })

        contacts.append({
            "number": number,
            "kind": kind,
            "name": final_name,
            "prefix": final_parsed.get("prefix", ""),
            "core_name": final_parsed.get("core_name", "") or final_name,
            "suffix": final_parsed.get("suffix", ""),
            "raw_chosen": best_candidate["raw"],
            "all_raw_names": all_raw_names,
            "notes": notes,
            "source_count": len(entries),
            "newest_date": newest_date.strftime("%Y-%m-%d"),
            "history": history
        })

    return contacts

# =============================================================================
# BACKGROUND SCAN WORKER
# =============================================================================

class ScanWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, folder=None, recursive=True, include_csv=True):
        super().__init__()
        # Self-healing path resolution
        if not folder or folder == ".":
            if Path("imports").exists() and any(Path("imports").glob("*.vcf")):
                self.folder = Path("imports")
            else:
                self.folder = Path(".")
        else:
            self.folder = Path(folder)

        self.recursive = recursive
        self.include_csv = include_csv

    def run(self):
        try:
            vcf_patterns = ["*.vcf", "*.VCF", "*.vcard", "*.VCARD"]
            files = []
            for pat in vcf_patterns:
                files.extend(self.folder.rglob(pat) if self.recursive else self.folder.glob(pat))

            if self.include_csv:
                csv_patterns = ["*.csv", "*.CSV"]
                for pat in csv_patterns:
                    files.extend(self.folder.rglob(pat) if self.recursive else self.folder.glob(pat))

            files = sorted(set(files))
            if not files:
                self.failed.emit(f"No .vcf or .csv files found in directory: {self.folder.resolve()}")
                return

            by_number = defaultdict(list)
            discovered_fields = Counter()
            parse_errors = []

            total_raw_cards = 0
            discarded_no_name = 0
            discarded_no_phone = 0
            total_associations = 0

            for idx, file_path in enumerate(files, 1):
                pct = int(idx * 75 / len(files))
                self.progress.emit(pct, f"Processing {file_path.name} ({idx}/{len(files)})")

                try:
                    if file_path.suffix.lower() == ".csv":
                        cards = parse_csv_contacts(file_path)
                    else:
                        cards = parse_vcf_file(file_path)

                    records, raw_count, no_name, no_phone = extract_phone_records(cards, file_path)

                    total_raw_cards += raw_count
                    discarded_no_name += no_name
                    discarded_no_phone += no_phone
                    total_associations += len(records)

                    for card in cards:
                        for p in card:
                            discovered_fields[p["name"]] += 1

                    for rec in records:
                        by_number[rec["number"]].append(rec)

                except Exception as e:
                    parse_errors.append(f"{file_path.name}: {str(e)}")

            self.progress.emit(82, "Analyzing corpus vocabulary for acronyms & bank tags...")
            all_raw_names = []
            for records_list in by_number.values():
                for r in records_list:
                    if r.get("raw_name"):
                        all_raw_names.append(r["raw_name"])
            learn_corpus_acronyms(all_raw_names)

            self.progress.emit(88, "Merging duplicates & resolving intelligent names...")
            contacts = merge_and_resolve_contacts(by_number)

            self.progress.emit(100, "Scan and intelligence formatting complete.")

            self.finished.emit({
                "files": files,
                "contacts": contacts,
                "raw_cards": total_raw_cards,
                "total_associations": total_associations,
                "discarded_no_name": discarded_no_name,
                "discarded_no_phone": discarded_no_phone,
                "fields": dict(discovered_fields),
                "parse_errors": parse_errors
            })

        except Exception:
            self.failed.emit(traceback.format_exc())

# =============================================================================
# MODERN DARK THEME STYLESHEET
# =============================================================================

MODERN_STYLE = """
QMainWindow, QWidget {
    background-color: #0b0f19;
    color: #f1f5f9;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
    font-size: 13px;
}

QGroupBox {
    background-color: #111827;
    border: 1px solid #1e293b;
    border-radius: 10px;
    margin-top: 12px;
    font-weight: 600;
    padding-top: 14px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 6px;
    color: #38bdf8;
    font-size: 13px;
}

QLineEdit {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 8px 12px;
    color: #ffffff;
    font-size: 13px;
}

QLineEdit:focus {
    border: 1px solid #38bdf8;
    background-color: #0f172a;
}

QPushButton {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 6px;
    padding: 8px 16px;
    color: #f8fafc;
    font-weight: 600;
}

QPushButton:hover {
    background-color: #334155;
    border-color: #475569;
}

QPushButton:pressed {
    background-color: #0f172a;
}

QPushButton#PrimaryButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #0284c7);
    border: none;
    color: #ffffff;
    font-weight: 700;
    font-size: 14px;
    border-radius: 8px;
}

QPushButton#PrimaryButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4ed8, stop:1 #0369a1);
}

QPushButton#ExportButton {
    background-color: #1e293b;
    border: 1px solid #0284c7;
    color: #38bdf8;
    font-weight: 600;
    border-radius: 6px;
    padding: 8px 14px;
}

QPushButton#ExportButton:hover {
    background-color: #0284c7;
    color: #ffffff;
}

QPushButton#ExportAllButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10b981);
    border: none;
    color: #ffffff;
    font-weight: 700;
    border-radius: 6px;
    padding: 8px 16px;
}

QPushButton#ExportAllButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #047857, stop:1 #059669);
}

QPushButton#FilterPill {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 14px;
    padding: 4px 14px;
    color: #94a3b8;
    font-size: 12px;
}

QPushButton#FilterPill:checked, QPushButton#FilterPill:hover {
    background-color: #0284c7;
    border-color: #38bdf8;
    color: #ffffff;
    font-weight: 600;
}

QTableWidget {
    background-color: #0f172a;
    border: 1px solid #1e293b;
    border-radius: 8px;
    gridline-color: #1e293b;
    color: #e2e8f0;
    selection-background-color: #1e3a8a;
    selection-color: #ffffff;
}

QHeaderView::section {
    background-color: #111827;
    color: #94a3b8;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #334155;
    font-weight: 600;
    font-size: 12px;
}

QTabWidget::pane {
    border: 1px solid #1e293b;
    background-color: #0f172a;
    border-radius: 8px;
    top: -1px;
}

QTabBar::tab {
    background-color: #111827;
    color: #94a3b8;
    padding: 10px 20px;
    margin-right: 4px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: 600;
}

QTabBar::tab:selected {
    background-color: #0f172a;
    color: #38bdf8;
    border-bottom: 2px solid #38bdf8;
}

QProgressBar {
    border: 1px solid #1e293b;
    border-radius: 5px;
    text-align: center;
    background-color: #111827;
    color: #ffffff;
    font-weight: 600;
    height: 18px;
}

QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #38bdf8);
    border-radius: 4px;
}

QScrollBar:vertical {
    border: none;
    background: #0b0f19;
    width: 8px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: #334155;
    min-height: 20px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #475569;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QSplitter::handle {
    background-color: #1e293b;
}
"""

class StatCard(QFrame):
    def __init__(self, title, value="0", color="#38bdf8", parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QFrame {
                background-color: #111827;
                border: 1px solid #1e293b;
                border-radius: 10px;
                padding: 10px;
            }
        """)
        v = QVBoxLayout(self)
        v.setContentsMargins(10, 8, 10, 8)
        v.setSpacing(2)

        self.title_lbl = QLabel(title.upper())
        self.title_lbl.setStyleSheet("color: #94a3b8; font-size: 11px; font-weight: 700; letter-spacing: 0.5px;")
        v.addWidget(self.title_lbl)

        self.value_lbl = QLabel(str(value))
        self.value_lbl.setStyleSheet(f"color: {color}; font-size: 20px; font-weight: 800;")
        v.addWidget(self.value_lbl)

    def set_value(self, val):
        self.value_lbl.setText(str(val))

# =============================================================================
# MAIN APPLICATION WINDOW
# =============================================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {VERSION}")
        self.resize(1620, 940)
        self.setStyleSheet(MODERN_STYLE)

        self.result = None
        self.worker = None
        self.current_filtered_contacts = []
        self.active_filter = "all"

        ensure_workspace_directories()
        self.build_ui()

    def build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(16, 14, 16, 14)
        main_layout.setSpacing(12)

        # 1. Header
        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel(APP_NAME)
        title.setStyleSheet("font-size: 22px; font-weight: 800; color: #f8fafc;")
        subtitle = QLabel("Deep Audit Engine • Google Contacts Ready • Number-Unique Anchor • All-Inclusive Standalone")
        subtitle.setStyleSheet("font-size: 12px; color: #38bdf8; font-weight: 500;")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box)
        header.addStretch()

        ver_badge = QLabel(VERSION)
        ver_badge.setStyleSheet("background-color: #1e293b; color: #38bdf8; padding: 4px 10px; border-radius: 12px; font-weight: 700; font-size: 11px;")
        header.addWidget(ver_badge)
        main_layout.addLayout(header)

        # 2. Controls Group
        controls_group = QGroupBox("SCAN SOURCE & DIRECTORY")
        cg_layout = QGridLayout(controls_group)
        cg_layout.setContentsMargins(14, 14, 14, 14)
        cg_layout.setHorizontalSpacing(10)
        cg_layout.setVerticalSpacing(8)

        cg_layout.addWidget(QLabel("Folder Path:"), 0, 0)
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("Select directory containing .vcf or Google export .csv files...")

        # Self-healing default folder: prefers 'imports' if present
        default_dir = Path("imports") if Path("imports").exists() and any(Path("imports").glob("*.vcf")) else Path.cwd()
        self.folder_input.setText(str(default_dir.resolve()))
        cg_layout.addWidget(self.folder_input, 0, 1, 1, 3)

        browse_btn = QPushButton("Browse Folder")
        browse_btn.clicked.connect(self.choose_folder)
        cg_layout.addWidget(browse_btn, 0, 4)

        self.recursive_chk = QCheckBox("Scan subfolders recursively")
        self.recursive_chk.setChecked(True)
        cg_layout.addWidget(self.recursive_chk, 1, 0, 1, 2)

        self.csv_chk = QCheckBox("Include Google Contacts CSV exports")
        self.csv_chk.setChecked(True)
        cg_layout.addWidget(self.csv_chk, 1, 2)

        self.scan_btn = QPushButton("⚡ START DEEP INTELLIGENT SCAN")
        self.scan_btn.setObjectName("PrimaryButton")
        self.scan_btn.setMinimumHeight(40)
        self.scan_btn.clicked.connect(self.start_scan)
        cg_layout.addWidget(self.scan_btn, 0, 5, 2, 1)

        main_layout.addWidget(controls_group)

        # 3. KPI Dashboard
        stats_grid = QHBoxLayout()
        stats_grid.setSpacing(10)

        self.card_files = StatCard("Files Scanned", "0", "#38bdf8")
        self.card_raw = StatCard("Raw Cards Read", "0", "#94a3b8")
        self.card_unique = StatCard("Unique Numbers", "0", "#10b981")
        self.card_merged = StatCard("Merged Duplicates", "0", "#f59e0b")
        self.card_formatted = StatCard("Names Cleaned", "0", "#8b5cf6")
        self.card_tags = StatCard("Prefix/Suffix Tagged", "0", "#ec4899")

        for c in (self.card_files, self.card_raw, self.card_unique, self.card_merged, self.card_formatted, self.card_tags):
            stats_grid.addWidget(c)

        main_layout.addLayout(stats_grid)

        # 4. Progress Bar
        prog_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        prog_layout.addWidget(self.progress_bar, 1)

        self.status_lbl = QLabel("Ready to scan.")
        self.status_lbl.setStyleSheet("color: #94a3b8; font-weight: 500;")
        prog_layout.addWidget(self.status_lbl)
        main_layout.addLayout(prog_layout)

        # 5. Tabbed Interface
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs, 1)

        # Tab 1: Final Contacts (Google Ready)
        self.tab_contacts = QWidget()
        tab_v = QVBoxLayout(self.tab_contacts)
        tab_v.setContentsMargins(10, 10, 10, 10)
        tab_v.setSpacing(10)

        filter_bar = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Search contacts by Name, Phone, Prefix, Suffix, or Notes...")
        self.search_input.textChanged.connect(self.apply_filters_and_refresh)
        filter_bar.addWidget(self.search_input, 2)

        self.pill_all = QPushButton("All Contacts")
        self.pill_all.setObjectName("FilterPill")
        self.pill_all.setCheckable(True)
        self.pill_all.setChecked(True)
        self.pill_all.clicked.connect(lambda: self.set_active_filter("all"))
        filter_bar.addWidget(self.pill_all)

        self.pill_tags = QPushButton("Prefix / Suffix Active")
        self.pill_tags.setObjectName("FilterPill")
        self.pill_tags.setCheckable(True)
        self.pill_tags.clicked.connect(lambda: self.set_active_filter("tags"))
        filter_bar.addWidget(self.pill_tags)

        self.pill_dups = QPushButton("Merged Number Duplicates")
        self.pill_dups.setObjectName("FilterPill")
        self.pill_dups.setCheckable(True)
        self.pill_dups.clicked.connect(lambda: self.set_active_filter("duplicates"))
        filter_bar.addWidget(self.pill_dups)

        self.pill_cleaned = QPushButton("Names Cleaned")
        self.pill_cleaned.setObjectName("FilterPill")
        self.pill_cleaned.setCheckable(True)
        self.pill_cleaned.clicked.connect(lambda: self.set_active_filter("cleaned"))
        filter_bar.addWidget(self.pill_cleaned)

        tab_v.addLayout(filter_bar)

        splitter = QSplitter(Qt.Horizontal)

        # Table
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels([
            "#", "Final Name", "Phone Number", "Prefix", "Core Name", "Suffix", "Sources", "Newest Date"
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Interactive)
        self.table.setColumnWidth(2, 180)
        self.table.setColumnWidth(3, 80)
        self.table.setColumnWidth(5, 140)
        self.table.setColumnWidth(7, 110)
        self.table.itemSelectionChanged.connect(self.on_table_row_selected)
        splitter.addWidget(self.table)

        # Live Inspector Drawer
        self.inspector = QFrame()
        self.inspector.setStyleSheet("""
            QFrame {
                background-color: #111827;
                border: 1px solid #1e293b;
                border-radius: 10px;
            }
        """)
        self.inspector.setMinimumWidth(380)
        self.inspector.setMaximumWidth(520)
        insp_v = QVBoxLayout(self.inspector)
        insp_v.setContentsMargins(16, 16, 16, 16)
        insp_v.setSpacing(12)

        insp_header = QHBoxLayout()
        self.avatar_lbl = QLabel("ID")
        self.avatar_lbl.setFixedSize(54, 54)
        self.avatar_lbl.setAlignment(Qt.AlignCenter)
        self.avatar_lbl.setStyleSheet("""
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #2563eb, stop:1 #06b6d4);
            color: #ffffff;
            font-size: 20px;
            font-weight: 800;
            border-radius: 27px;
        """)
        insp_header.addWidget(self.avatar_lbl)

        name_box = QVBoxLayout()
        self.insp_name = QLabel("Select a contact")
        self.insp_name.setStyleSheet("font-size: 17px; font-weight: 800; color: #ffffff;")
        self.insp_name.setWordWrap(True)
        self.insp_phone = QLabel("+91 ----- -----")
        self.insp_phone.setStyleSheet("font-size: 13px; font-weight: 600; color: #10b981;")
        name_box.addWidget(self.insp_name)
        name_box.addWidget(self.insp_phone)
        insp_header.addLayout(name_box)
        insp_v.addLayout(insp_header)

        insp_v.addWidget(QLabel("INTELLIGENT NAME BREAKDOWN"))
        self.tags_box = QHBoxLayout()
        self.tag_prefix = QLabel("PREFIX: -")
        self.tag_prefix.setStyleSheet("background-color: #1e293b; color: #38bdf8; padding: 3px 8px; border-radius: 4px; font-weight: 600; font-size: 11px;")
        self.tag_core = QLabel("CORE: -")
        self.tag_core.setStyleSheet("background-color: #1e293b; color: #e2e8f0; padding: 3px 8px; border-radius: 4px; font-weight: 600; font-size: 11px;")
        self.tag_suffix = QLabel("SUFFIX: -")
        self.tag_suffix.setStyleSheet("background-color: #1e293b; color: #f43f5e; padding: 3px 8px; border-radius: 4px; font-weight: 600; font-size: 11px;")
        self.tags_box.addWidget(self.tag_prefix)
        self.tags_box.addWidget(self.tag_core)
        self.tags_box.addWidget(self.tag_suffix)
        insp_v.addLayout(self.tags_box)

        insp_v.addWidget(QLabel("QUICK EDIT NAME"))
        edit_row = QHBoxLayout()
        self.edit_name_input = QLineEdit()
        self.edit_name_input.setPlaceholderText("Edit final name manually...")
        edit_row.addWidget(self.edit_name_input)
        save_edit_btn = QPushButton("Save Edit")
        save_edit_btn.clicked.connect(self.save_manual_edit)
        edit_row.addWidget(save_edit_btn)
        insp_v.addLayout(edit_row)

        insp_v.addWidget(QLabel("CHRONOLOGICAL SOURCE HISTORY"))
        self.history_text = QTextEdit()
        self.history_text.setReadOnly(True)
        self.history_text.setStyleSheet("background-color: #0b0f19; border: 1px solid #1e293b; border-radius: 6px; font-size: 12px; color: #cbd5e1;")
        insp_v.addWidget(self.history_text, 1)

        splitter.addWidget(self.inspector)
        splitter.setSizes([1050, 420])
        tab_v.addWidget(splitter, 1)

        self.tabs.addTab(self.tab_contacts, "Final Contacts (Google Ready)")

        # Tab 2: Duplicates Audit
        self.tab_dups = QWidget()
        tab_dup_v = QVBoxLayout(self.tab_dups)
        tab_dup_v.setContentsMargins(10, 10, 10, 10)
        self.dup_table = QTableWidget(0, 5)
        self.dup_table.setHorizontalHeaderLabels(["Phone Number", "Final Name", "Sources Count", "Newest Date", "All Associated Names"])
        self.dup_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.dup_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)
        self.dup_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.dup_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.dup_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.dup_table.setColumnWidth(1, 250)
        self.dup_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        tab_dup_v.addWidget(self.dup_table)
        self.tabs.addTab(self.tab_dups, "Number Duplicates Audit")

        # Tab 3: Discovered Fields
        self.tab_fields = QWidget()
        tab_f_v = QVBoxLayout(self.tab_fields)
        tab_f_v.setContentsMargins(10, 10, 10, 10)
        self.field_table = QTableWidget(0, 2)
        self.field_table.setHorizontalHeaderLabels(["VCF Property", "Total Occurrences"])
        self.field_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tab_f_v.addWidget(self.field_table)
        self.tabs.addTab(self.tab_fields, "Discovered Fields")

        # Tab 4: Scan Log
        self.tab_log = QWidget()
        tab_log_v = QVBoxLayout(self.tab_log)
        tab_log_v.setContentsMargins(10, 10, 10, 10)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        tab_log_v.addWidget(self.log_text)
        self.tabs.addTab(self.tab_log, "Scan Log")

        # 6. Export Toolbar
        export_bar = QHBoxLayout()
        export_bar.setSpacing(10)

        exp_all = QPushButton("⚡ Export All Formats to exports/")
        exp_all.setObjectName("ExportAllButton")
        exp_all.clicked.connect(self.export_all_formats)
        export_bar.addWidget(exp_all)

        exp_google_csv = QPushButton("📤 Google CSV (Official)")
        exp_google_csv.setObjectName("ExportButton")
        exp_google_csv.clicked.connect(self.export_google_csv)
        export_bar.addWidget(exp_google_csv)

        exp_vcf = QPushButton("📇 Google VCF 3.0")
        exp_vcf.setObjectName("ExportButton")
        exp_vcf.clicked.connect(self.export_google_vcf)
        export_bar.addWidget(exp_vcf)

        exp_xlsx = QPushButton("📊 Excel XLSX")
        exp_xlsx.setObjectName("ExportButton")
        exp_xlsx.clicked.connect(self.export_xlsx)
        export_bar.addWidget(exp_xlsx)

        exp_csv = QPushButton("📄 Clean CSV")
        exp_csv.setObjectName("ExportButton")
        exp_csv.clicked.connect(self.export_clean_csv)
        export_bar.addWidget(exp_csv)

        exp_dup_rep = QPushButton("📋 Duplicates Audit CSV")
        exp_dup_rep.setObjectName("ExportButton")
        exp_dup_rep.clicked.connect(self.export_duplicate_report)
        export_bar.addWidget(exp_dup_rep)

        main_layout.addLayout(export_bar)

    def choose_folder(self):
        f = QFileDialog.getExistingDirectory(self, "Select Contacts Folder", self.folder_input.text())
        if f:
            self.folder_input.setText(f)

    def start_scan(self):
        folder = self.folder_input.text().strip()
        if not folder or not Path(folder).is_dir():
            QMessageBox.warning(self, "Invalid Directory", "Please select a valid folder containing .vcf or .csv contact files.")
            return

        self.scan_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_lbl.setText("Scanning directory and parsing contacts...")

        self.worker = ScanWorker(
            folder=folder,
            recursive=self.recursive_chk.isChecked(),
            include_csv=self.csv_chk.isChecked()
        )
        self.worker.progress.connect(self.on_scan_progress)
        self.worker.finished.connect(self.on_scan_finished)
        self.worker.failed.connect(self.on_scan_failed)
        self.worker.start()

    def on_scan_progress(self, val, msg):
        self.progress_bar.setValue(val)
        self.status_lbl.setText(msg)

    def on_scan_failed(self, err):
        self.scan_btn.setEnabled(True)
        self.status_lbl.setText("Scan failed.")
        QMessageBox.critical(self, "Scan Error", f"An error occurred during scan:\n\n{err}")

    def on_scan_finished(self, result):
        self.result = result
        self.scan_btn.setEnabled(True)
        self.progress_bar.setValue(100)
        self.status_lbl.setText("Ready.")

        contacts = result["contacts"]
        dup_count = sum(1 for c in contacts if len(c["all_raw_names"]) > 1)
        cleaned_count = sum(1 for c in contacts if c["name"] != c["raw_chosen"])
        tagged_count = sum(1 for c in contacts if c["prefix"] or c["suffix"])

        self.card_files.set_value(f"{len(result['files']):,}")
        self.card_raw.set_value(f"{result['raw_cards']:,}")
        self.card_unique.set_value(f"{len(contacts):,}")
        self.card_merged.set_value(f"{dup_count:,}")
        self.card_formatted.set_value(f"{cleaned_count:,}")
        self.card_tags.set_value(f"{tagged_count:,}")

        log_lines = [
            f"=== SCAN COMPLETED AT {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===",
            f"Source Folder: {self.folder_input.text()}",
            f"Total Files Scanned: {len(result['files'])}",
            f"Raw vCards / Rows: {result['raw_cards']:,}",
            f"Valid Phone Associations: {result['total_associations']:,}",
            f"Final Unique Phone Numbers: {len(contacts):,}",
            f"Number Duplicate Groups Merged: {dup_count:,}",
            f"Names Formatted / Cleaned: {cleaned_count:,}",
            f"Prefix / Suffix Identified: {tagged_count:,}",
            f"Discarded No-Phone: {result['discarded_no_phone']:,}",
            f"Discarded No-Name: {result['discarded_no_name']:,}",
            "",
            "INTELLIGENCE RULES ENFORCED:",
            "- Unique Key: Strictly 1 Contact Record per Phone Number.",
            "- No Repetitive Words (e.g. 'Vyapar Vyapar' -> 'Vyapar').",
            "- No Trailing Counter Digits (e.g. 'Suraj Sharma 1' -> 'Suraj Sharma').",
            "- Prefix / Suffix formatted in ALL CAPS.",
            "- Core Personal Names formatted in Proper Title Case.",
            "- Chronological Resolution: Prioritized newest source names (2019-2026)."
        ]
        if result["parse_errors"]:
            log_lines += ["", f"Parse Errors ({len(result['parse_errors'])}):"]
            log_lines.extend(result["parse_errors"][:100])
        self.log_text.setPlainText("\n".join(log_lines))

        self.apply_filters_and_refresh()
        self.populate_duplicates_table()
        self.populate_fields_table()

        QMessageBox.information(
            self,
            "Scan Complete",
            f"Successfully processed {len(result['files'])} files!\n\n"
            f"• Unique Phone Numbers: {len(contacts):,}\n"
            f"• Duplicates Merged: {dup_count:,}\n"
            f"• Names Intelligently Formatted: {cleaned_count:,}\n\n"
            "All contacts are ready for Google Contacts export."
        )

    def set_active_filter(self, filter_name):
        self.active_filter = filter_name
        self.pill_all.setChecked(filter_name == "all")
        self.pill_tags.setChecked(filter_name == "tags")
        self.pill_dups.setChecked(filter_name == "duplicates")
        self.pill_cleaned.setChecked(filter_name == "cleaned")
        self.apply_filters_and_refresh()

    def apply_filters_and_refresh(self):
        if not self.result:
            return

        query = self.search_input.text().strip().casefold()
        filtered = []

        for c in self.result["contacts"]:
            if self.active_filter == "tags" and not (c["prefix"] or c["suffix"]):
                continue
            if self.active_filter == "duplicates" and len(c["all_raw_names"]) <= 1:
                continue
            if self.active_filter == "cleaned" and c["name"] == c["raw_chosen"]:
                continue

            if query:
                haystack = f"{c['name']} {c['number']} {c['prefix']} {c['core_name']} {c['suffix']} {c['notes']}".casefold()
                if query not in haystack:
                    continue

            filtered.append(c)

        self.current_filtered_contacts = filtered
        self.table.setRowCount(len(filtered))

        for row, c in enumerate(filtered):
            item_idx = QTableWidgetItem(str(row + 1))
            item_idx.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, item_idx)

            item_name = QTableWidgetItem(c["name"])
            item_name.setFont(QFont("Segoe UI", 10, QFont.Bold))
            self.table.setItem(row, 1, item_name)

            item_num = QTableWidgetItem(format_indian_number(c["number"]))
            item_num.setForeground(QColor("#34d399"))
            item_num.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.table.setItem(row, 2, item_num)

            item_pref = QTableWidgetItem(c["prefix"])
            if c["prefix"]:
                item_pref.setForeground(QColor("#38bdf8"))
            self.table.setItem(row, 3, item_pref)

            item_core = QTableWidgetItem(c["core_name"])
            self.table.setItem(row, 4, item_core)

            item_suf = QTableWidgetItem(c["suffix"])
            if c["suffix"]:
                item_suf.setForeground(QColor("#f43f5e"))
            self.table.setItem(row, 5, item_suf)

            item_src = QTableWidgetItem(str(c["source_count"]))
            item_src.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 6, item_src)

            item_date = QTableWidgetItem(c["newest_date"])
            item_date.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 7, item_date)

        if filtered:
            self.table.selectRow(0)

    def on_table_row_selected(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.current_filtered_contacts):
            return

        c = self.current_filtered_contacts[row]
        name = c["name"].strip()
        words = name.split()
        initials = (words[0][0] + (words[1][0] if len(words) > 1 else "")).upper() if words else "ID"
        self.avatar_lbl.setText(initials)

        self.insp_name.setText(c["name"])
        self.insp_phone.setText(f"{format_indian_number(c['number'])} ({c['kind'].upper()})")

        self.tag_prefix.setText(f"PREFIX: {c['prefix'] or 'None'}")
        self.tag_core.setText(f"CORE: {c['core_name'] or 'None'}")
        self.tag_suffix.setText(f"SUFFIX: {c['suffix'] or 'None'}")

        self.edit_name_input.setText(c["name"])

        lines = []
        lines.append(f"<b>Chosen Canonical Name:</b> {c['name']}")
        lines.append(f"<b>Raw Source Selected:</b> {c['raw_chosen']}")
        lines.append(f"<b>Newest Date:</b> {c['newest_date']}")
        lines.append("<br><b>All Sources Found Across Files:</b>")
        for h in c["history"]:
            badge = "<span style='color:#10b981; font-weight:bold;'>[NEWEST]</span>" if h["is_newest"] else ""
            lines.append(f"• <b>{h['raw_name']}</b> — {h['source_file']} ({h['date']}) {badge}")

        if c["notes"]:
            lines.append(f"<br><b>Google Export Notes:</b><br>{c['notes']}")

        self.history_text.setHtml("<br>".join(lines))

    def save_manual_edit(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.current_filtered_contacts):
            return
        c = self.current_filtered_contacts[row]
        new_name = self.edit_name_input.text().strip()
        if not new_name:
            return

        c["name"] = new_name
        parsed = parse_and_format_name(new_name)
        c["prefix"] = parsed["prefix"]
        c["core_name"] = parsed["core_name"]
        c["suffix"] = parsed["suffix"]

        self.table.item(row, 1).setText(new_name)
        self.table.item(row, 3).setText(c["prefix"])
        self.table.item(row, 4).setText(c["core_name"])
        self.table.item(row, 5).setText(c["suffix"])

        self.on_table_row_selected()
        QMessageBox.information(self, "Name Updated", f"Updated contact name to:\n{new_name}")

    def populate_duplicates_table(self):
        contacts = [c for c in self.result["contacts"] if len(c["all_raw_names"]) > 1]
        self.dup_table.setRowCount(len(contacts))

        for row, c in enumerate(contacts):
            self.dup_table.setItem(row, 0, QTableWidgetItem(format_indian_number(c["number"])))
            self.dup_table.setItem(row, 1, QTableWidgetItem(c["name"]))
            self.dup_table.setItem(row, 2, QTableWidgetItem(str(c["source_count"])))
            self.dup_table.setItem(row, 3, QTableWidgetItem(c["newest_date"]))
            self.dup_table.setItem(row, 4, QTableWidgetItem(" | ".join(c["all_raw_names"])))

    def populate_fields_table(self):
        fields = sorted(self.result["fields"].items(), key=lambda x: (-x[1], x[0]))
        self.field_table.setRowCount(len(fields))
        for row, (f, count) in enumerate(fields):
            self.field_table.setItem(row, 0, QTableWidgetItem(f))
            self.field_table.setItem(row, 1, QTableWidgetItem(f"{count:,}"))

    # =========================================================================
    # EXPORT IMPLEMENTATIONS
    # =========================================================================

    def export_all_formats(self):
        """One-click batch export of all formats into the exports/ directory."""
        if not self.result:
            QMessageBox.warning(self, "No Data", "Please perform a scan first.")
            return

        exports_dir = Path("exports")
        exports_dir.mkdir(exist_ok=True)

        contacts = self.result["contacts"]
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 1. Google CSV (chunked)
        self._write_google_csv_to_folder(exports_dir, chunk_size=3000)

        # 2. Google VCF 3.0
        vcf_path = exports_dir / f"google_contacts_{timestamp_str}.vcf"
        self._write_vcf_to_path(vcf_path)

        # 3. Excel XLSX
        xlsx_path = exports_dir / f"final_contacts_{timestamp_str}.xlsx"
        self._write_xlsx_to_path(xlsx_path)

        # 4. Clean 3-Column CSV
        clean_csv_path = exports_dir / f"final_contacts_clean_{timestamp_str}.csv"
        self._write_clean_csv_to_path(clean_csv_path)

        # 5. Duplicates Audit CSV
        dup_path = exports_dir / f"duplicates_audit_{timestamp_str}.csv"
        self._write_dup_report_to_path(dup_path)

        QMessageBox.information(
            self,
            "One-Click Export Complete",
            f"All export formats successfully generated in:\n{exports_dir.resolve()}\n\n"
            f"• Google Contacts CSV (<=3,000 chunked)\n"
            f"• Google vCard 3.0\n"
            f"• Excel XLSX with styling\n"
            f"• Clean 3-column CSV\n"
            f"• Duplicates Audit Report\n"
            f"• Total Contacts: {len(contacts):,}"
        )

    def _write_google_csv_to_folder(self, folder, chunk_size=3000):
        contacts = self.result["contacts"]
        headers = [
            "Name", "Given Name", "Additional Name", "Family Name", "Yomi Name",
            "Given Name Yomi", "Additional Name Yomi", "Family Name Yomi",
            "Name Prefix", "Name Suffix", "Initials", "Nickname", "Short Name",
            "Maiden Name", "Birthday", "Gender", "Location", "Billing Information",
            "Directory Server", "Mileage", "Occupation", "Hobby", "Sensitivity",
            "Priority", "Subject", "Notes", "Language", "Photo", "Group Membership",
            "Phone 1 - Type", "Phone 1 - Value"
        ]

        total_files = (len(contacts) + chunk_size - 1) // chunk_size if contacts else 1
        created_paths = []

        for idx in range(total_files):
            start = idx * chunk_size
            end = min(start + chunk_size, len(contacts))
            batch = contacts[start:end]

            fname = "Google_Contacts.csv" if total_files == 1 else f"Google_Contacts_Part_{idx + 1:02d}.csv"
            out_path = Path(folder) / fname
            created_paths.append(str(out_path))

            with open(out_path, "w", encoding="utf-8-sig", newline="") as fp:
                writer = csv.writer(fp)
                writer.writerow(headers)
                for c in batch:
                    core_words = c["core_name"].split()
                    given_name = core_words[0] if core_words else ""
                    family_name = " ".join(core_words[1:]) if len(core_words) > 1 else ""

                    row = [
                        c["name"], given_name, "", family_name,
                        "", "", "", "",
                        c["prefix"], c["suffix"],
                        "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                        c["notes"], "", "",
                        "* myContacts",
                        "Mobile" if c["kind"] == "mobile" else "Work",
                        format_indian_number(c["number"])
                    ]
                    writer.writerow(row)
        return created_paths

    def _write_vcf_to_path(self, path):
        with open(path, "w", encoding="utf-8", newline="") as fp:
            for c in self.result["contacts"]:
                fp.write("BEGIN:VCARD\r\n")
                fp.write("VERSION:3.0\r\n")
                escaped_name = c["name"].replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
                fp.write(f"FN:{escaped_name}\r\n")
                core_words = c["core_name"].split()
                given = core_words[0] if core_words else ""
                family = " ".join(core_words[1:]) if len(core_words) > 1 else ""
                fp.write(f"N:{family};{given};;{c['prefix']};{c['suffix']}\r\n")
                num_str = format_indian_number(c["number"])
                fp.write(f"TEL;TYPE=CELL,PREF:{num_str}\r\n")
                if c["notes"]:
                    esc_note = c["notes"].replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")
                    fp.write(f"NOTE:{esc_note}\r\n")
                fp.write("END:VCARD\r\n")

    def _write_xlsx_to_path(self, path):
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
        except ImportError:
            return False

        wb = Workbook()
        ws = wb.active
        ws.title = "Contacts"

        headers = ["#", "Final Name", "Phone Number", "Prefix", "Core Name", "Suffix", "Sources", "Newest Date", "Notes"]
        ws.append(headers)

        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for idx, c in enumerate(self.result["contacts"], 1):
            ws.append([
                idx,
                c["name"],
                format_indian_number(c["number"]),
                c["prefix"],
                c["core_name"],
                c["suffix"],
                c["source_count"],
                c["newest_date"],
                c["notes"]
            ])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 60)

        wb.save(path)
        return True

    def _write_clean_csv_to_path(self, path):
        with open(path, "w", encoding="utf-8-sig", newline="") as fp:
            writer = csv.writer(fp)
            writer.writerow(["Name", "Number", "Notes"])
            for c in self.result["contacts"]:
                writer.writerow([
                    c["name"],
                    format_indian_number(c["number"]),
                    c["notes"]
                ])

    def _write_dup_report_to_path(self, path):
        with open(path, "w", encoding="utf-8-sig", newline="") as fp:
            writer = csv.writer(fp)
            writer.writerow(["Phone Number", "Final Chosen Name", "Sources Count", "Newest Date", "All Associated Names"])
            for c in self.result["contacts"]:
                if len(c["all_raw_names"]) > 1:
                    writer.writerow([
                        format_indian_number(c["number"]),
                        c["name"],
                        c["source_count"],
                        c["newest_date"],
                        " | ".join(c["all_raw_names"])
                    ])

    def export_google_csv(self):
        if not self.result:
            QMessageBox.warning(self, "No Data", "Please perform a scan first.")
            return
        default_dir = str(Path("exports").resolve())
        folder = QFileDialog.getExistingDirectory(self, "Select Folder for Google Contacts CSV Export", default_dir)
        if not folder:
            return
        paths = self._write_google_csv_to_folder(folder, chunk_size=3000)
        QMessageBox.information(
            self,
            "Google CSV Export Complete",
            f"Successfully generated {len(paths)} Google-compatible CSV file(s) in:\n{folder}\n\n"
            f"• Total Contacts: {len(self.result['contacts']):,}\n"
            f"• Maximum per file: 3,000 (Respects Google Contacts import limits)"
        )

    def export_google_vcf(self):
        if not self.result:
            QMessageBox.warning(self, "No Data", "Please perform a scan first.")
            return
        default_path = str((Path("exports") / "Google_Contacts.vcf").resolve())
        path, _ = QFileDialog.getSaveFileName(self, "Export Google VCF 3.0", default_path, "vCard Files (*.vcf)")
        if not path:
            return
        self._write_vcf_to_path(path)
        QMessageBox.information(self, "VCF Export Complete", f"Successfully saved Google VCF 3.0 to:\n{path}")

    def export_xlsx(self):
        if not self.result:
            QMessageBox.warning(self, "No Data", "Please perform a scan first.")
            return
        default_path = str((Path("exports") / "Final_Contacts.xlsx").resolve())
        path, _ = QFileDialog.getSaveFileName(self, "Export Excel XLSX", default_path, "Excel Files (*.xlsx)")
        if not path:
            return
        success = self._write_xlsx_to_path(path)
        if success:
            QMessageBox.information(self, "XLSX Export Complete", f"Successfully saved Excel file to:\n{path}")
        else:
            QMessageBox.warning(self, "Package Missing", "Please install openpyxl to export to Excel.")

    def export_clean_csv(self):
        if not self.result:
            QMessageBox.warning(self, "No Data", "Please perform a scan first.")
            return
        default_path = str((Path("exports") / "Final_Contacts_Clean.csv").resolve())
        path, _ = QFileDialog.getSaveFileName(self, "Export 3-Column CSV", default_path, "CSV Files (*.csv)")
        if not path:
            return
        self._write_clean_csv_to_path(path)
        QMessageBox.information(self, "CSV Export Complete", f"Successfully saved clean CSV to:\n{path}")

    def export_duplicate_report(self):
        if not self.result:
            QMessageBox.warning(self, "No Data", "Please perform a scan first.")
            return
        default_path = str((Path("exports") / "Number_Duplicates_Audit.csv").resolve())
        path, _ = QFileDialog.getSaveFileName(self, "Export Duplicates Audit Report", default_path, "CSV Files (*.csv)")
        if not path:
            return
        self._write_dup_report_to_path(path)
        QMessageBox.information(self, "Audit Report Complete", f"Saved duplicates audit report to:\n{path}")

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait(1000)
        event.accept()

# =============================================================================
# MAIN ENTRYPOINT
# =============================================================================

def main():
    try:
        app = QApplication.instance() or QApplication(sys.argv)
        app.setStyle("Fusion")
        win = MainWindow()
        win.show()
        sys.exit(app.exec())
    except SystemExit:
        pass
    except Exception:
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
