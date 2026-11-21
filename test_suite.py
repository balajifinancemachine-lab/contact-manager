import sys
import os
import re
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Test standalone contact_manager.pyw directly
import importlib.util
spec = importlib.util.spec_from_file_location("contact_manager", "contact_manager.pyw")
contact_manager = importlib.util.module_from_spec(spec)
sys.modules["contact_manager"] = contact_manager
spec.loader.exec_module(contact_manager)

ScanWorker = contact_manager.ScanWorker

print("Running deep verification of standalone VCF Intelligent Contact Manager Pro...")

# Target imports directory where files were cleanly reorganized
target_dir = "imports" if Path("imports").exists() else "."
print(f"Target scan directory: {target_dir}")

worker = ScanWorker(folder=target_dir, recursive=False, include_csv=True)

class TestReceiver:
    def __init__(self):
        self.done = False
        self.result = None
        self.error = None

    def on_progress(self, val, msg):
        if val % 25 == 0 or val == 100:
            print(f"Progress: {val}% - {msg}")

    def on_finished(self, res):
        self.done = True
        self.result = res

    def on_failed(self, err):
        self.done = True
        self.error = err

rec = TestReceiver()
worker.progress.connect(rec.on_progress)
worker.finished.connect(rec.on_finished)
worker.failed.connect(rec.on_failed)

worker.run()

if rec.error:
    print("FAILED:", rec.error)
    sys.exit(1)

res = rec.result
print("\n" + "="*60)
print("VERIFICATION RESULTS")
print("="*60)
print(f"Files processed: {len(res['files'])}")
print(f"Raw cards read: {res['raw_cards']:,}")
print(f"Valid phone associations: {res['total_associations']:,}")
print(f"Final unique contacts: {len(res['contacts']):,}")

contacts = res["contacts"]

print("\n--- SAMPLE CONTACTS ---")
for c in contacts[100:112]:
    print(f"Name: {c['name']:35} | Phone: {c['number']:10} | Pref: {c['prefix']:5} | Suf: {c['suffix']:10} | Newest: {c['newest_date']}")

# 1. Verification: Exactly 1 phone number per contact
nums = [c["number"] for c in contacts]
assert len(contacts) == len(set(nums)), "ERROR: Duplicate numbers found in final contacts!"
print("\n[PASS] Strict 1-number-per-contact constraint verified! Exactly 1 unique number per record.")

# 2. Verification: Check specific previously corrupted names
check_numbers = {
    "9197527746": "Suraj Sharma",
    "9617036609": "Suraj Sharma",
    "9752774668": "Suraj Sharma",
    "7738607620": "Adi Shetty KOTAK",
    "9594195639": "Aditya Gupta",
    "8591190451": "Aditya Rokade Customer Flower Shop",
    "9619642570": "Ahmad Shaikh MOBILE SHOP ASALFA",
    "9819248643": "Ajay Pashte CHUNABHATTI BALAJI"
}

print("\n--- CHECKING PREVIOUSLY MANGLED CONTACTS ---")
for c in contacts:
    if c["number"] in check_numbers:
        expected_hint = check_numbers[c["number"]]
        print(f"Number {c['number']}: Name = '{c['name']}' (Expected hint: '{expected_hint}')")
        assert not re.search(r"([a-zA-Z])\1{2,}", c["name"]), f"Triple letter found in {c['name']}"
        assert not re.search(r"\s+\d+$", c["name"]), f"Trailing counter number found in {c['name']}"

print("\n[PASS] No trailing numbers (' 1', ' 2') or triple letters ('Surajjj') detected in verified contacts!")
print("[PASS] All verification assertions succeeded!")
