"""
Tes fungsional untuk status 'Diluar Modul'.
Memverifikasi bahwa:
  1. SEP dengan diagnosa TIDAK tercakup modul -> "Diluar Modul"
  2. SEP dengan diagnosa tercakup dan lulus    -> "ACC"
  3. SEP dengan diagnosa tercakup dan gagal    -> "PERLU REVISI"
  4. SEP dengan cek berkas fisik               -> "ACC (Cek Berkas Fisik)"
  5. Syntax validator.py dan app.py valid
"""

import ast
import sys
from pathlib import Path

# --- 1. Syntax check ---
print("=" * 60)
print("TEST 1: Syntax Check")
print("=" * 60)

errors = []
for fname in ("validator.py", "app.py"):
    fpath = Path(__file__).parent / fname
    try:
        ast.parse(fpath.read_text(encoding="utf-8"))
        print(f"  [PASS] {fname} — syntax OK")
    except SyntaxError as e:
        print(f"  [FAIL] {fname} — {e}")
        errors.append(fname)

if errors:
    print(f"\nSyntax errors in: {errors}")
    sys.exit(1)

# --- 2. Import ---
print()
print("=" * 60)
print("TEST 2: Import & Logic Tests")
print("=" * 60)

import pandas as pd
from validator import evaluate_row, load_rules

rules_data = load_rules()
passed = 0
failed = 0

def run_test(name, diaglist, proclist, expected_status, vent_hour=""):
    global passed, failed
    row = pd.Series({
        "SEP": "TEST-001",
        "DIAGLIST": diaglist,
        "PROCLIST": proclist,
        "VENT_HOUR": vent_hour,
    })
    result = evaluate_row(row, rules_data)
    actual = result.status
    ok = actual == expected_status
    tag = "PASS" if ok else "FAIL"
    print(f"  [{tag}] {name}")
    print(f"         DIAGLIST={diaglist}, PROCLIST={proclist}")
    print(f"         Expected: {expected_status}, Got: {actual}")
    if not ok:
        print(f"         any_rule_triggered={result.any_rule_triggered}")
        print(f"         matched_rules={result.matched_rules}")
        print(f"         revisions={result.revisions}")
        print(f"         physical_checks={result.physical_checks}")
        failed += 1
    else:
        passed += 1

# Test A: Diagnosa TIDAK ada di modul -> "Diluar Modul"
run_test(
    "Diagnosa Z00.0 (tidak di modul) -> Diluar Modul",
    "Z00.0", "", "Diluar Modul"
)

# Test B: Diagnosa Z99.9 (tidak di modul) -> "Diluar Modul"
run_test(
    "Diagnosa Z99.9 (tidak di modul) -> Diluar Modul",
    "Z99.9", "", "Diluar Modul"
)

# Test C: I50.9 (Gagal Jantung) DENGAN prosedur lengkap -> ACC (Cek Berkas Fisik)
#         Karena ada physical_document_check "bacaan ekokardiografi 6 bulan terakhir"
run_test(
    "I50.9 + prosedur 89.52;88.72;87.49 -> ACC (Cek Berkas Fisik)",
    "I50.9", "89.52;88.72;87.49", "ACC (Cek Berkas Fisik)"
)

# Test C2: J15.9 (Pneumonia Bakterial) DENGAN prosedur lengkap -> ACC (Cek Berkas Fisik)
#         PARU-002 memiliki physical_document_check PSI Score
run_test(
    "J15.9 + prosedur 87.49;90.42 -> ACC (Cek Berkas Fisik)",
    "J15.9", "87.49;90.42", "ACC (Cek Berkas Fisik)"
)

# Test D: I50.9 (Gagal Jantung) TANPA prosedur -> PERLU REVISI
run_test(
    "I50.9 tanpa prosedur -> PERLU REVISI",
    "I50.9", "", "PERLU REVISI"
)

# Test E: I21.0 (STEMI) + streptokinase -> ACC (Cek Berkas Fisik)
run_test(
    "I21.0 + 99.10 (streptokinase) -> ACC (Cek Berkas Fisik)",
    "I21.0", "89.52;88.72;87.49;99.10", "ACC (Cek Berkas Fisik)"
)

# Test F: Multiple diagnosa tapi semua diluar modul -> Diluar Modul
run_test(
    "Z00.0;Z01.1 (semua diluar modul) -> Diluar Modul",
    "Z00.0;Z01.1", "", "Diluar Modul"
)

# Test G: Satu diagnosa modul + satu diluar -> seharusnya trigger modul
run_test(
    "I50.9;Z00.0 tanpa prosedur -> PERLU REVISI (I50.9 trigger modul)",
    "I50.9;Z00.0", "", "PERLU REVISI"
)

print()
print("=" * 60)
print(f"RESULTS: {passed} passed, {failed} failed out of {passed + failed}")
print("=" * 60)

if failed:
    sys.exit(1)
else:
    print("\n✅ ALL TESTS PASSED!")
    sys.exit(0)
