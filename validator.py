"""
validator.py — Mesin Validasi Deterministik Berbasis Aturan
Validator Kelengkapan Berkas & Coding BPJS Rawat Inap

Modul ini membaca rules.json dan mengevaluasi setiap baris klaim (SEP)
terhadap seluruh aturan yang berlaku, menghasilkan status validasi:
  - ACC
  - ACC (Cek Berkas Fisik)
  - PERLU REVISI
"""

import json
import re
from pathlib import Path
from typing import Any

import pandas as pd

# ---------------------------------------------------------------------------
# 1. UTILITAS PENCOCOKAN KODE ICD
# ---------------------------------------------------------------------------

def _normalize_code(code: str) -> str:
    """Bersihkan spasi dan ubah ke uppercase."""
    return code.strip().upper()


def _code_matches_pattern(code: str, pattern: str) -> bool:
    """
    Cocokkan satu *code* terhadap satu *pattern*.
    Mendukung:
      - exact  : "I50.9" == "I50.9"
      - prefix : "I11.-" atau "I11*" cocok dengan "I11", "I11.0", "I11.9"
      - bare   : "I48" cocok exact "I48" atau prefix "I48.0", "I48.1" dll
    """
    code = _normalize_code(code)
    pattern = _normalize_code(pattern)

    # Wildcard suffix  .-  atau  *
    if pattern.endswith(".-"):
        prefix = pattern[:-2]
        return code == prefix or code.startswith(prefix + ".")
    if pattern.endswith("*"):
        prefix = pattern[:-1]
        return code.startswith(prefix)

    # Exact match
    return code == pattern


def _any_code_matches(codes: list[str], patterns: list[str]) -> list[str]:
    """Kembalikan daftar kode dari *codes* yang cocok minimal satu pattern."""
    matched: list[str] = []
    for c in codes:
        for p in patterns:
            if _code_matches_pattern(c, p):
                matched.append(c)
                break
    return matched


def _has_any_match(codes: list[str], patterns: list[str]) -> bool:
    return len(_any_code_matches(codes, patterns)) > 0


# ---------------------------------------------------------------------------
# 2. PARSING FIELD DARI BARIS KLAIM
# ---------------------------------------------------------------------------

def _parse_list_field(value: Any) -> list[str]:
    """Pecah field DIAGLIST / PROCLIST (pemisah ; atau ,) menjadi list kode."""
    if pd.isna(value) or str(value).strip() == "":
        return []
    raw = str(value).strip()
    # Coba titik koma dulu, lalu koma
    if ";" in raw:
        parts = raw.split(";")
    elif "," in raw:
        parts = raw.split(",")
    else:
        parts = [raw]
    return [_normalize_code(p) for p in parts if p.strip()]


def _parse_vent_hour(value: Any) -> float | None:
    if pd.isna(value) or str(value).strip() == "":
        return None
    try:
        return float(str(value).strip().replace(",", "."))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# 3. LOADER ATURAN
# ---------------------------------------------------------------------------

def load_rules(path: str | Path | None = None) -> dict:
    """Muat rules.json — gunakan path default jika tidak diberikan."""
    if path is None:
        path = Path(__file__).parent / "rules.json"
    else:
        path = Path(path)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# 4. EVALUASI ATURAN PER BARIS
# ---------------------------------------------------------------------------

class ValidationResult:
    """Hasil validasi untuk satu SEP."""

    def __init__(self):
        self.revisions: list[str] = []
        self.physical_checks: list[str] = []
        self.matched_rules: list[str] = []  # deskripsi singkat aturan terpenuhi
        self.any_rule_triggered: bool = False  # ada aturan modul yang berlaku?

    @property
    def status(self) -> str:
        if self.revisions:
            return "PERLU REVISI"
        if self.physical_checks:
            return "ACC (Cek Berkas Fisik)"
        if self.matched_rules or self.any_rule_triggered:
            return "ACC"
        return "Diluar Modul"

    @property
    def status_emoji(self) -> str:
        mapping = {
            "PERLU REVISI": "🔴",
            "ACC (Cek Berkas Fisik)": "🟡",
            "ACC": "🟢",
            "Diluar Modul": "⚪",
        }
        return mapping.get(self.status, "")


def _evaluate_standard_rule(rule: dict, diags: list[str], procs: list[str],
                             vent_hour: float | None, result: ValidationResult):
    """Evaluasi aturan standar: trigger diagnosis → cek required procedures."""
    trigger = rule.get("trigger_diagnoses", [])
    if not trigger:
        return

    matched_diags = _any_code_matches(diags, trigger)
    if not matched_diags:
        return  # aturan tidak berlaku

    result.any_rule_triggered = True
    matched_diag_str = ", ".join(matched_diags)

    # Cek required_procedures
    req_procs = rule.get("required_procedures")
    if req_procs and req_procs.get("codes"):
        mode = req_procs.get("mode", "all")
        required_codes = req_procs["codes"]

        # Cek pengecualian (misal Pneumonia + TB → hapus 90.42)
        exc = rule.get("exception_diagnoses_remove_procedure")
        if exc:
            if_diags = exc.get("if_diagnoses_present", [])
            if _has_any_match(diags, if_diags):
                remove_procs = exc.get("remove_procedures", [])
                required_codes = [c for c in required_codes if c not in remove_procs]

        if required_codes:
            if mode == "all":
                missing = [c for c in required_codes if not _has_any_match(procs, [c])]
                if missing:
                    msg = rule.get("revision_message", "Prosedur tidak lengkap")
                    msg = msg.replace("{matched_diag}", matched_diag_str)
                    msg += f" (kurang: {', '.join(missing)})"
                    result.revisions.append(msg)
                else:
                    result.matched_rules.append(rule.get("name", rule.get("id", "")))
            elif mode == "any":
                has_any = any(_has_any_match(procs, [c]) for c in required_codes)
                if not has_any:
                    msg = rule.get("revision_message", "Prosedur tidak lengkap")
                    msg = msg.replace("{matched_diag}", matched_diag_str)
                    result.revisions.append(msg)
                else:
                    result.matched_rules.append(rule.get("name", rule.get("id", "")))
        else:
            # Semua required dihapus oleh exception
            result.matched_rules.append(rule.get("name", rule.get("id", "")) + " (pengecualian berlaku)")

    elif req_procs is None:
        # Hanya physical check, tidak ada prosedur wajib
        result.matched_rules.append(rule.get("name", rule.get("id", "")))

    # Physical document check
    phys = rule.get("physical_document_check")
    if phys and matched_diags:
        result.physical_checks.append(f"[{rule.get('id','')}] {phys}")


def _evaluate_combination_check(rule: dict, diags: list[str], result: ValidationResult):
    """Evaluasi aturan kombinasi diagnosis yang salah."""
    group_a = rule.get("trigger_group_a", [])
    group_b = rule.get("trigger_group_b", [])
    if not group_a or not group_b:
        return

    has_a = _has_any_match(diags, group_a)
    has_b = _has_any_match(diags, group_b)

    if has_a and has_b:
        result.any_rule_triggered = True
        # Cek exception (misal sudah ada I13.-)
        exceptions = rule.get("exception_if_present", [])
        if exceptions and _has_any_match(diags, exceptions):
            result.matched_rules.append(rule.get("name", rule.get("id", "")) + " (sudah sesuai I13.-)")
            return
        result.revisions.append(rule.get("revision_message", "Kombinasi diagnosis perlu revisi"))


def _evaluate_single_diagnosis_check(rule: dict, diags: list[str], result: ValidationResult):
    """Evaluasi aturan diagnosis tunggal tanpa penyulit."""
    trigger = rule.get("trigger_diagnoses", [])
    if not trigger:
        return

    if not _has_any_match(diags, trigger):
        return

    result.any_rule_triggered = True

    if len(diags) == 1:
        result.revisions.append(rule.get("revision_message", "Diagnosis tunggal tanpa penyulit"))


def _evaluate_specific_code_check(rule: dict, diags: list[str], result: ValidationResult):
    """Evaluasi aturan kode tidak spesifik (misal CKD N18/N18.9)."""
    trigger = rule.get("trigger_diagnoses", [])
    if not trigger:
        return

    matched = _any_code_matches(diags, trigger)
    if matched:
        result.any_rule_triggered = True
        result.revisions.append(rule.get("revision_message", "Kode tidak spesifik"))


def _evaluate_procedure_combination_check(rule: dict, procs: list[str], result: ValidationResult):
    """Evaluasi kombinasi prosedur yang tidak seharusnya (misal intubasi + RJPO)."""
    required_procs = rule.get("trigger_procedures_all", [])
    if not required_procs:
        return

    all_present = all(_has_any_match(procs, [p]) for p in required_procs)
    if all_present:
        result.any_rule_triggered = True
        result.revisions.append(rule.get("revision_message", "Kombinasi prosedur perlu revisi"))


def _evaluate_ventilator_duration(rule: dict, procs: list[str],
                                   vent_hour: float | None, result: ValidationResult):
    """Evaluasi durasi ventilator < 6 jam."""
    trigger_procs = rule.get("trigger_procedures", [])
    if not trigger_procs:
        return

    has_vent_proc = _has_any_match(procs, trigger_procs)
    if not has_vent_proc:
        return

    result.any_rule_triggered = True

    if vent_hour is not None and vent_hour > 0 and vent_hour < rule.get("vent_hour_max_exclusive", 6):
        result.revisions.append(rule.get("revision_message", "Ventilator < 6 jam"))


def _evaluate_conditional_physical_check(rule: dict, diags: list[str],
                                          procs: list[str], result: ValidationResult):
    """Evaluasi cek berkas fisik bersyarat (misal streptokinase pada STEMI)."""
    trigger = rule.get("trigger_diagnoses", [])
    cond_proc = rule.get("conditional_procedure")
    if not trigger or not cond_proc:
        return

    if not _has_any_match(diags, trigger):
        return

    result.any_rule_triggered = True

    if _has_any_match(procs, [cond_proc]):
        phys = rule.get("physical_document_check")
        if phys:
            result.physical_checks.append(f"[{rule.get('id','')}] {phys}")


def _evaluate_delivery_complication(rule: dict, diags: list[str],
                                     procs: list[str], result: ValidationResult):
    """Evaluasi aturan persalinan: penyulit harus jadi Diagnosis Utama."""
    trigger_procs = rule.get("trigger_procedures", [])
    if not trigger_procs:
        return

    has_delivery_proc = _has_any_match(procs, trigger_procs)
    if not has_delivery_proc:
        return

    result.any_rule_triggered = True

    if len(diags) < 2:
        return

    delivery_codes = rule.get("delivery_method_codes", [])
    outcome_codes = rule.get("outcome_codes", ["Z37.-"])

    primary_diag = diags[0]
    secondary_diags = diags[1:]

    # Cek apakah primary diagnosis adalah kode persalinan normal (O80.-) atau outcome (Z37.-)
    is_primary_delivery = _has_any_match([primary_diag], delivery_codes)
    is_primary_outcome = _has_any_match([primary_diag], outcome_codes)

    # Cek apakah ada penyulit/komplikasi di sekunder (kode selain O80.- dan Z37.-)
    has_complication_in_secondary = False
    for d in secondary_diags:
        if not _has_any_match([d], delivery_codes) and not _has_any_match([d], outcome_codes):
            has_complication_in_secondary = True
            break

    if (is_primary_delivery or is_primary_outcome) and has_complication_in_secondary:
        result.revisions.append(rule.get("revision_message",
            "Penyulit harus menjadi Diagnosis Utama pada persalinan"))


def _evaluate_procedure_requires_diagnosis(rule: dict, diags: list[str],
                                            procs: list[str], result: ValidationResult):
    """Evaluasi prosedur yang memerlukan diagnosis tertentu (misal Perineoraphy)."""
    trigger_procs = rule.get("trigger_procedures", [])
    if not trigger_procs:
        return

    has_trigger = _has_any_match(procs, trigger_procs)
    if not has_trigger:
        return

    result.any_rule_triggered = True

    required_diags = rule.get("required_diagnoses", [])
    if not _has_any_match(diags, required_diags):
        result.revisions.append(rule.get("revision_message",
            "Prosedur memerlukan diagnosis tertentu"))
    else:
        result.matched_rules.append(rule.get("name", rule.get("id", "")))


def evaluate_row(row: pd.Series, rules_data: dict) -> ValidationResult:
    """
    Evaluasi seluruh aturan terhadap satu baris klaim.
    Mengembalikan objek ValidationResult.
    """
    diags = _parse_list_field(row.get("DIAGLIST", ""))
    procs = _parse_list_field(row.get("PROCLIST", ""))
    vent_hour = _parse_vent_hour(row.get("VENT_HOUR", None))

    result = ValidationResult()

    for category in rules_data.get("categories", []):
        for rule in category.get("rules", []):
            rule_type = rule.get("type", "standard")

            if rule_type == "combination_check":
                _evaluate_combination_check(rule, diags, result)
            elif rule_type == "single_diagnosis_check":
                _evaluate_single_diagnosis_check(rule, diags, result)
            elif rule_type == "specific_code_check":
                _evaluate_specific_code_check(rule, diags, result)
            elif rule_type == "procedure_combination_check":
                _evaluate_procedure_combination_check(rule, procs, result)
            elif rule_type == "ventilator_duration_check":
                _evaluate_ventilator_duration(rule, procs, vent_hour, result)
            elif rule_type == "conditional_physical_check":
                _evaluate_conditional_physical_check(rule, diags, procs, result)
            elif rule_type == "delivery_complication_check":
                _evaluate_delivery_complication(rule, diags, procs, result)
            elif rule_type == "procedure_requires_diagnosis":
                _evaluate_procedure_requires_diagnosis(rule, diags, procs, result)
            else:
                # standard rule
                _evaluate_standard_rule(rule, diags, procs, vent_hour, result)

    return result


# ---------------------------------------------------------------------------
# 5. VALIDASI MASSAL (DataFrame)
# ---------------------------------------------------------------------------

def validate_dataframe(df: pd.DataFrame, rules_data: dict) -> pd.DataFrame:
    """
    Validasi seluruh DataFrame dan kembalikan DataFrame dengan kolom tambahan:
      STATUS_VALIDASI, STATUS_EMOJI, KETERANGAN_REVISI, CEK_BERKAS_FISIK, ATURAN_TERPENUHI
    """
    statuses = []
    emojis = []
    revisions = []
    phys_checks = []
    matched = []

    for _, row in df.iterrows():
        res = evaluate_row(row, rules_data)
        statuses.append(res.status)
        emojis.append(res.status_emoji)
        revisions.append("\n".join(res.revisions) if res.revisions else "")
        phys_checks.append("\n".join(res.physical_checks) if res.physical_checks else "")
        matched.append("; ".join(res.matched_rules) if res.matched_rules else "")

    df = df.copy()
    df["STATUS_VALIDASI"] = statuses
    df["STATUS_EMOJI"] = emojis
    df["KETERANGAN_REVISI"] = revisions
    df["CEK_BERKAS_FISIK"] = phys_checks
    df["ATURAN_TERPENUHI"] = matched

    return df


# ---------------------------------------------------------------------------
# 6. SMART FILE LOADER
# ---------------------------------------------------------------------------

def smart_load_file(uploaded_file) -> pd.DataFrame:
    """
    Smart loader: coba baca sebagai TSV dulu (karena banyak file .xls ekspor
    SIMRS/INA-CBG sebenarnya TSV), lalu fallback ke read_excel.
    """
    import io

    filename = uploaded_file.name if hasattr(uploaded_file, "name") else str(uploaded_file)
    ext = Path(filename).suffix.lower()

    raw_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    # Deteksi apakah file biner Excel (magic bytes)
    is_binary_excel = False
    if raw_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":  # OLE2 (.xls)
        is_binary_excel = True
    elif raw_bytes[:4] == b"PK\x03\x04":  # ZIP (.xlsx)
        is_binary_excel = True

    if ext in (".tsv", ".txt") or (ext == ".xls" and not is_binary_excel):
        # Baca sebagai TSV
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(io.BytesIO(raw_bytes), sep="\t", dtype=str, encoding="utf-8")
            if len(df.columns) <= 1:
                raise ValueError("Only 1 column detected, probably not TSV")
            return df
        except Exception:
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(io.BytesIO(raw_bytes), sep="\t", dtype=str, encoding="latin-1")
                if len(df.columns) <= 1:
                    raise ValueError("Only 1 column detected")
                return df
            except Exception:
                pass

    if ext in (".xlsx", ".xls") or is_binary_excel:
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(io.BytesIO(raw_bytes), dtype=str)
            return df
        except Exception:
            pass

    # Fallback: coba CSV biasa
    try:
        df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str, encoding="utf-8")
        return df
    except Exception:
        df = pd.read_csv(io.BytesIO(raw_bytes), dtype=str, encoding="latin-1")
        return df
