"""
app.py — Antarmuka Utama Streamlit
Validator Kelengkapan Berkas & Coding BPJS Rawat Inap

Fitur:
  - Upload file data klaim (MIX) dengan smart loader
  - Dashboard ringkasan (KPI metrics)
  - Pencarian & filter SEP berdasarkan status, DPJP, diagnosis
  - Detail SEP via dialog (@st.dialog)
  - Edit status per-SEP & bulk edit dengan log riwayat perubahan
  - Tab Kelola Aturan: lihat, edit, tambah, download rules.json
  - Export hasil validasi ke Excel (.xlsx)
"""

import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from validator import (
    evaluate_row,
    load_rules,
    smart_load_file,
    validate_dataframe,
)

# ---------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Validator Berkas & Coding BPJS RI",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

RULES_PATH = Path(__file__).parent / "rules.json"

STATUS_OPTIONS = ["ACC", "ACC (Cek Berkas Fisik)", "PERLU REVISI", "Diluar Modul"]
STATUS_EMOJI_MAP = {
    "ACC": "🟢",
    "ACC (Cek Berkas Fisik)": "🟡",
    "PERLU REVISI": "🔴",
    "Diluar Modul": "⚪",
}
TUJUAN_OPTIONS = ["-", "Coder", "Admin Ruangan", "Petugas Entry"]

# ---------------------------------------------------------------------------
# SESSION STATE INITIALIZATION
# ---------------------------------------------------------------------------

if "status_overrides" not in st.session_state:
    # {sep_no: {"status": "...", "catatan": "...", "verifikator": "..."}}
    st.session_state["status_overrides"] = {}

if "change_history" not in st.session_state:
    # list of {sep, status_lama, status_baru, catatan, verifikator, waktu}
    st.session_state["change_history"] = []

if "pending_notes" not in st.session_state:
    # {sep_no: "keterangan perbaikan..."}
    st.session_state["pending_notes"] = {}

if "pending_file_bytes" not in st.session_state:
    st.session_state["pending_file_bytes"] = None
    st.session_state["pending_file_name"] = None

if "inline_editor_rev" not in st.session_state:
    st.session_state["inline_editor_rev"] = 0

if "last_verifikator" not in st.session_state:
    st.session_state["last_verifikator"] = ""


# ---------------------------------------------------------------------------
# UTILITAS
# ---------------------------------------------------------------------------

def get_rules() -> dict:
    """Muat rules, prioritaskan session_state jika user sudah edit."""
    if "rules_data" in st.session_state:
        return st.session_state["rules_data"]
    rules = load_rules(RULES_PATH)
    st.session_state["rules_data"] = rules
    return rules


def save_rules_to_disk(rules_data: dict):
    """Simpan rules_data kembali ke rules.json."""
    with open(RULES_PATH, "w", encoding="utf-8") as f:
        json.dump(rules_data, f, ensure_ascii=False, indent=2)


@st.cache_data(show_spinner=False)
def cached_load_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Cache pembacaan file agar tidak dibaca ulang setiap rerun."""
    buf = io.BytesIO(file_bytes)
    buf.name = filename
    return smart_load_file(buf)


@st.cache_data(show_spinner=False)
def cached_validate(df_work: pd.DataFrame, rules_json_str: str) -> pd.DataFrame:
    """Cache hasil validasi agar tidak mengulang evaluasi 12rb+ baris setiap interaksi."""
    rules_dict = json.loads(rules_json_str)
    return validate_dataframe(df_work, rules_dict)


def apply_status_overrides(df: pd.DataFrame) -> pd.DataFrame:
    """Terapkan override status dari session_state ke DataFrame (hanya baris yang diedit)."""
    overrides = st.session_state.get("status_overrides", {})
    if not overrides:
        return df

    df = df.copy()
    sep_series = df["SEP"].astype(str)
    mask = sep_series.isin(overrides.keys())
    for idx in df.index[mask]:
        sep = sep_series.at[idx]
        override = overrides[sep]
        new_status = override["status"]
        df.at[idx, "STATUS_VALIDASI"] = new_status
        df.at[idx, "STATUS_EMOJI"] = STATUS_EMOJI_MAP.get(new_status, "")

        catatan = override.get("catatan", "")
        verifikator = override.get("verifikator", "")
        tujuan = override.get("tujuan", "")
        tujuan_tag = f" (Tujuan: {tujuan})" if tujuan and tujuan != "-" else ""
        override_note = (
            f"[MANUAL] Status diubah oleh {verifikator}{tujuan_tag}: {catatan}"
            if catatan
            else f"[MANUAL] Status diubah oleh {verifikator}{tujuan_tag}"
        )

        existing_revisi = str(df.at[idx, "KETERANGAN_REVISI"]) if df.at[idx, "KETERANGAN_REVISI"] else ""
        if new_status == "PERLU REVISI" and catatan:
            df.at[idx, "KETERANGAN_REVISI"] = (existing_revisi + "\n" + override_note).strip()
        elif new_status in ("ACC", "ACC (Cek Berkas Fisik)", "Diluar Modul"):
            df.at[idx, "KETERANGAN_REVISI"] = override_note

    return df


def record_change(sep: str, status_lama: str, status_baru: str,
                   catatan: str, verifikator: str, tujuan: str = "-"):
    """Catat perubahan status ke log riwayat."""
    st.session_state["change_history"].append({
        "sep": sep,
        "status_lama": status_lama,
        "status_baru": status_baru,
        "catatan": catatan,
        "verifikator": verifikator,
        "tujuan": tujuan or "-",
        "waktu": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


def update_status(sep: str, status_lama: str, status_baru: str,
                   catatan: str, verifikator: str, tujuan: str = "-"):
    """Update override status untuk satu SEP dan catat log."""
    tujuan_clean = tujuan if tujuan and tujuan != "-" else ""
    st.session_state["status_overrides"][sep] = {
        "status": status_baru,
        "catatan": catatan,
        "verifikator": verifikator,
        "tujuan": tujuan_clean,
    }
    record_change(sep, status_lama, status_baru, catatan, verifikator, tujuan_clean or "-")


def build_excel_report(df_validated: pd.DataFrame) -> bytes:
    """Buat file Excel yang rapi dan berwarna dengan openpyxl styling."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()
    display_cols = [c for c in df_validated.columns
                    if c not in ("STATUS_EMOJI", "ATURAN_TERPENUHI")]

    # --- Palet warna ---
    HEADER_FILL = PatternFill("solid", fgColor="1F4E79")      # biru tua
    HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    ACC_FILL = PatternFill("solid", fgColor="D9F2D9")          # hijau muda
    CEK_FILL = PatternFill("solid", fgColor="FFF2CC")          # kuning muda
    REV_FILL = PatternFill("solid", fgColor="F4CCCC")          # merah muda
    DILUAR_FILL = PatternFill("solid", fgColor="E0E0E0")        # abu-abu muda
    ACC_FONT = Font(name="Calibri", bold=True, color="1B7A1B") # hijau tua
    CEK_FONT = Font(name="Calibri", bold=True, color="B45F06") # oranye tua
    REV_FONT = Font(name="Calibri", bold=True, color="CC0000") # merah tua
    DILUAR_FONT = Font(name="Calibri", bold=True, color="666666") # abu-abu tua
    NORMAL_FONT = Font(name="Calibri", size=10)
    TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
    SUBTITLE_FONT = Font(name="Calibri", italic=True, size=10, color="666666")
    THIN_BORDER = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
    CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
    CENTER_WRAP = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # Kolom lebar khusus (key=nama kolom, value=lebar)
    COL_WIDTHS = {
        "SEP": 26, "NAMA_PASIEN": 22, "MRN": 14, "DPJP": 20,
        "ADMISSION_DATE": 14, "DISCHARGE_DATE": 14, "LOS": 7,
        "INACBG": 12, "DESKRIPSI_INACBG": 30, "PTD": 6,
        "JENIS_RAWAT": 14,
        "DIAGLIST": 32, "PROCLIST": 28, "VENT_HOUR": 10,
        "STATUS_VALIDASI": 22, "KETERANGAN_REVISI": 45,
        "CEK_BERKAS_FISIK": 40,
    }

    # Label header yang lebih ramah pembaca
    COL_LABELS = {
        "SEP": "No. SEP",
        "NAMA_PASIEN": "Nama Pasien",
        "MRN": "No. RM",
        "DPJP": "DPJP",
        "ADMISSION_DATE": "Tgl Masuk",
        "DISCHARGE_DATE": "Tgl Keluar",
        "LOS": "LOS",
        "INACBG": "Kode INA-CBG",
        "DESKRIPSI_INACBG": "Deskripsi INA-CBG",
        "PTD": "PTD",
        "JENIS_RAWAT": "Jenis Rawat",
        "DIAGLIST": "Daftar Diagnosis (ICD-10)",
        "PROCLIST": "Daftar Prosedur (ICD-9-CM)",
        "VENT_HOUR": "Vent (Jam)",
        "STATUS_VALIDASI": "Status Validasi",
        "KETERANGAN_REVISI": "Keterangan Revisi",
        "CEK_BERKAS_FISIK": "Cek Berkas Fisik",
    }

    def _style_sheet(ws, df_sheet, cols, sheet_title, sheet_subtitle=""):
        """Styling helper: judul, header, baris berwarna, border, lebar kolom."""
        # -- Baris judul (merge) --
        ws.insert_rows(1, 2)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
        title_cell = ws.cell(row=1, column=1, value=sheet_title)
        title_cell.font = TITLE_FONT
        title_cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cols))
        sub_cell = ws.cell(row=2, column=1, value=sheet_subtitle)
        sub_cell.font = SUBTITLE_FONT

        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 16

        # -- Header (baris ke-3 setelah insert) --
        header_row = 3
        for col_idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.value = COL_LABELS.get(col_name, col_name)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER_WRAP
            cell.border = THIN_BORDER
        ws.row_dimensions[header_row].height = 30

        # -- Data rows --
        status_col_idx = None
        if "STATUS_VALIDASI" in cols:
            status_col_idx = cols.index("STATUS_VALIDASI") + 1

        for row_idx in range(header_row + 1, header_row + 1 + len(df_sheet)):
            df_row_idx = row_idx - header_row - 1
            status_val = ""
            if status_col_idx and df_row_idx < len(df_sheet):
                status_val = str(df_sheet.iloc[df_row_idx].get("STATUS_VALIDASI", ""))

            # Warna baris berdasarkan status
            if status_val == "ACC":
                row_fill, status_font = ACC_FILL, ACC_FONT
            elif status_val == "ACC (Cek Berkas Fisik)":
                row_fill, status_font = CEK_FILL, CEK_FONT
            elif status_val == "PERLU REVISI":
                row_fill, status_font = REV_FILL, REV_FONT
            elif status_val == "Diluar Modul":
                row_fill, status_font = DILUAR_FILL, DILUAR_FONT
            else:
                row_fill, status_font = None, NORMAL_FONT

            for col_idx in range(1, len(cols) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = THIN_BORDER
                cell.font = NORMAL_FONT
                cell.alignment = WRAP_ALIGN

                if row_fill:
                    cell.fill = row_fill
                # Font khusus untuk kolom status
                if col_idx == status_col_idx:
                    cell.font = status_font
                    cell.alignment = CENTER_WRAP

            # Tinggi baris otomatis (estimasi)
            ws.row_dimensions[row_idx].height = 32

        # -- Lebar kolom --
        for col_idx, col_name in enumerate(cols, 1):
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = COL_WIDTHS.get(col_name, 15)

        # -- Freeze panes di bawah header --
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)

        # -- Auto-filter --
        last_col_letter = get_column_letter(len(cols))
        last_row = header_row + len(df_sheet)
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{last_row}"

    # --- Statistik untuk subtitle ---
    total = len(df_validated)
    n_ri_xl = len(df_validated[df_validated.get("JENIS_RAWAT", pd.Series(dtype=str)) == "Rawat Inap"]) if "JENIS_RAWAT" in df_validated.columns else 0
    n_rj_xl = len(df_validated[df_validated.get("JENIS_RAWAT", pd.Series(dtype=str)) == "Rawat Jalan"]) if "JENIS_RAWAT" in df_validated.columns else 0
    n_acc = len(df_validated[df_validated["STATUS_VALIDASI"] == "ACC"])
    n_cek = len(df_validated[df_validated["STATUS_VALIDASI"] == "ACC (Cek Berkas Fisik)"])
    n_rev = len(df_validated[df_validated["STATUS_VALIDASI"] == "PERLU REVISI"])
    n_diluar_xl = len(df_validated[df_validated["STATUS_VALIDASI"] == "Diluar Modul"])
    tgl = datetime.now().strftime("%d-%m-%Y %H:%M")
    summary = (
        f"Tanggal: {tgl}  |  Total: {total} SEP  |  "
        f"RI: {n_ri_xl}  |  RJ: {n_rj_xl}  |  "
        f"ACC: {n_acc}  |  Cek Berkas: {n_cek}  |  Revisi: {n_rev}  |  "
        f"Diluar Modul: {n_diluar_xl}"
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # === Sheet 1: Semua SEP ===
        df_validated[display_cols].to_excel(
            writer, sheet_name="Semua SEP", index=False
        )
        _style_sheet(
            writer.sheets["Semua SEP"], df_validated, display_cols,
            "Hasil Validasi Kelengkapan Berkas & Coding BPJS",
            summary,
        )

        # === Sheet 2: Perlu Revisi ===
        df_revisi = df_validated[df_validated["STATUS_VALIDASI"] == "PERLU REVISI"]
        if not df_revisi.empty:
            df_revisi[display_cols].to_excel(
                writer, sheet_name="Perlu Revisi", index=False
            )
            _style_sheet(
                writer.sheets["Perlu Revisi"], df_revisi, display_cols,
                "Daftar SEP — PERLU REVISI",
                f"Total: {len(df_revisi)} SEP  |  Tanggal: {tgl}",
            )

        # === Sheet 3: Cek Berkas Fisik ===
        df_cek = df_validated[df_validated["STATUS_VALIDASI"] == "ACC (Cek Berkas Fisik)"]
        if not df_cek.empty:
            df_cek[display_cols].to_excel(
                writer, sheet_name="Cek Berkas Fisik", index=False
            )
            _style_sheet(
                writer.sheets["Cek Berkas Fisik"], df_cek, display_cols,
                "Daftar SEP — ACC (Cek Berkas Fisik)",
                f"Total: {len(df_cek)} SEP  |  Tanggal: {tgl}",
            )

        # === Sheet 4: ACC ===
        df_acc = df_validated[df_validated["STATUS_VALIDASI"] == "ACC"]
        if not df_acc.empty:
            df_acc[display_cols].to_excel(
                writer, sheet_name="ACC", index=False
            )
            _style_sheet(
                writer.sheets["ACC"], df_acc, display_cols,
                "Daftar SEP — ACC (Sesuai)",
                f"Total: {len(df_acc)} SEP  |  Tanggal: {tgl}",
            )

        # === Sheet 5: Diluar Modul ===
        df_diluar = df_validated[df_validated["STATUS_VALIDASI"] == "Diluar Modul"]
        if not df_diluar.empty:
            df_diluar[display_cols].to_excel(
                writer, sheet_name="Diluar Modul", index=False
            )
            _style_sheet(
                writer.sheets["Diluar Modul"], df_diluar, display_cols,
                "Daftar SEP — Diluar Modul (Diagnosis Tidak Tercakup)",
                f"Total: {len(df_diluar)} SEP  |  Tanggal: {tgl}",
            )

        # === Sheet 6: Riwayat Perubahan ===
        history = st.session_state.get("change_history", [])
        if history:
            df_history = pd.DataFrame(history).reindex(
                columns=["sep", "status_lama", "status_baru", "tujuan", "catatan",
                         "verifikator", "waktu"]
            ).fillna("")
            df_history.columns = [
                "SEP", "Status Lama", "Status Baru", "Tujuan",
                "Catatan", "Verifikator", "Waktu",
            ]
            hist_cols = list(df_history.columns)
            df_history.to_excel(
                writer, sheet_name="Riwayat Perubahan", index=False
            )
            ws_hist = writer.sheets["Riwayat Perubahan"]

            # Styling riwayat (lebih sederhana, tanpa row coloring)
            ws_hist.insert_rows(1, 2)
            ws_hist.merge_cells(
                start_row=1, start_column=1, end_row=1, end_column=len(hist_cols)
            )
            ws_hist.cell(row=1, column=1, value="Riwayat Perubahan Status").font = TITLE_FONT
            ws_hist.merge_cells(
                start_row=2, start_column=1, end_row=2, end_column=len(hist_cols)
            )
            ws_hist.cell(
                row=2, column=1,
                value=f"Total: {len(history)} perubahan  |  Tanggal cetak: {tgl}",
            ).font = SUBTITLE_FONT

            HIST_WIDTHS = {"SEP": 26, "Status Lama": 22, "Status Baru": 22,
                           "Tujuan": 18, "Catatan": 40, "Verifikator": 18, "Waktu": 20}
            header_row = 3
            for ci, cn in enumerate(hist_cols, 1):
                cell = ws_hist.cell(row=header_row, column=ci)
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = CENTER_WRAP
                cell.border = THIN_BORDER
                letter = get_column_letter(ci)
                ws_hist.column_dimensions[letter].width = HIST_WIDTHS.get(cn, 15)

            for ri in range(header_row + 1, header_row + 1 + len(df_history)):
                for ci in range(1, len(hist_cols) + 1):
                    cell = ws_hist.cell(row=ri, column=ci)
                    cell.border = THIN_BORDER
                    cell.font = NORMAL_FONT
                    cell.alignment = WRAP_ALIGN

            ws_hist.freeze_panes = ws_hist.cell(row=header_row + 1, column=1)

    return output.getvalue()


@st.cache_data(show_spinner=False)
def get_pending_sheet_info(file_bytes: bytes) -> dict[str, list[str]]:
    """Baca daftar sheet dan kolom masing-masing sheet dari Excel data pending."""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    info: dict[str, list[str]] = {}
    for sheet in xl.sheet_names:
        cols = [str(c).strip().upper() for c in pd.read_excel(xl, sheet, nrows=0).columns]
        info[sheet] = cols
    return info


@st.cache_data(show_spinner=False)
def load_pending_excel(file_bytes: bytes, selected_sheets: tuple[str, ...]) -> pd.DataFrame:
    """Muat dan gabungkan sheet terpilih dari Excel data pending berdasarkan NOSEP."""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    frames: list[pd.DataFrame] = []

    for sheet in selected_sheets:
        df_s = pd.read_excel(xl, sheet, dtype=str)
        df_s.columns = [str(c).strip().upper() for c in df_s.columns]

        # Cari kolom SEP (NOSEP, SEP, atau NOSJP)
        sep_col = None
        for cand in ("NOSEP", "SEP", "NOSJP"):
            if cand in df_s.columns:
                sep_col = cand
                break
        if sep_col is None:
            continue

        out = pd.DataFrame()
        out["NOSEP"] = df_s[sep_col].fillna("").astype(str).str.strip().str.upper()
        out = out[out["NOSEP"] != ""].copy()

        for target_col, fallbacks in [
            ("FILTRASI", ("FILTRASI",)),
            ("JNSPENDING", ("JNSPENDING", "KET PRE OPS")),
            ("ALASANPENDING", ("ALASANPENDING", "ALASAN_TIDAK_LAYAK")),
        ]:
            val_series = pd.Series("", index=out.index)
            for fb in fallbacks:
                if fb in df_s.columns:
                    val_series = df_s.loc[out.index, fb].fillna("").astype(str).str.strip()
                    break
            out[target_col] = val_series.replace({"nan": "", "None": ""})

        frames.append(out)

    if not frames:
        return pd.DataFrame(columns=["NOSEP", "FILTRASI", "JNSPENDING", "ALASANPENDING"])

    combined = pd.concat(frames, ignore_index=True)

    # Jika ada NOSEP ganda, gabungkan nilai unik yang tidak kosong
    def _join_unique(series: pd.Series) -> str:
        vals = []
        for v in series:
            v_str = str(v).strip()
            if v_str and v_str not in vals:
                vals.append(v_str)
        return " | ".join(vals)

    grouped = (
        combined.groupby("NOSEP", as_index=False)
        .agg({
            "FILTRASI": _join_unique,
            "JNSPENDING": _join_unique,
            "ALASANPENDING": _join_unique,
        })
    )
    return grouped


def match_pending_data(
    df_validated: pd.DataFrame,
    df_pending: pd.DataFrame,
    pending_notes: dict[str, str],
) -> pd.DataFrame:
    """Cocokkan data tervalidasi dengan data pending berdasarkan No. SEP."""
    if df_validated.empty or df_pending.empty:
        return pd.DataFrame()

    df_left = df_validated.copy()
    df_left["_SEP_KEY"] = df_left["SEP"].astype(str).str.strip().str.upper()

    merged = df_left.merge(
        df_pending[["NOSEP", "FILTRASI", "JNSPENDING", "ALASANPENDING"]],
        left_on="_SEP_KEY",
        right_on="NOSEP",
        how="inner",
    )
    merged.drop(columns=["_SEP_KEY", "NOSEP"], inplace=True)

    # Tambahkan kolom KETERANGAN dari session_state
    merged["KETERANGAN"] = merged["SEP"].astype(str).str.strip().str.upper().map(
        lambda s: pending_notes.get(s, "")
    )
    return merged


def build_pending_excel_report(df_pending_matched: pd.DataFrame) -> bytes:
    """Buat file Excel khusus data yang termasuk dalam data pending (+ 4 kolom pending)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    output = io.BytesIO()
    priority_cols = [
        "SEP", "NAMA_PASIEN", "MRN", "DPJP", "JENIS_RAWAT", "INACBG",
        "DIAGLIST", "PROCLIST", "STATUS_VALIDASI", "KETERANGAN_REVISI",
        "CEK_BERKAS_FISIK", "FILTRASI", "JNSPENDING", "ALASANPENDING", "KETERANGAN",
    ]
    export_cols = [c for c in priority_cols if c in df_pending_matched.columns]

    HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
    PENDING_HEADER_FILL = PatternFill("solid", fgColor="833C0C")  # cokelat/oranye tua untuk 4 kolom pending
    HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    ACC_FILL = PatternFill("solid", fgColor="D9F2D9")
    CEK_FILL = PatternFill("solid", fgColor="FFF2CC")
    REV_FILL = PatternFill("solid", fgColor="F4CCCC")
    DILUAR_FILL = PatternFill("solid", fgColor="E0E0E0")
    ACC_FONT = Font(name="Calibri", bold=True, color="1B7A1B")
    CEK_FONT = Font(name="Calibri", bold=True, color="B45F06")
    REV_FONT = Font(name="Calibri", bold=True, color="CC0000")
    DILUAR_FONT = Font(name="Calibri", bold=True, color="666666")
    NORMAL_FONT = Font(name="Calibri", size=10)
    TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
    SUBTITLE_FONT = Font(name="Calibri", italic=True, size=10, color="666666")
    THIN_BORDER = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
    CENTER_WRAP = Alignment(horizontal="center", vertical="top", wrap_text=True)

    COL_WIDTHS = {
        "SEP": 25, "NAMA_PASIEN": 22, "MRN": 14, "DPJP": 20,
        "JENIS_RAWAT": 14, "INACBG": 12, "DIAGLIST": 28, "PROCLIST": 24,
        "STATUS_VALIDASI": 22, "KETERANGAN_REVISI": 38, "CEK_BERKAS_FISIK": 34,
        "FILTRASI": 24, "JNSPENDING": 36, "ALASANPENDING": 42, "KETERANGAN": 36,
    }
    COL_LABELS = {
        "SEP": "No. SEP", "NAMA_PASIEN": "Nama Pasien", "MRN": "No. RM",
        "DPJP": "DPJP", "JENIS_RAWAT": "Jenis Rawat", "INACBG": "Kode INA-CBG",
        "DIAGLIST": "Daftar Diagnosis (ICD-10)", "PROCLIST": "Daftar Prosedur (ICD-9-CM)",
        "STATUS_VALIDASI": "Status Validasi Modul", "KETERANGAN_REVISI": "Keterangan Revisi Modul",
        "CEK_BERKAS_FISIK": "Cek Berkas Fisik",
        "FILTRASI": "FILTRASI", "JNSPENDING": "JNSPENDING",
        "ALASANPENDING": "ALASANPENDING", "KETERANGAN": "KETERANGAN",
    }
    PENDING_COLS = {"FILTRASI", "JNSPENDING", "ALASANPENDING", "KETERANGAN"}

    def _style_pending_sheet(ws, df_sheet, cols, title, subtitle):
        ws.insert_rows(1, 2)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
        ws.cell(row=1, column=1, value=title).font = TITLE_FONT
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(cols))
        ws.cell(row=2, column=1, value=subtitle).font = SUBTITLE_FONT
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 16

        header_row = 3
        for ci, cn in enumerate(cols, 1):
            cell = ws.cell(row=header_row, column=ci)
            cell.value = COL_LABELS.get(cn, cn)
            cell.fill = PENDING_HEADER_FILL if cn in PENDING_COLS else HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER_WRAP
            cell.border = THIN_BORDER
            ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(cn, 16)
        ws.row_dimensions[header_row].height = 30

        status_col_idx = cols.index("STATUS_VALIDASI") + 1 if "STATUS_VALIDASI" in cols else None
        for ri in range(header_row + 1, header_row + 1 + len(df_sheet)):
            df_ri = ri - header_row - 1
            st_val = str(df_sheet.iloc[df_ri].get("STATUS_VALIDASI", ""))
            if st_val == "ACC":
                r_fill, st_font = ACC_FILL, ACC_FONT
            elif st_val == "ACC (Cek Berkas Fisik)":
                r_fill, st_font = CEK_FILL, CEK_FONT
            elif st_val == "PERLU REVISI":
                r_fill, st_font = REV_FILL, REV_FONT
            elif st_val == "Diluar Modul":
                r_fill, st_font = DILUAR_FILL, DILUAR_FONT
            else:
                r_fill, st_font = None, NORMAL_FONT

            for ci in range(1, len(cols) + 1):
                cell = ws.cell(row=ri, column=ci)
                cell.border = THIN_BORDER
                cell.font = st_font if ci == status_col_idx else NORMAL_FONT
                cell.alignment = CENTER_WRAP if ci == status_col_idx else WRAP_ALIGN
                if r_fill:
                    cell.fill = r_fill
            ws.row_dimensions[ri].height = 36

        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
        last_col_letter = get_column_letter(len(cols))
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row + len(df_sheet)}"

    tgl = datetime.now().strftime("%d-%m-%Y %H:%M")
    total_p = len(df_pending_matched)
    n_ri_p = len(df_pending_matched[df_pending_matched["JENIS_RAWAT"] == "Rawat Inap"]) if "JENIS_RAWAT" in df_pending_matched.columns else 0
    n_rj_p = len(df_pending_matched[df_pending_matched["JENIS_RAWAT"] == "Rawat Jalan"]) if "JENIS_RAWAT" in df_pending_matched.columns else 0
    n_ket = int((df_pending_matched["KETERANGAN"].astype(str).str.strip() != "").sum()) if "KETERANGAN" in df_pending_matched.columns else 0
    sub_all = f"Tanggal: {tgl}  |  Total Cocok Pending: {total_p} SEP  |  RI: {n_ri_p}  |  RJ: {n_rj_p}  |  Keterangan Diisi: {n_ket}"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_pending_matched[export_cols].to_excel(writer, sheet_name="Data Pending", index=False)
        _style_pending_sheet(
            writer.sheets["Data Pending"], df_pending_matched, export_cols,
            "Hasil Pengecekan Data Pending BPJS", sub_all,
        )

        if "JENIS_RAWAT" in df_pending_matched.columns:
            for jenis, s_name in [("Rawat Inap", "Pending Rawat Inap"), ("Rawat Jalan", "Pending Rawat Jalan")]:
                df_sub = df_pending_matched[df_pending_matched["JENIS_RAWAT"] == jenis]
                if not df_sub.empty:
                    df_sub[export_cols].to_excel(writer, sheet_name=s_name, index=False)
                    _style_pending_sheet(
                        writer.sheets[s_name], df_sub, export_cols,
                        f"Daftar {s_name}",
                        f"Total: {len(df_sub)} SEP  |  Tanggal: {tgl}",
                    )

    return output.getvalue()


def flatten_rules_to_table(rules_data: dict) -> pd.DataFrame:
    """Ubah rules.json menjadi tabel datar untuk ditampilkan."""
    rows = []
    for cat in rules_data.get("categories", []):
        cat_name = cat.get("name", "")
        cat_id = cat.get("id", "")
        for rule in cat.get("rules", []):
            rows.append({
                "Kategori": cat_name,
                "Kategori_ID": cat_id,
                "Rule_ID": rule.get("id", ""),
                "Nama": rule.get("name", ""),
                "Deskripsi": rule.get("description", ""),
                "Tipe": rule.get("type", "standard"),
                "Trigger Diagnosis": ", ".join(rule.get("trigger_diagnoses", rule.get("trigger_group_a", []))),
                "Required Procedures": json.dumps(rule.get("required_procedures", {}), ensure_ascii=False) if rule.get("required_procedures") else "",
                "Pesan Revisi": rule.get("revision_message", ""),
                "Cek Berkas Fisik": rule.get("physical_document_check", ""),
            })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# DIALOG DETAIL SEP (dengan fitur Edit Status)
# ---------------------------------------------------------------------------

@st.dialog("✏️ Konfirmasi Perubahan Status", width="medium")
def show_inline_status_dialog(changed_rows: list[dict], changes_sig: tuple):
    """Modal popup otomatis saat user mengubah status pada tabel validasi."""
    st.session_state["dismissed_inline_sig"] = changes_sig
    st.info(f"**{len(changed_rows)}** SEP diubah statusnya. Lengkapi keterangan di bawah untuk menyimpan.")

    for ch in changed_rows:
        e_lama = STATUS_EMOJI_MAP.get(ch["status_lama"], "")
        e_baru = STATUS_EMOJI_MAP.get(ch["status_baru"], "")
        st.write(f"• **{ch['sep']}**: {e_lama} {ch['status_lama']} → {e_baru} {ch['status_baru']}")

    st.divider()
    with st.form("inline_dialog_form", clear_on_submit=False):
        tujuan = st.selectbox(
            "Tujuan Perubahan",
            options=TUJUAN_OPTIONS,
            index=0,
            help="Pilih tujuan tindak lanjut (Coder, Admin Ruangan, atau Petugas Entry).",
        )
        catatan = st.text_area(
            "Catatan / Alasan Perubahan *",
            placeholder="Contoh: Berkas fisik sudah lengkap, Koding sudah diperbaiki, dll.",
        )
        verifikator = st.text_input(
            "Nama Verifikator *",
            value=st.session_state.get("last_verifikator", ""),
            placeholder="Nama petugas verifikator",
        )

        c1, c2 = st.columns(2)
        with c1:
            save_clicked = st.form_submit_button(
                f"💾 Simpan {len(changed_rows)} Perubahan",
                type="primary",
                use_container_width=True,
            )
        with c2:
            cancel_clicked = st.form_submit_button(
                "❌ Batalkan Perubahan",
                type="secondary",
                use_container_width=True,
            )

        if save_clicked:
            if not catatan.strip():
                st.error("❌ Catatan/alasan perubahan wajib diisi!")
            elif not verifikator.strip():
                st.error("❌ Nama verifikator wajib diisi!")
            else:
                for ch in changed_rows:
                    update_status(
                        ch["sep"],
                        ch["status_lama"],
                        ch["status_baru"],
                        catatan.strip(),
                        verifikator.strip(),
                        tujuan,
                    )
                st.session_state["last_verifikator"] = verifikator.strip()
                st.session_state["inline_editor_rev"] += 1
                st.session_state.pop("dismissed_inline_sig", None)
                st.session_state["flash_message"] = (
                    f"✅ Berhasil menyimpan perubahan status **{len(changed_rows)} SEP**"
                    + (f" (Tujuan: **{tujuan}**)" if tujuan != "-" else "")
                    + "!"
                )
                st.rerun()

        if cancel_clicked:
            st.session_state["inline_editor_rev"] += 1
            st.session_state.pop("dismissed_inline_sig", None)
            st.rerun()


@st.dialog("Detail SEP", width="large")
def show_sep_detail(row: pd.Series, rules_data: dict):
    """Tampilkan modal detail SEP dengan form edit status."""
    sep_no = str(row.get("SEP", "N/A"))
    status = row.get("STATUS_VALIDASI", "")
    emoji = row.get("STATUS_EMOJI", "")

    # Cek apakah ada override
    override = st.session_state["status_overrides"].get(sep_no)
    is_overridden = override is not None

    st.subheader(f"{emoji} SEP: {sep_no} — {status}")

    if is_overridden:
        tujuan_ov = override.get("tujuan") or "-"
        st.info(
            f"ℹ️ Status telah diubah manual oleh **{override.get('verifikator', '-')}** "
            f"— Tujuan: **{tujuan_ov}** — Catatan: _{override.get('catatan', '-')}_"
        )

    if status == "ACC":
        st.success(f"**{sep_no}** — Semua kode sesuai aturan modul ✅")
    elif status == "ACC (Cek Berkas Fisik)":
        st.warning(f"**{sep_no}** — Kode sesuai, namun ada berkas fisik yang harus diverifikasi ⚠️")
    elif status == "Diluar Modul":
        st.info(f"**{sep_no}** — Diagnosis tidak tercakup dalam modul aturan ⚪")
    else:
        st.error(f"**{sep_no}** — Ditemukan pelanggaran aturan, perlu revisi ❌")

    st.divider()

    # Identitas Pasien
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Identitas Pasien**")
        st.write(f"**Nama:** {row.get('NAMA_PASIEN', '-')}")
        st.write(f"**MRN:** {row.get('MRN', '-')}")
        st.write(f"**DPJP:** {row.get('DPJP', '-')}")
    with col2:
        st.markdown("**Rawat Inap**")
        st.write(f"**Tgl Masuk:** {row.get('ADMISSION_DATE', '-')}")
        st.write(f"**Tgl Keluar:** {row.get('DISCHARGE_DATE', '-')}")
        st.write(f"**LOS:** {row.get('LOS', '-')} hari")

    st.divider()

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**INA-CBG**")
        st.write(f"**Kode:** {row.get('INACBG', '-')}")
        st.write(f"**Deskripsi:** {row.get('DESKRIPSI_INACBG', '-')}")
    with col4:
        st.markdown("**Ventilator**")
        vent = row.get("VENT_HOUR", "-")
        st.write(f"**Durasi:** {vent} jam")

    st.divider()

    # Daftar Diagnosis
    st.markdown("**📋 Daftar Diagnosis (DIAGLIST)**")
    diag_raw = str(row.get("DIAGLIST", ""))
    diags = [d.strip() for d in diag_raw.replace(",", ";").split(";") if d.strip()]
    if diags:
        for i, d in enumerate(diags):
            label = "🔹 Diagnosis Utama" if i == 0 else f"🔸 Diagnosis Sekunder #{i}"
            st.write(f"  {label}: `{d}`")
    else:
        st.write("  _(Tidak ada data)_")

    # Daftar Prosedur
    st.markdown("**🔧 Daftar Prosedur (PROCLIST)**")
    proc_raw = str(row.get("PROCLIST", ""))
    procs = [p.strip() for p in proc_raw.replace(",", ";").split(";") if p.strip()]
    if procs:
        for p in procs:
            st.write(f"  • `{p}`")
    else:
        st.write("  _(Tidak ada data)_")

    st.divider()

    # Hasil Validasi
    st.markdown("**📝 Hasil Validasi**")

    revisi = row.get("KETERANGAN_REVISI", "")
    if revisi:
        st.error("**Keterangan Revisi:**")
        for line in revisi.split("\n"):
            if line.strip():
                st.write(f"  ❌ {line.strip()}")

    berkas = row.get("CEK_BERKAS_FISIK", "")
    if berkas:
        st.warning("**Syarat Berkas Fisik yang Harus Diverifikasi:**")
        for line in berkas.split("\n"):
            if line.strip():
                st.write(f"  📄 {line.strip()}")

    aturan = row.get("ATURAN_TERPENUHI", "")
    if aturan:
        st.success("**Aturan Terpenuhi:**")
        for a in aturan.split(";"):
            if a.strip():
                st.write(f"  ✅ {a.strip()}")

    if not revisi and not berkas and not aturan:
        st.info("Tidak ada aturan modul yang berlaku untuk diagnosis/prosedur pada SEP ini.")

    # --- Riwayat perubahan SEP ini ---
    sep_history = [h for h in st.session_state["change_history"] if h["sep"] == sep_no]
    if sep_history:
        st.divider()
        st.markdown("**📜 Riwayat Perubahan Status SEP Ini**")
        for h in reversed(sep_history):
            e_lama = STATUS_EMOJI_MAP.get(h["status_lama"], "")
            e_baru = STATUS_EMOJI_MAP.get(h["status_baru"], "")
            tujuan_h = h.get("tujuan") or "-"
            st.write(
                f"  🕐 `{h['waktu']}` — {e_lama} {h['status_lama']} → {e_baru} {h['status_baru']} "
                f"oleh **{h['verifikator']}** (Tujuan: **{tujuan_h}**) — _{h['catatan']}_"
            )

    # --- Form Edit Status ---
    st.divider()
    st.markdown("### ✏️ Ubah Status SEP Ini")

    current_idx = STATUS_OPTIONS.index(status) if status in STATUS_OPTIONS else 0
    cur_tujuan = (override or {}).get("tujuan") or "-"
    tujuan_idx = TUJUAN_OPTIONS.index(cur_tujuan) if cur_tujuan in TUJUAN_OPTIONS else 0

    with st.form(f"edit_status_{sep_no}", clear_on_submit=False):
        new_status = st.selectbox(
            "Status Baru",
            options=STATUS_OPTIONS,
            index=current_idx,
            key=f"new_status_{sep_no}",
        )
        tujuan = st.selectbox(
            "Tujuan Perubahan",
            options=TUJUAN_OPTIONS,
            index=tujuan_idx,
            key=f"tujuan_{sep_no}",
        )
        catatan = st.text_area(
            "Catatan / Alasan Perubahan *",
            placeholder="Contoh: Berkas fisik sudah diverifikasi lengkap, Koding sudah diperbaiki oleh coder, dll.",
            key=f"catatan_{sep_no}",
        )
        verifikator = st.text_input(
            "Nama Verifikator *",
            value=st.session_state.get("last_verifikator", ""),
            placeholder="Nama petugas verifikator",
            key=f"verifikator_{sep_no}",
        )

        submitted = st.form_submit_button("💾 Simpan Perubahan Status", type="primary")

        if submitted:
            if not catatan.strip():
                st.error("❌ Catatan/alasan perubahan wajib diisi!")
            elif not verifikator.strip():
                st.error("❌ Nama verifikator wajib diisi!")
            elif new_status == status and not is_overridden:
                st.warning("⚠️ Status tidak berubah.")
            else:
                update_status(sep_no, status, new_status, catatan.strip(), verifikator.strip(), tujuan)
                st.session_state["last_verifikator"] = verifikator.strip()
                st.success(f"✅ Status SEP **{sep_no}** berhasil diubah menjadi **{new_status}**")
                st.rerun()


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------

st.sidebar.image("https://img.icons8.com/color/96/hospital-3.png", width=64)
st.sidebar.title("🏥 Validator BPJS")
st.sidebar.caption("Kelengkapan Berkas & Coding Rawat Inap & Rawat Jalan")
st.sidebar.divider()

uploaded_file = st.sidebar.file_uploader(
    "📂 Upload File Data Klaim (MIX)",
    type=["xls", "xlsx", "tsv", "txt", "csv"],
    help="File ekspor SIMRS/INA-CBG. Mendukung .xls (TSV), .xlsx, .tsv, .txt, .csv",
)

st.sidebar.divider()

# Filter (muncul setelah file diupload)
filter_status = st.sidebar.selectbox(
    "🔍 Filter Status Validasi",
    ["Semua", "PERLU REVISI", "ACC (Cek Berkas Fisik)", "ACC", "Diluar Modul"],
    index=0,
)

search_query = st.sidebar.text_input(
    "🔎 Cari SEP / MRN / Nama Pasien",
    placeholder="Ketik nomor SEP, MRN, atau nama...",
)

filter_dpjp = st.sidebar.text_input(
    "👨‍⚕️ Filter DPJP",
    placeholder="Nama DPJP...",
)

filter_diag = st.sidebar.text_input(
    "🏷️ Filter Kode Diagnosis",
    placeholder="Contoh: I50.9, A09.9...",
)

filter_jenis_rawat = st.sidebar.selectbox(
    "🏨 Filter Jenis Rawat",
    ["Semua", "Rawat Inap", "Rawat Jalan"],
    index=0,
)

# Sidebar: ringkasan override
overrides = st.session_state.get("status_overrides", {})
if overrides:
    st.sidebar.divider()
    st.sidebar.markdown(f"✏️ **Status Diedit:** {len(overrides)} SEP")
    if st.sidebar.button("🔄 Reset Semua Edit Status"):
        st.session_state["status_overrides"] = {}
        st.session_state["change_history"] = []
        st.rerun()

# ---------------------------------------------------------------------------
# TABS UTAMA
# ---------------------------------------------------------------------------

tab_validasi, tab_pending, tab_edit, tab_riwayat, tab_aturan = st.tabs([
    "📊 Validasi Klaim",
    "🔍 Cek Data Pending",
    "✏️ Edit Status (Bulk)",
    "📜 Riwayat Perubahan",
    "⚙️ Kelola Aturan Modul",
])

# ---------------------------------------------------------------------------
# TAB 1: VALIDASI KLAIM
# ---------------------------------------------------------------------------

with tab_validasi:
    st.title("📊 Validator Kelengkapan Berkas & Coding BPJS")
    st.caption("Modul Kelengkapan Berkas Rekam Medis Rawat Inap & Rawat Jalan Berdasarkan Diagnosis")

    if uploaded_file is None:
        st.info(
            "👆 Silakan upload file data klaim (MIX) di sidebar untuk memulai validasi.\n\n"
            "**Format yang didukung:** `.xls` (TSV), `.xlsx`, `.tsv`, `.txt`, `.csv`\n\n"
            "**Kolom kunci yang diharapkan:** `SEP`, `PTD`, `DIAGLIST`, `PROCLIST`, `VENT_HOUR`"
        )
        st.stop()

    # --- Muat dan Validasi Data ---
    with st.spinner("📂 Membaca file..."):
        try:
            df_raw = cached_load_file(uploaded_file.getvalue(), uploaded_file.name)
        except Exception as e:
            st.error(f"❌ Gagal membaca file: {e}")
            st.stop()

    # Normalisasi nama kolom
    df_raw.columns = [c.strip().upper() for c in df_raw.columns]

    # Cek kolom kunci
    required_cols = ["SEP", "DIAGLIST"]
    missing_cols = [c for c in required_cols if c not in df_raw.columns]
    if missing_cols:
        st.error(f"❌ Kolom kunci tidak ditemukan: **{', '.join(missing_cols)}**")
        st.write("Kolom yang terdeteksi pada file:")
        st.code(", ".join(df_raw.columns.tolist()))
        st.stop()

    # Tambahkan kolom JENIS_RAWAT berdasarkan PTD
    df_work = df_raw.copy()
    if "PTD" in df_work.columns:
        df_work["JENIS_RAWAT"] = df_work["PTD"].astype(str).str.strip().map(
            {"1": "Rawat Inap", "2": "Rawat Jalan"}
        ).fillna("Lainnya")
    else:
        df_work["JENIS_RAWAT"] = "Tidak Diketahui"

    # Pastikan kolom PROCLIST dan VENT_HOUR ada
    if "PROCLIST" not in df_work.columns:
        df_work["PROCLIST"] = ""
    if "VENT_HOUR" not in df_work.columns:
        df_work["VENT_HOUR"] = ""

    # Validasi otomatis (di-cache berdasarkan data & rules)
    rules_data = get_rules()
    rules_json_str = json.dumps(rules_data, sort_keys=True)
    with st.spinner("🔍 Menjalankan validasi aturan..."):
        df_validated = cached_validate(df_work, rules_json_str)

    # Terapkan override status manual
    df_validated = apply_status_overrides(df_validated)

    # --- KPI Metrics ---
    st.divider()
    total = len(df_validated)
    n_ri = len(df_validated[df_validated["JENIS_RAWAT"] == "Rawat Inap"])
    n_rj = len(df_validated[df_validated["JENIS_RAWAT"] == "Rawat Jalan"])
    n_acc = len(df_validated[df_validated["STATUS_VALIDASI"] == "ACC"])
    n_cek = len(df_validated[df_validated["STATUS_VALIDASI"] == "ACC (Cek Berkas Fisik)"])
    n_rev = len(df_validated[df_validated["STATUS_VALIDASI"] == "PERLU REVISI"])
    n_diluar = len(df_validated[df_validated["STATUS_VALIDASI"] == "Diluar Modul"])
    n_edited = len([s for s in df_validated["SEP"].astype(str) if s in overrides])

    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    m1.metric("📋 Total SEP", total)
    m2.metric("🏥 Rawat Inap", n_ri)
    m3.metric("🏨 Rawat Jalan", n_rj)
    m4.metric("🟢 ACC", n_acc)
    m5.metric("🟡 Cek Berkas", n_cek)
    m6.metric("🔴 Perlu Revisi", n_rev)
    m7.metric("⚪ Diluar Modul", n_diluar)

    # Progress bar visual (4 kolom sama lebar agar teks tidak terpotong)
    if total > 0:
        col_bar = st.columns(4)
        bar_items = [
            (col_bar[0], "🟢 ACC", n_acc),
            (col_bar[1], "🟡 Cek Berkas", n_cek),
            (col_bar[2], "🔴 Perlu Revisi", n_rev),
            (col_bar[3], "⚪ Diluar Modul", n_diluar),
        ]
        for col, label, count in bar_items:
            pct = (count / total) * 100
            with col:
                st.progress(count / total, text=f"{label}: {count:,}/{total:,} ({pct:.1f}%)")

    st.divider()

    # --- Filter & Pencarian ---
    df_display = df_validated.copy()

    # Filter jenis rawat dari sidebar
    if filter_jenis_rawat != "Semua":
        df_display = df_display[df_display["JENIS_RAWAT"] == filter_jenis_rawat]

    if filter_status != "Semua":
        df_display = df_display[df_display["STATUS_VALIDASI"] == filter_status]

    if search_query:
        q = search_query.strip().upper()
        mask = pd.Series([False] * len(df_display), index=df_display.index)
        for col in ["SEP", "MRN", "NAMA_PASIEN"]:
            if col in df_display.columns:
                mask = mask | df_display[col].astype(str).str.upper().str.contains(q, na=False)
        df_display = df_display[mask]

    if filter_dpjp and "DPJP" in df_display.columns:
        df_display = df_display[
            df_display["DPJP"].astype(str).str.upper().str.contains(filter_dpjp.strip().upper(), na=False)
        ]

    if filter_diag:
        q_diag = filter_diag.strip().upper()
        df_display = df_display[
            df_display["DIAGLIST"].astype(str).str.upper().str.contains(q_diag, na=False)
        ]

    # --- Banner Pencarian SEP Spesifik ---
    if search_query and len(df_display) == 1:
        row = df_display.iloc[0]
        status = row.get("STATUS_VALIDASI", "")
        sep_no = str(row.get("SEP", ""))
        is_edited = sep_no in overrides
        edited_tag = " _(diedit manual)_" if is_edited else ""

        if status == "ACC":
            st.success(f"**SEP {sep_no} — ACC** ✅ Semua kode sesuai aturan modul.{edited_tag}")
            aturan = row.get("ATURAN_TERPENUHI", "")
            if aturan:
                for a in aturan.split(";"):
                    if a.strip():
                        st.write(f"  ✅ {a.strip()}")
        elif status == "ACC (Cek Berkas Fisik)":
            st.warning(f"**SEP {sep_no} — ACC (Cek Berkas Fisik)** ⚠️{edited_tag}")
            berkas = row.get("CEK_BERKAS_FISIK", "")
            if berkas:
                for line in berkas.split("\n"):
                    if line.strip():
                        st.write(f"  📄 {line.strip()}")
        elif status == "PERLU REVISI":
            st.error(f"**SEP {sep_no} — PERLU REVISI** ❌{edited_tag}")
            revisi = row.get("KETERANGAN_REVISI", "")
            if revisi:
                for line in revisi.split("\n"):
                    if line.strip():
                        st.write(f"  ❌ {line.strip()}")
        elif status == "Diluar Modul":
            st.info(f"**SEP {sep_no} — Diluar Modul** ⚪ Diagnosis tidak tercakup dalam modul aturan.{edited_tag}")
    elif search_query and len(df_display) == 0:
        st.warning(f"Tidak ditemukan SEP/MRN/Nama yang cocok dengan **\"{search_query}\"**.")

    # --- Tombol Filter Cepat Berdasarkan Jenis Rawat ---
    st.subheader(f"📋 Daftar SEP")

    # Tombol filter Jenis Rawat
    btn_ri = len(df_display[df_display["JENIS_RAWAT"] == "Rawat Inap"])
    btn_rj = len(df_display[df_display["JENIS_RAWAT"] == "Rawat Jalan"])
    btn_all_jenis = len(df_display)

    if "table_jenis_filter" not in st.session_state:
        st.session_state["table_jenis_filter"] = "Semua"

    jc1, jc2, jc3 = st.columns(3)
    with jc1:
        if st.button(
            f"🏥 Semua ({btn_all_jenis})",
            use_container_width=True,
            type="primary" if st.session_state["table_jenis_filter"] == "Semua" else "secondary",
            key="btn_jenis_semua",
        ):
            st.session_state["table_jenis_filter"] = "Semua"
            st.rerun()
    with jc2:
        if st.button(
            f"🛏️ Rawat Inap ({btn_ri})",
            use_container_width=True,
            type="primary" if st.session_state["table_jenis_filter"] == "Rawat Inap" else "secondary",
            key="btn_jenis_ri",
        ):
            st.session_state["table_jenis_filter"] = "Rawat Inap"
            st.rerun()
    with jc3:
        if st.button(
            f"🚶 Rawat Jalan ({btn_rj})",
            use_container_width=True,
            type="primary" if st.session_state["table_jenis_filter"] == "Rawat Jalan" else "secondary",
            key="btn_jenis_rj",
        ):
            st.session_state["table_jenis_filter"] = "Rawat Jalan"
            st.rerun()

    # Terapkan filter jenis rawat
    active_jenis_filter = st.session_state["table_jenis_filter"]
    if active_jenis_filter != "Semua":
        df_display = df_display[df_display["JENIS_RAWAT"] == active_jenis_filter]

    # --- Tombol Filter Cepat Berdasarkan Status ---
    # Hitung jumlah per status dari df_display (sudah difilter sidebar + jenis rawat)
    btn_n_all = len(df_display)
    btn_n_acc = len(df_display[df_display["STATUS_VALIDASI"] == "ACC"])
    btn_n_cek = len(df_display[df_display["STATUS_VALIDASI"] == "ACC (Cek Berkas Fisik)"])
    btn_n_rev = len(df_display[df_display["STATUS_VALIDASI"] == "PERLU REVISI"])
    btn_n_diluar = len(df_display[df_display["STATUS_VALIDASI"] == "Diluar Modul"])

    # Inisialisasi state tombol filter tabel
    if "table_status_filter" not in st.session_state:
        st.session_state["table_status_filter"] = "Semua"

    btn_col1, btn_col2, btn_col3, btn_col4, btn_col5 = st.columns(5)

    with btn_col1:
        if st.button(
            f"📋 Semua ({btn_n_all})",
            use_container_width=True,
            type="primary" if st.session_state["table_status_filter"] == "Semua" else "secondary",
            key="btn_filter_semua",
        ):
            st.session_state["table_status_filter"] = "Semua"
            st.rerun()
    with btn_col2:
        if st.button(
            f"🟢 ACC ({btn_n_acc})",
            use_container_width=True,
            type="primary" if st.session_state["table_status_filter"] == "ACC" else "secondary",
            key="btn_filter_acc",
        ):
            st.session_state["table_status_filter"] = "ACC"
            st.rerun()
    with btn_col3:
        if st.button(
            f"🟡 Cek Berkas ({btn_n_cek})",
            use_container_width=True,
            type="primary" if st.session_state["table_status_filter"] == "ACC (Cek Berkas Fisik)" else "secondary",
            key="btn_filter_cek",
        ):
            st.session_state["table_status_filter"] = "ACC (Cek Berkas Fisik)"
            st.rerun()
    with btn_col4:
        if st.button(
            f"🔴 Perlu Revisi ({btn_n_rev})",
            use_container_width=True,
            type="primary" if st.session_state["table_status_filter"] == "PERLU REVISI" else "secondary",
            key="btn_filter_revisi",
        ):
            st.session_state["table_status_filter"] = "PERLU REVISI"
            st.rerun()
    with btn_col5:
        if st.button(
            f"⚪ Diluar Modul ({btn_n_diluar})",
            use_container_width=True,
            type="primary" if st.session_state["table_status_filter"] == "Diluar Modul" else "secondary",
            key="btn_filter_diluar",
        ):
            st.session_state["table_status_filter"] = "Diluar Modul"
            st.rerun()

    # Terapkan filter tombol
    active_table_filter = st.session_state["table_status_filter"]
    if active_table_filter != "Semua":
        df_display = df_display[df_display["STATUS_VALIDASI"] == active_table_filter]

    # Paginasi tabel agar render browser tetap cepat untuk ribuan data
    total_filtered = len(df_display)
    pg_col1, pg_col2 = st.columns([1, 2])
    with pg_col1:
        page_size_opt = st.selectbox(
            "Baris per halaman:",
            [100, 250, 500, "Semua"],
            index=0,
            key="table_page_size",
        )

    if page_size_opt == "Semua" or total_filtered == 0:
        start_idx = 0
        df_page = df_display
        page_info = f"Menampilkan **{total_filtered}** SEP"
    else:
        page_size = int(page_size_opt)
        total_pages = max(1, (total_filtered + page_size - 1) // page_size)
        if st.session_state.get("table_page_num", 1) > total_pages:
            st.session_state["table_page_num"] = total_pages
        with pg_col2:
            current_page = st.number_input(
                f"Halaman (1 - {total_pages}):",
                min_value=1,
                max_value=total_pages,
                value=1,
                step=1,
                key="table_page_num",
            )
        start_idx = (current_page - 1) * page_size
        end_idx = min(start_idx + page_size, total_filtered)
        df_page = df_display.iloc[start_idx:end_idx]
        page_info = f"Menampilkan **{start_idx + 1}–{end_idx}** dari **{total_filtered}** SEP"

    st.caption(
        page_info
        + (f" — jenis: **{active_jenis_filter}**" if active_jenis_filter != "Semua" else "")
        + (f" — status: **{active_table_filter}**" if active_table_filter != "Semua" else "")
        + " | 💡 Klik kolom **Status ✏️** pada baris mana pun untuk mengubah status langsung"
    )

    table_cols = ["SEP", "STATUS_EMOJI", "STATUS_VALIDASI", "JENIS_RAWAT", "DIAGLIST", "PROCLIST",
                  "KETERANGAN_REVISI", "CEK_BERKAS_FISIK"]
    # Tambahkan kolom opsional jika ada
    for opt_col in ["NAMA_PASIEN", "DPJP", "LOS", "INACBG"]:
        if opt_col in df_page.columns:
            table_cols.insert(1, opt_col)

    available_cols = [c for c in table_cols if c in df_page.columns]

    # Siapkan DataFrame untuk editor — reset index agar stabil & tambahkan kolom nomor urut
    df_editable = df_page[available_cols].reset_index(drop=True).copy()
    df_editable.insert(0, "NO", range(start_idx + 1, start_idx + 1 + len(df_editable)))

    # Kolom yang TIDAK boleh diedit (semua kecuali STATUS_VALIDASI)
    disabled_cols = [c for c in df_editable.columns if c != "STATUS_VALIDASI"]

    _flash = st.session_state.pop("flash_message", None)
    if _flash:
        st.success(_flash)

    # Placeholder untuk bar aksi cepat tepat di atas tabel jika ada perubahan tertunda
    inline_bar_placeholder = st.container()

    editor_key = f"sep_table_editor_{st.session_state['inline_editor_rev']}"
    edited_df = st.data_editor(
        df_editable,
        use_container_width=True,
        height=400,
        num_rows="fixed",
        disabled=disabled_cols,
        column_config={
            "NO": st.column_config.NumberColumn("No.", width=60),
            "STATUS_EMOJI": st.column_config.TextColumn("", width=30),
            "STATUS_VALIDASI": st.column_config.SelectboxColumn(
                "Status ✏️",
                options=STATUS_OPTIONS,
                width=170,
                required=True,
            ),
            "SEP": st.column_config.TextColumn("SEP", width=180),
            "JENIS_RAWAT": st.column_config.TextColumn("Jenis Rawat", width=100),
            "KETERANGAN_REVISI": st.column_config.TextColumn("Keterangan Revisi", width=300),
            "CEK_BERKAS_FISIK": st.column_config.TextColumn("Cek Berkas Fisik", width=250),
        },
        hide_index=True,
        key=editor_key,
    )

    # --- Deteksi perubahan status -> buka modal popup otomatis (tanpa scroll ke bawah) ---
    changed_rows = []
    if not df_editable.empty:
        diff_mask = df_editable["STATUS_VALIDASI"] != edited_df["STATUS_VALIDASI"]
        for i in df_editable.index[diff_mask]:
            changed_rows.append({
                "sep": str(df_editable.at[i, "SEP"]),
                "status_lama": df_editable.at[i, "STATUS_VALIDASI"],
                "status_baru": edited_df.at[i, "STATUS_VALIDASI"],
            })

    if changed_rows:
        changes_sig = tuple((c["sep"], c["status_lama"], c["status_baru"]) for c in changed_rows)
        with inline_bar_placeholder:
            bar_c1, bar_c2, bar_c3 = st.columns([3, 1.3, 1])
            with bar_c1:
                st.warning(f"⚡ **{len(changed_rows)}** perubahan status belum disimpan.")
            with bar_c2:
                if st.button("📝 Isi Keterangan & Simpan", type="primary", use_container_width=True, key="reopen_inline_dialog_btn"):
                    st.session_state.pop("dismissed_inline_sig", None)
                    st.rerun()
            with bar_c3:
                if st.button("↩️ Batalkan", use_container_width=True, key="cancel_inline_dialog_btn"):
                    st.session_state["inline_editor_rev"] += 1
                    st.session_state.pop("dismissed_inline_sig", None)
                    st.rerun()

        if st.session_state.get("dismissed_inline_sig") != changes_sig:
            show_inline_status_dialog(changed_rows, changes_sig)

    st.divider()

    # --- Detail SEP (pilih dari selectbox) ---
    st.subheader("🔍 Lihat Detail & Edit Status SEP")

    sep_list = df_display["SEP"].astype(str).tolist()
    if sep_list:
        selected_sep = st.selectbox(
            "Pilih nomor SEP untuk melihat detail:",
            options=sep_list,
            index=0,
        )
        if st.button("📄 Lihat Detail & Edit Status", type="primary"):
            row_data = df_display[df_display["SEP"].astype(str) == selected_sep].iloc[0]
            show_sep_detail(row_data, rules_data)

    # --- Export (on-demand agar tidak memperlambat setiap rerun) ---
    st.divider()
    st.subheader("📥 Export Laporan")

    excel_sig = (
        uploaded_file.name,
        len(df_validated),
        len(st.session_state.get("change_history", [])),
    )
    excel_cache = st.session_state.get("excel_cache")

    if st.button("📊 Siapkan File Excel (.xlsx)", type="secondary"):
        with st.spinner("⏳ Membuat file Excel..."):
            st.session_state["excel_cache"] = {
                "sig": excel_sig,
                "bytes": build_excel_report(df_validated),
            }
        excel_cache = st.session_state["excel_cache"]

    if excel_cache and excel_cache.get("sig") == excel_sig:
        st.download_button(
            label="⬇️ Download Hasil Validasi (.xlsx)",
            data=excel_cache["bytes"],
            file_name="Hasil_Validasi_BPJS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.caption("💡 Klik **Siapkan File Excel (.xlsx)** di atas saat Anda siap mengunduh laporan.")

# ---------------------------------------------------------------------------
# TAB: CEK DATA PENDING
# ---------------------------------------------------------------------------

def _render_tab_pending(uploaded_file, df_validated):
    """Render isi tab Cek Data Pending tanpa st.stop() agar tab lain tidak terblokir."""
    st.title("🔍 Pengecekan Data Pending Pasien")
    st.caption(
        "Cocokkan hasil validasi modul dengan file Excel Data Pending berdasarkan No. SEP, "
        "lengkapi kolom KETERANGAN perbaikan, lalu download khusus data pending."
    )

    if uploaded_file is None or df_validated is None:
        st.info("👆 Upload file data klaim (MIX) terlebih dahulu di sidebar.")
        return

    pending_file = st.file_uploader(
        "📂 Upload File Excel Data Pending",
        type=["xlsx", "xls"],
        help="Contoh: Data Pending Agustus RSDK Reguler 2026.xlsx",
        key="pending_excel_uploader",
    )

    # Simpan ke session_state agar tidak hilang saat rerun / pindah tab / klik filter
    if pending_file is not None:
        new_bytes = pending_file.getvalue()
        if (
            st.session_state.get("pending_file_bytes") != new_bytes
            or st.session_state.get("pending_file_name") != pending_file.name
        ):
            st.session_state["pending_file_bytes"] = new_bytes
            st.session_state["pending_file_name"] = pending_file.name
            st.session_state.pop("pending_selected_sheets", None)

    pending_bytes = st.session_state.get("pending_file_bytes")
    pending_name = st.session_state.get("pending_file_name")

    if not pending_bytes:
        st.info(
            "👆 Upload file Excel Data Pending di atas untuk mencocokkan **No. SEP** "
            "dengan data yang sudah divalidasi.\n\n"
            "**Kolom yang akan ditambahkan:** `FILTRASI`, `JNSPENDING`, `ALASANPENDING`, dan `KETERANGAN` (dapat diisi langsung)."
        )
        return

    info_col, clr_col = st.columns([4, 1])
    with info_col:
        st.success(f"📄 File Data Pending aktif: **{pending_name}**")
    with clr_col:
        if st.button("🗑️ Hapus File", use_container_width=True, key="clear_pending_file_btn"):
            st.session_state["pending_file_bytes"] = None
            st.session_state["pending_file_name"] = None
            st.session_state.pop("pending_excel_cache", None)
            st.session_state.pop("pending_selected_sheets", None)
            st.rerun()

    try:
        sheet_info = get_pending_sheet_info(pending_bytes)
    except Exception as e:
        st.error(f"❌ Gagal membaca file Excel Data Pending: {e}")
        return

    all_sheets = list(sheet_info.keys())
    # Pilih otomatis sheet yang memiliki kolom NOSEP/SEP dan JNSPENDING/ALASANPENDING
    default_sheets = [
        s for s, cols in sheet_info.items()
        if any(c in cols for c in ("NOSEP", "SEP"))
        and any(c in cols for c in ("JNSPENDING", "ALASANPENDING"))
    ]
    if not default_sheets:
        default_sheets = [
            s for s, cols in sheet_info.items()
            if any(c in cols for c in ("NOSEP", "SEP", "NOSJP"))
        ] or all_sheets[:1]

    selected_sheets = st.multiselect(
        "📑 Pilih Sheet Data Pending yang Akan Dicocokkan:",
        options=all_sheets,
        default=default_sheets,
        key="pending_selected_sheets",
    )

    if not selected_sheets:
        st.warning("⚠️ Pilih minimal 1 sheet dari file Excel Data Pending.")
        return

    with st.spinner("🔍 Mencocokkan No. SEP dengan Data Pending..."):
        df_pending_raw = load_pending_excel(pending_bytes, tuple(selected_sheets))
        df_pending_matched = match_pending_data(
            df_validated,
            df_pending_raw,
            st.session_state.get("pending_notes", {}),
        )

    st.divider()

    # KPI Ringkasan Pending
    n_val_total = len(df_validated)
    n_pend_src = len(df_pending_raw)
    n_matched = len(df_pending_matched)
    n_pend_ri = (
        len(df_pending_matched[df_pending_matched["JENIS_RAWAT"] == "Rawat Inap"])
        if not df_pending_matched.empty and "JENIS_RAWAT" in df_pending_matched.columns
        else 0
    )
    n_pend_rj = (
        len(df_pending_matched[df_pending_matched["JENIS_RAWAT"] == "Rawat Jalan"])
        if not df_pending_matched.empty and "JENIS_RAWAT" in df_pending_matched.columns
        else 0
    )
    n_filled_ket = (
        int((df_pending_matched["KETERANGAN"].astype(str).str.strip() != "").sum())
        if not df_pending_matched.empty
        else 0
    )

    pm1, pm2, pm3, pm4, pm5 = st.columns(5)
    pm1.metric("📋 Data Tervalidasi", n_val_total)
    pm2.metric("📄 SEP di File Pending", n_pend_src)
    pm3.metric("⚠️ Cocok Masuk Pending", n_matched)
    pm4.metric("🛏️ Pending RI / 🚶 RJ", f"{n_pend_ri} / {n_pend_rj}")
    pm5.metric("📝 Keterangan Diisi", f"{n_filled_ket}/{n_matched}")

    if df_pending_matched.empty:
        st.warning(
            "Tidak ditemukan No. SEP pada data tervalidasi yang cocok dengan Data Pending pada sheet terpilih."
        )
        return

    st.divider()

    # Filter tabel pending
    pf_c1, pf_c2, pf_c3 = st.columns(3)
    with pf_c1:
        pend_filter_jenis = st.selectbox(
            "Filter Jenis Rawat:",
            ["Semua", "Rawat Inap", "Rawat Jalan"],
            key="pend_filter_jenis",
        )
    with pf_c2:
        pend_filter_status = st.selectbox(
            "Filter Status Validasi:",
            ["Semua"] + STATUS_OPTIONS,
            key="pend_filter_status",
        )
    with pf_c3:
        pend_search = st.text_input(
            "Cari SEP / Nama / Alasan Pending:",
            placeholder="Ketik No. SEP, nama, atau kata kunci...",
            key="pend_search",
        )

    df_pend_view = df_pending_matched.copy()
    if pend_filter_jenis != "Semua" and "JENIS_RAWAT" in df_pend_view.columns:
        df_pend_view = df_pend_view[df_pend_view["JENIS_RAWAT"] == pend_filter_jenis]
    if pend_filter_status != "Semua":
        df_pend_view = df_pend_view[df_pend_view["STATUS_VALIDASI"] == pend_filter_status]
    if pend_search.strip():
        q_p = pend_search.strip().upper()
        p_mask = pd.Series(False, index=df_pend_view.index)
        for col in ("SEP", "NAMA_PASIEN", "FILTRASI", "JNSPENDING", "ALASANPENDING", "KETERANGAN"):
            if col in df_pend_view.columns:
                p_mask = p_mask | df_pend_view[col].astype(str).str.upper().str.contains(q_p, na=False)
        df_pend_view = df_pend_view[p_mask]

    st.caption(
        f"Menampilkan **{len(df_pend_view)}** dari **{n_matched}** SEP yang masuk Data Pending "
        "| 💡 Klik kolom **KETERANGAN** untuk mengisi keterangan perbaikan langsung pada tabel, lalu klik **Simpan Keterangan**."
    )

    pend_table_cols = [
        "SEP", "NAMA_PASIEN", "JENIS_RAWAT", "STATUS_EMOJI", "STATUS_VALIDASI",
        "DIAGLIST", "PROCLIST", "KETERANGAN_REVISI",
        "FILTRASI", "JNSPENDING", "ALASANPENDING", "KETERANGAN",
    ]
    pend_avail_cols = [c for c in pend_table_cols if c in df_pend_view.columns]
    df_pend_edit = df_pend_view[pend_avail_cols].reset_index(drop=True).copy()
    df_pend_edit.insert(0, "NO", range(1, len(df_pend_edit) + 1))
    pend_disabled_cols = [c for c in df_pend_edit.columns if c != "KETERANGAN"]

    edited_pend_df = st.data_editor(
        df_pend_edit,
        use_container_width=True,
        height=440,
        num_rows="fixed",
        disabled=pend_disabled_cols,
        column_config={
            "NO": st.column_config.NumberColumn("No.", width=60),
            "STATUS_EMOJI": st.column_config.TextColumn("", width=30),
            "SEP": st.column_config.TextColumn("No. SEP", width=180),
            "JENIS_RAWAT": st.column_config.TextColumn("Jenis Rawat", width=100),
            "STATUS_VALIDASI": st.column_config.TextColumn("Status Modul", width=155),
            "KETERANGAN_REVISI": st.column_config.TextColumn("Keterangan Revisi Modul", width=220),
            "FILTRASI": st.column_config.TextColumn("FILTRASI", width=160),
            "JNSPENDING": st.column_config.TextColumn("JNSPENDING", width=220),
            "ALASANPENDING": st.column_config.TextColumn("ALASANPENDING", width=260),
            "KETERANGAN": st.column_config.TextColumn(
                "KETERANGAN (Perbaikan) ✏️",
                width=280,
                help="Isi keterangan perbaikan untuk SEP pending ini",
            ),
        },
        hide_index=True,
        key="pending_table_editor",
    )

    # Deteksi perubahan kolom KETERANGAN
    if not df_pend_edit.empty:
        diff_ket_mask = (
            df_pend_edit["KETERANGAN"].fillna("").astype(str)
            != edited_pend_df["KETERANGAN"].fillna("").astype(str)
        )
        n_ket_changed = int(diff_ket_mask.sum())
        if n_ket_changed > 0:
            st.info(f"✏️ Ada **{n_ket_changed}** baris dengan perubahan `KETERANGAN` yang belum disimpan.")
            if st.button(f"💾 Simpan {n_ket_changed} Perubahan Keterangan", type="primary", key="save_pending_notes_btn"):
                for i in df_pend_edit.index[diff_ket_mask]:
                    sep_key = str(df_pend_edit.at[i, "SEP"]).strip().upper()
                    st.session_state["pending_notes"][sep_key] = str(edited_pend_df.at[i, "KETERANGAN"] or "").strip()
                st.success(f"✅ Berhasil menyimpan **{n_ket_changed}** keterangan perbaikan!")
                st.rerun()

    st.divider()
    st.subheader("📥 Download Excel Data Pending")
    st.caption("File Excel hanya berisi data SEP yang termasuk ke dalam Data Pending beserta 4 kolom tambahan (`FILTRASI`, `JNSPENDING`, `ALASANPENDING`, `KETERANGAN`).")

    pend_excel_sig = (
        uploaded_file.name,
        pending_name,
        tuple(selected_sheets),
        len(df_pending_matched),
        tuple(sorted(st.session_state.get("pending_notes", {}).items())),
    )
    pend_excel_cache = st.session_state.get("pending_excel_cache")

    if st.button("📊 Siapkan File Excel Data Pending (.xlsx)", type="secondary", key="prep_pending_excel_btn"):
        with st.spinner("⏳ Membuat file Excel Data Pending..."):
            st.session_state["pending_excel_cache"] = {
                "sig": pend_excel_sig,
                "bytes": build_pending_excel_report(df_pending_matched),
            }
        pend_excel_cache = st.session_state["pending_excel_cache"]

    if pend_excel_cache and pend_excel_cache.get("sig") == pend_excel_sig:
        st.download_button(
            label=f"⬇️ Download Excel Data Pending ({n_matched} SEP) (.xlsx)",
            data=pend_excel_cache["bytes"],
            file_name="Hasil_Cek_Data_Pending_BPJS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="dl_pending_excel_btn",
        )


with tab_pending:
    _render_tab_pending(uploaded_file, df_validated if uploaded_file is not None else None)


# ---------------------------------------------------------------------------
# TAB 2: EDIT STATUS (BULK)
# ---------------------------------------------------------------------------

with tab_edit:
    st.title("✏️ Edit Status Massal (Bulk Edit)")
    st.caption("Ubah status beberapa SEP sekaligus dalam satu langkah")

    if uploaded_file is None:
        st.info("👆 Upload file data klaim terlebih dahulu di sidebar.")
        st.stop()

    st.divider()

    # Quick filter untuk bulk
    st.subheader("1️⃣ Pilih SEP yang Akan Diubah")

    n_bulk_all = len(df_validated)
    n_bulk_ri = len(df_validated[df_validated["JENIS_RAWAT"] == "Rawat Inap"]) if "JENIS_RAWAT" in df_validated.columns else 0
    n_bulk_rj = len(df_validated[df_validated["JENIS_RAWAT"] == "Rawat Jalan"]) if "JENIS_RAWAT" in df_validated.columns else 0

    bulk_jenis_filter = st.radio(
        "Filter Jenis Rawat:",
        ["Rawat Inap", "Rawat Jalan", "Semua"],
        format_func=lambda x: {
            "Rawat Inap": f"🛏️ Rawat Inap ({n_bulk_ri})",
            "Rawat Jalan": f"🚶 Rawat Jalan ({n_bulk_rj})",
            "Semua": f"🏥 Semua ({n_bulk_all})",
        }[x],
        horizontal=True,
        key="bulk_jenis_filter",
    )

    if bulk_jenis_filter != "Semua" and "JENIS_RAWAT" in df_validated.columns:
        df_bulk_base = df_validated[df_validated["JENIS_RAWAT"] == bulk_jenis_filter]
    else:
        df_bulk_base = df_validated

    all_seps = df_bulk_base["SEP"].astype(str).tolist()

    bulk_filter_mode = st.radio(
        "Cara memilih SEP:",
        ["Pilih berdasarkan status saat ini", "Pilih manual dari daftar"],
        horizontal=True,
    )

    selected_seps = []

    if bulk_filter_mode == "Pilih berdasarkan status saat ini":
        bulk_status_filter = st.selectbox(
            "Pilih semua SEP dengan status:",
            STATUS_OPTIONS,
            index=2,
            key="bulk_status_filter",
        )
        filtered = df_bulk_base[df_bulk_base["STATUS_VALIDASI"] == bulk_status_filter]
        filtered_seps = filtered["SEP"].astype(str).tolist()
        st.write(
            f"Ditemukan **{len(filtered_seps)}** SEP **{bulk_jenis_filter}** dengan status **{bulk_status_filter}**"
        )

        if filtered_seps:
            selected_seps = st.multiselect(
                "Pilih SEP (default: semua terpilih):",
                options=filtered_seps,
                default=filtered_seps,
                key=f"bulk_selected_seps_{bulk_jenis_filter}_{bulk_status_filter}",
            )
    else:
        selected_seps = st.multiselect(
            f"Pilih SEP {bulk_jenis_filter} dari daftar:",
            options=all_seps,
            key=f"bulk_manual_seps_{bulk_jenis_filter}",
        )

    if selected_seps:
        st.write(f"**{len(selected_seps)}** SEP terpilih")

        st.divider()
        st.subheader("2️⃣ Status Baru & Alasan")

        with st.form("bulk_edit_form", clear_on_submit=True):
            bulk_new_status = st.selectbox(
                "Status Baru untuk Semua SEP Terpilih",
                options=STATUS_OPTIONS,
                index=0,
                key="bulk_new_status",
            )
            bulk_tujuan = st.selectbox(
                "Tujuan Perubahan",
                options=TUJUAN_OPTIONS,
                index=0,
                key="bulk_tujuan",
            )
            bulk_catatan = st.text_area(
                "Catatan / Alasan Perubahan *",
                placeholder="Contoh: Berkas fisik sudah diverifikasi semua, Koding sudah diperbaiki batch, dll.",
                key="bulk_catatan",
            )
            bulk_verifikator = st.text_input(
                "Nama Verifikator *",
                value=st.session_state.get("last_verifikator", ""),
                placeholder="Nama petugas verifikator",
                key="bulk_verifikator",
            )

            st.warning(
                f"⚠️ Anda akan mengubah status **{len(selected_seps)} SEP** "
                f"menjadi **{bulk_new_status}**"
            )

            bulk_submitted = st.form_submit_button(
                f"💾 Ubah Status {len(selected_seps)} SEP", type="primary"
            )

            if bulk_submitted:
                if not bulk_catatan.strip():
                    st.error("❌ Catatan/alasan perubahan wajib diisi!")
                elif not bulk_verifikator.strip():
                    st.error("❌ Nama verifikator wajib diisi!")
                else:
                    count = 0
                    for sep in selected_seps:
                        row_match = df_validated[df_validated["SEP"].astype(str) == sep]
                        if not row_match.empty:
                            old_status = row_match.iloc[0]["STATUS_VALIDASI"]
                            update_status(
                                sep, old_status, bulk_new_status,
                                bulk_catatan.strip(), bulk_verifikator.strip(),
                                bulk_tujuan,
                            )
                            count += 1
                    st.session_state["last_verifikator"] = bulk_verifikator.strip()
                    st.success(
                        f"✅ Berhasil mengubah status **{count} SEP** "
                        f"menjadi **{bulk_new_status}**"
                    )
                    st.rerun()
    else:
        st.info("Pilih minimal 1 SEP untuk melakukan bulk edit.")

# ---------------------------------------------------------------------------
# TAB 3: RIWAYAT PERUBAHAN
# ---------------------------------------------------------------------------

with tab_riwayat:
    st.title("📜 Riwayat Perubahan Status")
    st.caption("Log seluruh perubahan status yang dilakukan oleh verifikator")

    history = st.session_state.get("change_history", [])

    if not history:
        st.info("Belum ada perubahan status yang dilakukan.")
    else:
        st.metric("📝 Total Perubahan", len(history))
        st.divider()

        # Tabel riwayat
        df_hist = pd.DataFrame(reversed(history)).reindex(
            columns=["sep", "status_lama", "status_baru", "tujuan", "catatan", "verifikator", "waktu"]
        ).fillna("")
        df_hist.columns = ["SEP", "Status Lama", "Status Baru", "Tujuan", "Catatan", "Verifikator", "Waktu"]

        # Filter riwayat
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            hist_filter_sep = st.text_input(
                "🔎 Filter SEP", placeholder="Ketik nomor SEP...", key="hist_filter_sep"
            )
        with col_f2:
            hist_filter_verif = st.text_input(
                "👨‍⚕️ Filter Verifikator", placeholder="Nama verifikator...", key="hist_filter_verif"
            )
        with col_f3:
            hist_filter_tujuan = st.selectbox(
                "🎯 Filter Tujuan", ["Semua"] + TUJUAN_OPTIONS[1:], key="hist_filter_tujuan"
            )

        df_hist_display = df_hist.copy()
        if hist_filter_sep:
            df_hist_display = df_hist_display[
                df_hist_display["SEP"].astype(str).str.contains(hist_filter_sep.strip(), case=False, na=False)
            ]
        if hist_filter_verif:
            df_hist_display = df_hist_display[
                df_hist_display["Verifikator"].astype(str).str.contains(hist_filter_verif.strip(), case=False, na=False)
            ]
        if hist_filter_tujuan != "Semua":
            df_hist_display = df_hist_display[
                df_hist_display["Tujuan"].astype(str) == hist_filter_tujuan
            ]

        st.dataframe(
            df_hist_display,
            use_container_width=True,
            height=400,
            hide_index=True,
            column_config={
                "SEP": st.column_config.TextColumn("SEP", width=180),
                "Status Lama": st.column_config.TextColumn("Status Lama", width=140),
                "Status Baru": st.column_config.TextColumn("Status Baru", width=140),
                "Tujuan": st.column_config.TextColumn("Tujuan", width=130),
                "Catatan": st.column_config.TextColumn("Catatan", width=280),
                "Verifikator": st.column_config.TextColumn("Verifikator", width=120),
                "Waktu": st.column_config.TextColumn("Waktu", width=140),
            },
        )

        # Detail timeline per SEP
        st.divider()
        st.subheader("🕐 Timeline Per SEP")

        unique_seps_in_history = df_hist["SEP"].unique().tolist()
        if unique_seps_in_history:
            timeline_sep = st.selectbox(
                "Pilih SEP untuk melihat timeline:", unique_seps_in_history
            )
            sep_entries = df_hist[df_hist["SEP"] == timeline_sep]
            for _, entry in sep_entries.iterrows():
                e_lama = STATUS_EMOJI_MAP.get(entry["Status Lama"], "")
                e_baru = STATUS_EMOJI_MAP.get(entry["Status Baru"], "")
                tujuan_e = entry["Tujuan"] if str(entry["Tujuan"]).strip() else "-"
                st.write(
                    f"🕐 `{entry['Waktu']}` — {e_lama} **{entry['Status Lama']}** → "
                    f"{e_baru} **{entry['Status Baru']}** oleh **{entry['Verifikator']}** (Tujuan: **{tujuan_e}**)"
                )
                st.write(f"   📝 _{entry['Catatan']}_")

        st.divider()

        # Export & clear
        col_exp, col_clr = st.columns(2)
        with col_exp:
            hist_json = json.dumps(history, ensure_ascii=False, indent=2)
            st.download_button(
                "⬇️ Download Riwayat (.json)",
                data=hist_json.encode("utf-8"),
                file_name="riwayat_perubahan_status.json",
                mime="application/json",
            )
        with col_clr:
            if st.button("🗑️ Hapus Semua Riwayat", type="secondary"):
                st.session_state["change_history"] = []
                st.session_state["status_overrides"] = {}
                st.rerun()

# ---------------------------------------------------------------------------
# TAB 4: KELOLA ATURAN
# ---------------------------------------------------------------------------

with tab_aturan:
    st.title("⚙️ Kelola Aturan Modul")
    st.caption("Lihat, edit, tambah, dan download aturan validasi (rules.json)")

    rules_data = get_rules()

    # Tampilkan tabel aturan
    st.subheader("📋 Daftar Aturan Saat Ini")
    df_rules = flatten_rules_to_table(rules_data)
    st.dataframe(
        df_rules,
        use_container_width=True,
        height=400,
        hide_index=True,
        column_config={
            "Deskripsi": st.column_config.TextColumn("Deskripsi", width=250),
            "Pesan Revisi": st.column_config.TextColumn("Pesan Revisi", width=250),
            "Cek Berkas Fisik": st.column_config.TextColumn("Cek Berkas Fisik", width=200),
        },
    )

    st.divider()

    # --- Form Tambah Aturan Baru ---
    st.subheader("➕ Tambah Aturan Baru")

    with st.form("add_rule_form", clear_on_submit=True):
        cat_names = [c["name"] for c in rules_data.get("categories", [])]
        new_cat_option = st.selectbox(
            "Kategori",
            options=cat_names + ["+ Kategori Baru"],
            index=0,
        )

        new_cat_name = ""
        new_cat_id = ""
        if new_cat_option == "+ Kategori Baru":
            new_cat_name = st.text_input("Nama Kategori Baru", placeholder="Contoh: Anak")
            new_cat_id = st.text_input("ID Kategori Baru", placeholder="Contoh: anak")

        col_a, col_b = st.columns(2)
        with col_a:
            new_rule_id = st.text_input("Rule ID *", placeholder="Contoh: ANAK-001")
            new_rule_name = st.text_input("Nama Aturan *", placeholder="Contoh: Thorax foto untuk Pneumonia Anak")
        with col_b:
            new_rule_type = st.selectbox("Tipe Aturan", [
                "standard", "combination_check", "single_diagnosis_check",
                "specific_code_check", "procedure_combination_check",
                "ventilator_duration_check", "conditional_physical_check",
                "delivery_complication_check", "procedure_requires_diagnosis",
            ])

        new_description = st.text_area("Deskripsi", placeholder="Deskripsi lengkap aturan...")

        col_c, col_d = st.columns(2)
        with col_c:
            new_trigger_diag = st.text_input(
                "Trigger Diagnosis (pisahkan koma)",
                placeholder="Contoh: J15.-, A09.9",
            )
            new_req_proc_mode = st.selectbox("Mode Prosedur Wajib", ["all", "any", "none"])
            new_req_proc_codes = st.text_input(
                "Kode Prosedur Wajib (pisahkan koma)",
                placeholder="Contoh: 87.49, 90.42",
            )
        with col_d:
            new_revision_msg = st.text_area(
                "Pesan Revisi",
                placeholder="Pesan yang muncul jika aturan dilanggar...",
            )
            new_physical_check = st.text_area(
                "Cek Berkas Fisik",
                placeholder="Syarat berkas fisik yang harus diverifikasi...",
            )

        submitted = st.form_submit_button("💾 Simpan Aturan Baru", type="primary")

        if submitted:
            if not new_rule_id or not new_rule_name:
                st.error("Rule ID dan Nama Aturan wajib diisi!")
            else:
                # Build rule object
                new_rule = {
                    "id": new_rule_id.strip(),
                    "name": new_rule_name.strip(),
                    "description": new_description.strip(),
                    "type": new_rule_type if new_rule_type != "standard" else None,
                }

                # Remove None type for standard
                if new_rule["type"] is None:
                    del new_rule["type"]

                if new_trigger_diag.strip():
                    new_rule["trigger_diagnoses"] = [d.strip() for d in new_trigger_diag.split(",") if d.strip()]
                    new_rule["trigger_mode"] = "any"

                if new_req_proc_mode != "none" and new_req_proc_codes.strip():
                    new_rule["required_procedures"] = {
                        "mode": new_req_proc_mode,
                        "codes": [c.strip() for c in new_req_proc_codes.split(",") if c.strip()],
                    }
                else:
                    new_rule["required_procedures"] = None

                if new_revision_msg.strip():
                    new_rule["revision_message"] = new_revision_msg.strip()
                if new_physical_check.strip():
                    new_rule["physical_document_check"] = new_physical_check.strip()

                # Tambahkan ke rules_data
                target_cat = None
                if new_cat_option == "+ Kategori Baru" and new_cat_name:
                    target_cat = {
                        "id": new_cat_id.strip() or new_cat_name.lower().replace(" ", "_"),
                        "name": new_cat_name.strip(),
                        "rules": [],
                    }
                    rules_data["categories"].append(target_cat)
                else:
                    for cat in rules_data["categories"]:
                        if cat["name"] == new_cat_option:
                            target_cat = cat
                            break

                if target_cat:
                    target_cat["rules"].append(new_rule)
                    st.session_state["rules_data"] = rules_data
                    save_rules_to_disk(rules_data)
                    st.success(f"✅ Aturan **{new_rule_id}** berhasil ditambahkan ke kategori **{target_cat['name']}**!")
                    st.rerun()
                else:
                    st.error("Kategori tidak ditemukan.")

    st.divider()

    # --- Edit Aturan Existing ---
    st.subheader("✏️ Edit Aturan Existing")

    all_rule_ids = []
    for cat in rules_data.get("categories", []):
        for rule in cat.get("rules", []):
            all_rule_ids.append(f"{cat['name']} / {rule['id']} - {rule['name']}")

    if all_rule_ids:
        selected_rule_label = st.selectbox("Pilih aturan untuk diedit:", all_rule_ids)

        # Parse selection
        sel_parts = selected_rule_label.split(" / ")
        sel_cat_name = sel_parts[0]
        sel_rule_id = sel_parts[1].split(" - ")[0].strip()

        # Find the rule
        edit_rule = None
        edit_cat = None
        for cat in rules_data["categories"]:
            if cat["name"] == sel_cat_name:
                edit_cat = cat
                for rule in cat["rules"]:
                    if rule["id"] == sel_rule_id:
                        edit_rule = rule
                        break
                break

        if edit_rule:
            with st.form("edit_rule_form"):
                st.write(f"**Mengedit:** {edit_rule['id']} — {edit_rule['name']}")

                edit_name = st.text_input("Nama", value=edit_rule.get("name", ""))
                edit_desc = st.text_area("Deskripsi", value=edit_rule.get("description", ""))

                diag_list = edit_rule.get("trigger_diagnoses", edit_rule.get("trigger_group_a", []))
                edit_trigger = st.text_input(
                    "Trigger Diagnosis (koma)",
                    value=", ".join(diag_list),
                )

                req_procs = edit_rule.get("required_procedures")
                if req_procs and req_procs.get("codes"):
                    edit_procs = st.text_input(
                        "Prosedur Wajib (koma)",
                        value=", ".join(req_procs["codes"]),
                    )
                    edit_proc_mode = st.selectbox(
                        "Mode", ["all", "any"],
                        index=0 if req_procs.get("mode") == "all" else 1,
                    )
                else:
                    edit_procs = st.text_input("Prosedur Wajib (koma)", value="")
                    edit_proc_mode = st.selectbox("Mode", ["all", "any", "none"], index=2)

                edit_revision = st.text_area("Pesan Revisi", value=edit_rule.get("revision_message", ""))
                edit_physical = st.text_area("Cek Berkas Fisik", value=edit_rule.get("physical_document_check", "") or "")

                save_edit = st.form_submit_button("💾 Simpan Perubahan", type="primary")

                if save_edit:
                    edit_rule["name"] = edit_name.strip()
                    edit_rule["description"] = edit_desc.strip()

                    if edit_trigger.strip():
                        new_diags = [d.strip() for d in edit_trigger.split(",") if d.strip()]
                        if "trigger_diagnoses" in edit_rule:
                            edit_rule["trigger_diagnoses"] = new_diags
                        elif "trigger_group_a" in edit_rule:
                            edit_rule["trigger_group_a"] = new_diags

                    if edit_proc_mode != "none" and edit_procs.strip():
                        edit_rule["required_procedures"] = {
                            "mode": edit_proc_mode,
                            "codes": [c.strip() for c in edit_procs.split(",") if c.strip()],
                        }
                    elif edit_proc_mode == "none":
                        edit_rule["required_procedures"] = None

                    if edit_revision.strip():
                        edit_rule["revision_message"] = edit_revision.strip()
                    if edit_physical.strip():
                        edit_rule["physical_document_check"] = edit_physical.strip()
                    elif "physical_document_check" in edit_rule:
                        edit_rule["physical_document_check"] = None

                    st.session_state["rules_data"] = rules_data
                    save_rules_to_disk(rules_data)
                    st.success(f"✅ Aturan **{edit_rule['id']}** berhasil diperbarui!")
                    st.rerun()

    st.divider()

    # --- Hapus Aturan ---
    st.subheader("🗑️ Hapus Aturan")
    if all_rule_ids:
        del_rule_label = st.selectbox("Pilih aturan untuk dihapus:", all_rule_ids, key="del_rule")
        del_parts = del_rule_label.split(" / ")
        del_cat_name = del_parts[0]
        del_rule_id = del_parts[1].split(" - ")[0].strip()

        if st.button("🗑️ Hapus Aturan Ini", type="secondary"):
            for cat in rules_data["categories"]:
                if cat["name"] == del_cat_name:
                    cat["rules"] = [r for r in cat["rules"] if r["id"] != del_rule_id]
                    break
            st.session_state["rules_data"] = rules_data
            save_rules_to_disk(rules_data)
            st.success(f"✅ Aturan **{del_rule_id}** berhasil dihapus!")
            st.rerun()

    st.divider()

    # --- Download & Upload rules.json ---
    st.subheader("📥📤 Import / Export Aturan")

    col_dl, col_ul = st.columns(2)
    with col_dl:
        rules_json_str = json.dumps(rules_data, ensure_ascii=False, indent=2)
        st.download_button(
            label="⬇️ Download rules.json Terbaru",
            data=rules_json_str.encode("utf-8"),
            file_name="rules.json",
            mime="application/json",
        )

    with col_ul:
        uploaded_rules = st.file_uploader(
            "📂 Import rules.json",
            type=["json"],
            help="Upload file rules.json untuk mengganti aturan saat ini",
        )
        if uploaded_rules:
            try:
                new_rules = json.load(uploaded_rules)
                if "categories" in new_rules:
                    st.session_state["rules_data"] = new_rules
                    save_rules_to_disk(new_rules)
                    st.success("✅ Rules.json berhasil diimport!")
                    st.rerun()
                else:
                    st.error("❌ Format rules.json tidak valid (harus ada key 'categories').")
            except json.JSONDecodeError:
                st.error("❌ File bukan JSON yang valid.")

# ---------------------------------------------------------------------------
# FOOTER
# ---------------------------------------------------------------------------

st.sidebar.divider()
st.sidebar.caption("© 2026 Validator BPJS Pengendali v1.0")
st.sidebar.caption("github.com/Rikiyp")
