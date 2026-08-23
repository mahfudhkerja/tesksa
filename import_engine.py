"""
import_engine.py
=================
Logic inti yang dipakai bersama oleh semua script import_*.py, supaya
SOURCE_SHEET_ID dan SHEETS_TO_IMPORT tidak lagi hardcoded di tiap file,
melainkan dibaca dari config.json (satu folder yang sama) saat script
dijalankan. Config itu sendiri diisi lewat halaman "Input Data" di
index.html -> endpoint Flask di app.py -> file ini.

Alur:
  1. Load (paste link)   -> extract_id_from_link() + detect_sheets()
  2. Pilih Sheet (centang)-> update_source(sheets=[...])
  3. Refresh / Run All    -> run_all.py memanggil tiap import_*.py,
                             yang masing2 memanggil run_gsheet_import()
                             atau run_excel_import() di file ini.
"""

import io
import json
import re
import time
import datetime
from pathlib import Path

import gspread
import pandas as pd
from google.auth.exceptions import TransportError
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from requests.exceptions import ConnectionError as RequestsConnectionError, ReadTimeout

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
CREDENTIALS_FILE = BASE_DIR / "credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ============================================================
# RETRY — dipakai untuk semua panggilan yang menyentuh jaringan
# (auth token ke oauth2.googleapis.com, buka spreadsheet, dsb),
# supaya timeout/koneksi putus sesaat tidak langsung menggagalkan
# seluruh script (lihat kasus import_ex.py yang gagal di tahap
# refresh token, bukan karena datanya bermasalah).
# ============================================================

RETRYABLE_EXCEPTIONS = (TransportError, RequestsConnectionError, ReadTimeout, TimeoutError)


def _with_retry(fn, *args, attempts=3, base_delay=5, label="", **kwargs):
    """Panggil fn(*args, **kwargs), otomatis coba ulang kalau kena error
    jaringan/timeout yang retryable. Delay antar percobaan naik bertahap
    (5s, 10s, ...). Kalau semua percobaan gagal, error asli dilempar lagi."""
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return fn(*args, **kwargs)
        except RETRYABLE_EXCEPTIONS as e:
            last_exc = e
            if attempt == attempts:
                break
            wait = base_delay * attempt
            tag = f" ({label})" if label else ""
            print(f"   ⚠️ Koneksi bermasalah{tag}: {e}. Coba lagi dalam {wait}s... [percobaan {attempt}/{attempts}]")
            time.sleep(wait)
    raise last_exc

# ============================================================
# CONFIG HELPERS
# ============================================================

def load_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(f"config.json tidak ditemukan di {CONFIG_PATH}")
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(cfg):
    tmp_path = CONFIG_PATH.with_suffix(".json.tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
    tmp_path.replace(CONFIG_PATH)  # tulis atomik, hindari config.json korup kalau proses lain baca bersamaan


def get_source(source_key):
    cfg = load_config()
    src = cfg.get("sources", {}).get(source_key)
    if src is None:
        raise KeyError(f"Source '{source_key}' tidak ada di config.json")
    return cfg, src


def update_source(source_key, **fields):
    cfg, src = get_source(source_key)
    src.update(fields)
    cfg["sources"][source_key] = src
    save_config(cfg)
    return src


def set_import_result(source_key, status, rows_written=None, error=None):
    update_source(
        source_key,
        last_import=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        last_status=status,
        last_rows=rows_written,
        last_error=error,
    )


# ============================================================
# SHEET "ListMesin" — simpan link tiap source ke kolom B,
# dicocokkan lewat nama mesin di kolom A.
# ============================================================

LIST_MESIN_SHEET_NAME = "ListMesin"

# source_key (di config.json) -> nama persis di kolom A sheet ListMesin.
SOURCE_KEY_TO_MESIN_NAME = {
    "printing_1": "PRINTING 1",
    "printing_2": "PRINTING 2",
    "printing_3": "PRINTING 3",
    "printing_4": "PRINTING 4",
    "printing_5": "PRINTING 5",
    "dry_1": "DRY 1",
    "dry_2": "DRY 2",
    "dry_3": "DRY 3",
    "dry_4": "DRY 4",
    "dry_5": "DRY 5",
    "rw": "REWIND BESAR",
    "ex": "EXTRUSI",
    "sf": "SOLVENT FREE",
    "sl": "SLITTING",
}


def update_list_mesin_link(source_key, link):
    """Simpan link yang baru dipaste user (di halaman Input Data) ke sheet
    'ListMesin' -> kolom B (LINK), pada baris yang kolom A-nya cocok dengan
    nama mesin untuk source_key ini (lihat SOURCE_KEY_TO_MESIN_NAME).

    Dipanggil dari endpoint /api/produksi/load setelah link berhasil
    terhubung. Kalau nama mesin tidak ada di mapping, atau baris tidak
    ditemukan di sheet ListMesin, fungsi ini tidak melempar error keras
    (supaya tidak menggagalkan proses Load utama) — cukup return False.
    """
    mesin_name = SOURCE_KEY_TO_MESIN_NAME.get(source_key)
    if not mesin_name:
        return False

    cfg = load_config()
    target_id = cfg.get("target_sheet_id")
    if not target_id:
        return False

    client = get_gspread_client()
    target_sp = client.open_by_key(target_id)
    try:
        ws = target_sp.worksheet(LIST_MESIN_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return False

    col_a = ws.col_values(1)  # kolom A, termasuk header
    row_idx = None
    for i, val in enumerate(col_a, start=1):
        if _norm(val) == _norm(mesin_name):
            row_idx = i
            break

    if row_idx is None:
        return False

    ws.update_cell(row_idx, 2, link)  # kolom B = LINK
    return True


# ============================================================
# LINK PARSING
# ============================================================

def extract_id_from_link(text):
    """Terima link Google Sheets / Drive, ATAU langsung ID mentah,
    kembalikan spreadsheet/file ID-nya."""
    text = (text or "").strip()
    m = re.search(r"/d/([a-zA-Z0-9_-]{15,})", text)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{15,})", text)
    if m:
        return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]{15,}", text):
        return text
    raise ValueError("Link/ID spreadsheet tidak valid.")


# ============================================================
# GOOGLE CLIENTS
# ============================================================

def _build_gspread_client():
    creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    return gspread.authorize(creds)


def get_gspread_client():
    return _with_retry(_build_gspread_client, label="auth gspread")


def _build_drive_service():
    creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def get_drive_service():
    return _with_retry(_build_drive_service, label="auth drive")


def _download_drive_file(file_id):
    drive_service = get_drive_service()

    def _do_download():
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        return fh

    return _with_retry(_do_download, label=f"download {file_id}")


# ============================================================
# DETEKSI NAMA SHEET/TAB — dipakai endpoint "Load"
# ============================================================

def detect_sheets(source_id, source_type):
    """Kembalikan (list_nama_sheet, nama_file) untuk sumber tertentu."""
    if source_type == "gsheet":
        client = get_gspread_client()
        sp = client.open_by_key(source_id)
        names = [ws.title for ws in sp.worksheets()]
        return names, sp.title
    elif source_type == "excel":
        drive_service = get_drive_service()
        meta = drive_service.files().get(fileId=source_id, fields="name").execute()
        fh = _download_drive_file(source_id)
        excel_file = pd.ExcelFile(fh)
        return excel_file.sheet_names, meta.get("name", "")
    else:
        raise ValueError(f"source_type tidak dikenal: {source_type}")


# ============================================================
# HELPER PARSING TABEL — dipakai kedua varian import di bawah
# ============================================================

def _norm(text):
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    clean = str(text).strip().upper()
    clean = clean.replace("_", " ").replace("-", " ")
    return " ".join(clean.split())


def find_header_row(rows, keywords):
    for i, row in enumerate(rows):
        if not row:
            continue
        row_clean = [_norm(c) for c in row if c is not None]
        for keyword in keywords:
            kw = _norm(keyword)
            if any(kw in cell for cell in row_clean):
                return i
    return None


def get_column_mapping(header_row, target_headers, max_cols=100):
    source_map = {}
    for idx, cell in enumerate(header_row[:max_cols]):
        clean_name = _norm(cell)
        if clean_name and clean_name not in source_map:
            source_map[clean_name] = idx
    return {th: source_map.get(_norm(th)) for th in target_headers}


def _is_cell_empty(cell):
    return cell is None or (isinstance(cell, float) and pd.isna(cell)) or str(cell).strip() == ""


def get_data_rows(rows, header_index, max_empty_streak=20):
    """Ambil baris data setelah header; berhenti kalau ada 20 baris
    berturut-turut dengan kolom A-F kosong (dianggap akhir data sheet ini)."""
    data = []
    empty_streak = 0
    for i in range(header_index + 1, len(rows)):
        row = rows[i]
        cols_a_f = row[:6] if len(row) >= 6 else list(row) + [""] * (6 - len(row))
        if all(_is_cell_empty(c) for c in cols_a_f):
            empty_streak += 1
            if empty_streak >= max_empty_streak:
                break
            continue
        empty_streak = 0

        row_text = " ".join(str(c).strip().upper() for c in row if c and not _is_cell_empty(c))
        if "TGL/BLN/THN" in row_text or "SHIFT/OPERATOR" in row_text:
            continue
        junk_keywords = ["NO ROLL", "METER AKHIR", "KG", "JAM", "SETING", "MESIN RUSAK", "TOTAL", "JUMLAH"]
        if any(k in row_text for k in junk_keywords):
            continue
        data.append(row)
    return data


def _map_row(row, target_headers, mapping, sanitize_fn):
    row = list(row) + [""] * (len(target_headers) - len(row))
    out = []
    for th in target_headers:
        idx = mapping.get(th)
        val = row[idx] if idx is not None and idx < len(row) else ""
        out.append(sanitize_fn(val))
    return out


def _write_target(target_sp, target_sheet_name, all_rows, target_headers):
    try:
        old_ws = target_sp.worksheet(target_sheet_name)
        target_sp.del_worksheet(old_ws)
        print("   Sheet lama dihapus.")
    except gspread.exceptions.WorksheetNotFound:
        pass

    rows_count = max(len(all_rows) + 10, 100)
    cols_count = max(len(target_headers) + 5, 26)
    target_ws = target_sp.add_worksheet(title=target_sheet_name, rows=rows_count, cols=cols_count)
    print(f"   Sheet baru '{target_sheet_name}' dibuat.")

    if all_rows:
        target_ws.update(range_name="A1", values=all_rows, value_input_option="RAW")
        print(f"✅ Berhasil menulis {len(all_rows)-1} baris data.")

    _reorder_target_sheets(target_sp)

    return len(all_rows) - 1 if all_rows else 0


# ============================================================
# RAPIKAN URUTAN TAB — semua sheet hasil import ditaruh
# berurutan tepat SETELAH tab "Data", biar rapi & konsisten
# tiap kali Refresh Semua dijalankan.
# ============================================================

# Urutan yang diinginkan untuk sheet hasil import (samakan dengan urutan
# card di halaman Input Data / SCRIPTS_ORDER di run_all.py).
TARGET_SHEET_ORDER = [
    "PRINTING_2", "PRINTING_3", "PRINTING_4", "PRINTING_5",
    "RW_1",
    "DRY_1", "DRY_2", "DRY_3", "DRY_4", "DRY_5",
    "EX_1",
    "SF_1",
    "SL_1",
]

ANCHOR_SHEET_NAME = "Data"


def _reorder_target_sheets(target_sp, anchor_sheet_name=ANCHOR_SHEET_NAME):
    """Susun ulang urutan tab di spreadsheet target: semua sheet hasil
    import (lihat TARGET_SHEET_ORDER) dipindah supaya berurutan tepat
    setelah tab 'Data', tanpa mengganggu urutan tab lain (Login, Validasi,
    PIC, ListMesin, dll) yang tetap di posisi relatifnya masing-masing.

    Kalau tab 'Data' tidak ditemukan, fallback: taruh semua sheet hasil
    import di akhir (perilaku lama). Kalau reorder gagal karena sebab
    apa pun, jangan sampai menggagalkan proses import (cukup print
    peringatan)."""
    try:
        all_ws = target_sp.worksheets()  # urutan tab saat ini
        by_title = {ws.title: ws for ws in all_ws}

        target_titles = set(TARGET_SHEET_ORDER)
        produced_ws_in_order = [by_title[name] for name in TARGET_SHEET_ORDER if name in by_title]
        other_ws = [ws for ws in all_ws if ws.title not in target_titles]

        anchor_idx = None
        for i, ws in enumerate(other_ws):
            if _norm(ws.title) == _norm(anchor_sheet_name):
                anchor_idx = i
                break

        if anchor_idx is None:
            new_order = other_ws + produced_ws_in_order
        else:
            before = other_ws[: anchor_idx + 1]  # termasuk tab 'Data' itu sendiri
            after = other_ws[anchor_idx + 1:]
            new_order = before + produced_ws_in_order + after

        target_sp.reorder_worksheets(new_order)
    except Exception as e:
        print(f"   ⚠️ Gagal merapikan urutan tab (dilewati, data tetap tertulis): {e}")


# ============================================================
# DETEKSI & URUTKAN SHEET PER BULAN — supaya hasil di sheet target
# selalu kronologis (JAN -> DES), walau nama tab tiap sumber beda-beda
# format (JAN/JANUARI/JANU, AGS/AGUST/AGUSTUS/AGUS, dst) dan urutan
# tab di spreadsheet sumber / urutan centang user tidak berurutan.
# ============================================================

# Tiap bulan dipetakan ke beberapa kemungkinan awalan (prefix) nama,
# mencakup ejaan Indonesia & Inggris, singkatan panjang/pendek, dan
# typo umum yang sudah ditemukan di sumber-sumber yang ada (mis. JULIE).
_MONTH_PREFIXES = [
    (1, ("JANUARI", "JANUAR", "JANU", "JAN")),
    (2, ("FEBRUARI", "FEBRUARY", "FEBR", "FEB")),
    (3, ("MARET", "MARCH", "MAR")),
    (4, ("APRIL", "APR")),
    (5, ("MEI", "MAY")),
    (6, ("JUNI", "JUNE", "JUN")),
    (7, ("JULI", "JULY", "JULIE", "JUL")),
    (8, ("AGUSTUS", "AUGUST", "AGUST", "AGUS", "AGS", "AUG", "AG")),
    (9, ("SEPTEMBER", "SEPT", "SEP")),
    (10, ("OKTOBER", "OCTOBER", "OKT", "OCT")),
    (11, ("NOVEMBER", "NOV")),
    (12, ("DESEMBER", "DECEMBER", "DES", "DEC")),
]


def _detect_month_number(sheet_name):
    """Deteksi nomor bulan (1-12) dari nama sheet/tab, walau formatnya
    beda-beda antar sumber (mis. 'JAN 2026', 'JANUARI', 'Agus 2026',
    'AGUST 26', 'FEBRUARI2026', 'MAR', dst). Return None kalau nama
    sheet-nya sama sekali tidak mengandung nama bulan yang dikenali."""
    letters = re.match(r"[A-Za-z]+", (sheet_name or "").strip())
    if not letters:
        return None
    word = letters.group(0).upper()
    for month_num, prefixes in _MONTH_PREFIXES:
        for p in prefixes:
            if word.startswith(p):
                return month_num
    return None


def _detect_year(sheet_name):
    """Deteksi tahun dari nama sheet kalau ada (4 digit, atau 2 digit
    seperti 'AGUST 26' -> 2026). Return 0 kalau tidak ada info tahun
    sama sekali (mis. sheet cuma bernama 'MARET')."""
    text = sheet_name or ""
    m = re.search(r"(20\d{2})", text)
    if m:
        return int(m.group(1))
    m = re.search(r"(?<!\d)(\d{2})(?!\d)", text)
    if m:
        return 2000 + int(m.group(1))
    return 0


def _sheet_month_sort_key(sheet_name):
    """Key buat sorted(): urut utama per BULAN (Jan->Des), tahun cuma
    tiebreaker kalau ada beberapa tahun bercampur. Sheet yang nama
    bulannya tidak dikenali ditaruh paling akhir (bukan bikin proses
    gagal), urutan aslinya di antara sesama yang 'tidak dikenal' tetap
    terjaga karena Python sorted() itu stable sort."""
    month = _detect_month_number(sheet_name)
    year = _detect_year(sheet_name)
    if month is None:
        return (99, 9999)
    return (month, year)


# ============================================================
# VARIAN 1: IMPORT LANGSUNG DARI GOOGLE SHEET (printing/rw/sl/sf)
# ============================================================

def _sanitize_cell_gsheet(cell):
    if cell is None or (isinstance(cell, str) and cell.strip() == ""):
        return "-"
    if isinstance(cell, (datetime.datetime, datetime.date, datetime.time)):
        return cell.isoformat()
    return cell


def import_sheets_aligned(source_id, target_id, sheets_to_import, target_sheet_name,
                           target_headers, header_keywords):
    client = get_gspread_client()
    source_sp = _with_retry(client.open_by_key, source_id, label=f"open source {source_id}")
    target_sp = _with_retry(client.open_by_key, target_id, label=f"open target {target_id}")

    ws_list = []
    if sheets_to_import:
        for name in sheets_to_import:
            try:
                ws_list.append(source_sp.worksheet(name))
            except gspread.exceptions.WorksheetNotFound:
                print(f"⚠️ Sheet '{name}' tidak ditemukan, dilewati.")
    else:
        ws_list = source_sp.worksheets()
        print(f"📋 Mengimpor semua sheet ({len(ws_list)} sheet).")

    if not ws_list:
        print("Tidak ada sheet yang akan diimpor.")
        return 0

    # Urutkan per bulan (bukan urutan tab di sumber / urutan centang user),
    # supaya hasil akhir di sheet target selalu kronologis JAN -> DES.
    ws_list = sorted(ws_list, key=lambda ws: _sheet_month_sort_key(ws.title))
    print("   🗓️ Urutan proses (per bulan): " + ", ".join(ws.title for ws in ws_list))

    all_rows = [target_headers]
    for ws in ws_list:
        print(f"\n🔍 Memproses sheet: {ws.title}")
        rows = ws.get_all_values()
        if not rows:
            print("   Sheet kosong, dilewati.")
            continue
        header_idx = find_header_row(rows, header_keywords)
        if header_idx is None:
            print("   ❌ Tidak ditemukan baris header, dilewati.")
            continue
        header_row = rows[header_idx]
        print(f"   ✅ Header ditemukan di baris {header_idx+1}")
        mapping = get_column_mapping(header_row, target_headers)
        data_rows = get_data_rows(rows, header_idx)
        print(f"   📊 Jumlah baris data valid: {len(data_rows)}")
        for row in data_rows:
            all_rows.append(_map_row(row, target_headers, mapping, _sanitize_cell_gsheet))

    print(f"\n📝 Menulis ke sheet tujuan '{target_sheet_name}'...")
    return _write_target(target_sp, target_sheet_name, all_rows, target_headers)


def run_gsheet_import(source_key, target_sheet_name, target_headers, header_keywords):
    """Dipanggil dari tiap import_printing_X.py / import_rw.py / import_sl.py / import_sf.py"""
    cfg, src = get_source(source_key)
    source_id = src.get("source_id")
    sheets = src.get("sheets") or []
    target_id = cfg["target_sheet_id"]
    if not source_id:
        raise RuntimeError(f"'{source_key}': belum ada link spreadsheet sumber (isi lewat halaman Input Data).")
    if not sheets:
        raise RuntimeError(f"'{source_key}': belum ada sheet yang dicentang (isi lewat halaman Input Data).")
    try:
        rows_written = import_sheets_aligned(source_id, target_id, sheets, target_sheet_name, target_headers, header_keywords)
        set_import_result(source_key, "OK", rows_written=rows_written)
        return rows_written
    except Exception as e:
        set_import_result(source_key, "ERROR", error=str(e))
        raise


# ============================================================
# VARIAN 2: IMPORT DARI FILE EXCEL (.xlsx) DI GOOGLE DRIVE (dry_1..5)
# ============================================================

def _sanitize_cell_excel(cell):
    if cell is None or (isinstance(cell, float) and pd.isna(cell)) or (isinstance(cell, str) and cell.strip() == ""):
        return "-"
    if isinstance(cell, (datetime.datetime, datetime.date)):
        return cell.strftime("%d-%m-%Y")
    if isinstance(cell, datetime.time):
        return cell.strftime("%H:%M:%S")
    if isinstance(cell, datetime.timedelta):
        total_seconds = int(cell.total_seconds())
        h, m, s = total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"
    return cell


def import_excel_from_drive(source_id, target_id, sheets_to_import, target_sheet_name,
                             target_headers, header_keywords):
    client = get_gspread_client()
    target_sp = _with_retry(client.open_by_key, target_id, label=f"open target {target_id}")

    fh = _download_drive_file(source_id)
    excel_file = pd.ExcelFile(fh)
    print("✅ Berhasil membaca file Excel dari Drive.")
    available_sheets = excel_file.sheet_names

    # Urutkan per bulan (bukan urutan centang user), supaya hasil akhir
    # di sheet target selalu kronologis JAN -> DES.
    sheets_to_import = sorted(sheets_to_import, key=_sheet_month_sort_key)
    print("   🗓️ Urutan proses (per bulan): " + ", ".join(sheets_to_import))

    all_rows = [target_headers]
    for sheet_name in sheets_to_import:
        if sheet_name not in available_sheets:
            print(f"⚠️ Sheet '{sheet_name}' tidak ditemukan di file Excel, dilewati.")
            continue
        print(f"\n🔍 Memproses sheet: {sheet_name}")
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
        rows = df.values.tolist()
        if not rows:
            print("   Sheet kosong, dilewati.")
            continue
        header_idx = find_header_row(rows, header_keywords)
        if header_idx is None:
            print("   ❌ Tidak ditemukan baris header, dilewati.")
            continue
        header_row = rows[header_idx]
        print(f"   ✅ Header ditemukan di baris {header_idx+1}")
        mapping = get_column_mapping(header_row, target_headers)
        data_rows = get_data_rows(rows, header_idx)
        print(f"   📊 Jumlah baris data valid: {len(data_rows)}")
        for row in data_rows:
            row = ["" if (isinstance(c, float) and pd.isna(c)) else c for c in row]
            all_rows.append(_map_row(row, target_headers, mapping, _sanitize_cell_excel))

    print(f"\n📝 Menulis ke sheet tujuan '{target_sheet_name}'...")
    return _write_target(target_sp, target_sheet_name, all_rows, target_headers)


def run_excel_import(source_key, target_sheet_name, target_headers, header_keywords):
    """Dipanggil dari tiap import_dry_X.py"""
    cfg, src = get_source(source_key)
    source_id = src.get("source_id")
    sheets = src.get("sheets") or []
    target_id = cfg["target_sheet_id"]
    if not source_id:
        raise RuntimeError(f"'{source_key}': belum ada link file Excel sumber (isi lewat halaman Input Data).")
    if not sheets:
        raise RuntimeError(f"'{source_key}': belum ada sheet yang dicentang (isi lewat halaman Input Data).")
    try:
        rows_written = import_excel_from_drive(source_id, target_id, sheets, target_sheet_name, target_headers, header_keywords)
        set_import_result(source_key, "OK", rows_written=rows_written)
        return rows_written
    except Exception as e:
        set_import_result(source_key, "ERROR", error=str(e))
        raise
