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
from collections import OrderedDict

import gspread
import pandas as pd
from google.auth.exceptions import TransportError
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

import classify_gudang_sheets  # klasifikasi BJB/BJL, dipanggil dari run_gudang_import()
from googleapiclient.http import MediaIoBaseDownload
from requests.exceptions import ConnectionError as RequestsConnectionError, ReadTimeout

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"

# Folder tempat file Excel/WPS "Data Gudang" hasil UPLOAD lewat browser
# disimpan di server. Sengaja BUKAN path folder di komputer user -- dulu
# fitur ini baca file langsung dari path lokal (mis. Z:\...), yang cuma
# jalan kalau app.py dijalankan di komputer yang sama dengan foldernya.
# Begitu di-deploy online (Render dst), server tidak punya akses ke
# harddisk komputer user, jadi diganti jadi "upload file ke server dulu,
# baru server yang baca file itu dari disknya sendiri".
GUDANG_UPLOAD_DIR = BASE_DIR / "gudang_uploads"

# Folder tempat file Excel hasil UPLOAD untuk source "update_stock" (kartu
# "Update Stock" di Input Data Produksi, mode file -- lihat
# save_update_stock_upload() & run_update_stock_import() di bawah).
# Sama prinsipnya dengan GUDANG_UPLOAD_DIR di atas, folder terpisah supaya
# file Gudang & Update Stock tidak saling menimpa.
UPDATE_STOCK_UPLOAD_DIR = BASE_DIR / "update_stock_uploads"

# Sama polanya dengan UPDATE_STOCK_UPLOAD_DIR di atas, tapi untuk dua kartu
# upload-file baru: "FORM SERAH TERIMA 2" (source_key "form_st_2") & "VALIDASI 2"
# (source_key "val_2"). Keduanya SUMBER KEDUA/tambahan -- terpisah dari
# import_form_st.py & import_val.py yang lama (yang jalan lewat
# /api/validasi/refresh-import), dan dari import_update_stock.py di atas.
# Folder terpisah supaya file yang diupload lewat ketiga kartu ini (Update
# Stock, Form Serah Terima 2, Validasi 2) tidak saling menimpa.
FORM_ST2_UPLOAD_DIR = BASE_DIR / "form_st2_uploads"
VAL2_UPLOAD_DIR = BASE_DIR / "val2_uploads"

# source_key -> folder upload-nya masing2, dipakai save_local_excel_upload()
# untuk SEMUA source yang file-nya diupload lokal (bukan link) -- baik yang
# dibaca lewat blok kolom lebar 10 tanpa header (update_stock, lihat
# _read_stacked_block_excel_sheet) MAUPUN yang header-based (form_st_2,
# val_2, lihat run_local_excel_import / VARIAN 5) -- supaya ketiganya bisa
# pakai fungsi upload yang SAMA tanpa duplikasi kode 3x.
LOCAL_EXCEL_UPLOAD_DIRS = {
    "update_stock": UPDATE_STOCK_UPLOAD_DIR,
    "form_st_2": FORM_ST2_UPLOAD_DIR,
    "val_2": VAL2_UPLOAD_DIR,
}

# Dipakai app.py (produksi_sources()) buat mengecualikan ketiga source ini
# dari daftar kartu generik link+checklist -- sama seperti GUDANG_SOURCES,
# ketiganya punya kartu upload-file hardcoded sendiri di index.html.
FILE_UPLOAD_EXTRA_SOURCES = set(LOCAL_EXCEL_UPLOAD_DIRS.keys())

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


def _is_quota_error(e):
    """True kalau exception ini gspread.exceptions.APIError dengan status
    429 (RESOURCE_EXHAUSTED / 'Quota exceeded ... per minute') dari Google
    Sheets API. Beda dari RETRYABLE_EXCEPTIONS di atas (yang soal jaringan/
    timeout) -- 429 ini soal terlalu banyak request dalam satu menit,
    jadi butuh delay lebih panjang (kuota reset per menit) supaya
    percobaan ulang berikutnya nggak langsung kena 429 lagi juga."""
    if not isinstance(e, gspread.exceptions.APIError):
        return False
    try:
        status = e.response.status_code
    except AttributeError:
        status = None
    if status == 429:
        return True
    # Jaga-jaga kalau status code-nya nggak kebaca tapi pesannya jelas kuota.
    return "RESOURCE_EXHAUSTED" in str(e) or "Quota exceeded" in str(e)


def _with_retry(fn, *args, attempts=3, base_delay=5, label="", **kwargs):
    """Panggil fn(*args, **kwargs), otomatis coba ulang kalau kena error
    jaringan/timeout ATAU kena limit kuota Google Sheets (429). Delay
    antar percobaan untuk error jaringan naik bertahap (5s, 10s, ...);
    untuk 429 sengaja dipaksa nunggu lebih lama (kuota "per menit", jadi
    percobaan ulang perlu nunggu minimal ~20-40 detik supaya jendela
    kuotanya beneran reset, bukan langsung kena 429 lagi). Kalau semua
    percobaan gagal, error asli dilempar lagi."""
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
        except gspread.exceptions.APIError as e:
            if not _is_quota_error(e):
                raise
            last_exc = e
            if attempt == attempts:
                break
            wait = 20 * attempt  # 20s, 40s, ... -- kuota Sheets API reset per menit
            tag = f" ({label})" if label else ""
            print(f"   ⚠️ Kena limit kuota Google Sheets{tag}: {e}. Coba lagi dalam {wait}s... [percobaan {attempt}/{attempts}]")
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
    "rewind_kecil": "REWIND KECIL",
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
# DETEKSI OTOMATIS TIPE SOURCE (gsheet vs excel) — dipakai endpoint
# "Load" supaya user TIDAK PERLU pilih manual lagi (dulu ada dropdown
# khusus di kartu "Rewind Kecil" karena sumbernya bisa dua-duanya).
#
# Google Sheets asli (dibuat langsung dari Sheets, atau "convert" hasil
# upload) punya mimeType "application/vnd.google-apps.spreadsheet" di
# Drive API. File Excel/WPS biasa yang cuma DISIMPAN di Drive (belum
# dikonversi) punya mimeType lain (mis. .xlsx/.xls) -- itu yang perlu
# didownload dulu baru dibaca lewat pandas (lihat run_excel_import).
# ============================================================

GOOGLE_SHEETS_MIME_TYPE = "application/vnd.google-apps.spreadsheet"


def detect_source_type(source_id):
    """Kembalikan 'gsheet' atau 'excel' berdasarkan mimeType file ini di
    Google Drive -- tidak perlu tahu isinya, cukup metadata (cepat)."""
    drive_service = get_drive_service()

    def _get_meta():
        return drive_service.files().get(fileId=source_id, fields="mimeType").execute()

    meta = _with_retry(_get_meta, label=f"deteksi tipe {source_id}")
    mime = meta.get("mimeType", "")
    return "gsheet" if mime == GOOGLE_SHEETS_MIME_TYPE else "excel"


# ============================================================
# DETEKSI NAMA SHEET/TAB — dipakai endpoint "Load"
# ============================================================

def detect_sheets(source_id, source_type=None):
    """Kembalikan (list_nama_sheet, nama_file) untuk sumber tertentu.
    source_type dideteksi otomatis lewat detect_source_type() kalau
    tidak dikirim eksplisit oleh pemanggil."""
    if source_type is None:
        source_type = detect_source_type(source_id)
    if source_type == "gsheet":
        client = get_gspread_client()
        sp = client.open_by_key(source_id)
        names = [ws.title for ws in sp.worksheets()]
        return names, sp.title
    elif source_type == "excel":
        drive_service = get_drive_service()
        meta = drive_service.files().get(fileId=source_id, fields="name").execute()
        fh = _download_drive_file(source_id)
        names = _fast_excel_sheet_names(fh)
        return names, meta.get("name", "")
    else:
        raise ValueError(f"source_type tidak dikenal: {source_type}")


def _fast_excel_sheet_names(fh):
    """Ambil daftar nama sheet dari file .xlsx TANPA membuka/parsing tiap
    worksheet-nya lewat openpyxl.

    Sebelumnya `pd.ExcelFile(fh)` dipakai cuma untuk ambil nama sheet,
    padahal itu bikin openpyxl scan XML SETIAP worksheet satu per satu
    (buat hitung dimensi datanya) -- kerja yang mahal & lambat kalau
    file-nya besar/banyak sheet, dan ini penyebab WORKER TIMEOUT (lalu
    worker di-kill) di endpoint /api/produksi/load.

    File .xlsx sebenarnya adalah file ZIP, dan daftar nama sheet-nya ada
    di dalam xl/workbook.xml -- file kecil yang isinya cuma metadata,
    TIDAK berisi data sel sama sekali. Jadi baca file itu saja: hasilnya
    instan berapapun besar/banyak data di tiap sheet.
    """
    import zipfile
    from xml.etree import ElementTree as ET

    fh.seek(0)
    try:
        with zipfile.ZipFile(fh) as z:
            xml_bytes = z.read("xl/workbook.xml")
        root = ET.fromstring(xml_bytes)
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheets_el = root.find("main:sheets", ns)
        names = [s.get("name") for s in sheets_el] if sheets_el is not None else []
        if names:
            return names
    except Exception:
        pass  # bukan .xlsx (zip) valid -- kemungkinan .xls lama, fallback di bawah

    # Fallback untuk format yang bukan .xlsx modern (mis. .xls lama):
    # tetap pakai cara lama, lebih lambat tapi tetap benar.
    fh.seek(0)
    excel_file = pd.ExcelFile(fh)
    return excel_file.sheet_names


# ============================================================
# HELPER PARSING TABEL — dipakai kedua varian import di bawah
# ============================================================

def _norm(text):
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    clean = str(text).strip().upper()
    clean = clean.replace("_", " ").replace("-", " ")
    return " ".join(clean.split())


def find_header_row(rows, keywords, min_matches=1):
    """Cari baris header: baris pertama yang punya minimal `min_matches`
    kata kunci (dari `keywords`) muncul BERSAMAAN di baris yang SAMA.

    Default min_matches=1 (perilaku lama: cukup 1 kata kunci cocok di
    mana saja di baris itu) -- dipakai source yang kata kuncinya khas
    dan hampir mustahil muncul sendirian di baris data (mis. "TANGGAL",
    "SHIFT/OPERATOR").

    Untuk source yang kata kuncinya cukup umum dan bisa saja kebetulan
    muncul sendirian di baris data (mis. "PRODUK", "PROSES" di LP),
    panggil dengan min_matches lebih tinggi (mis. sejumlah semua
    keywords) supaya baris cuma dianggap header kalau kata-kata itu
    ketemu BARENGAN dalam 1 baris yang sama -- bukan asal 1 kata cocok
    di baris manapun secara acak."""
    for i, row in enumerate(rows):
        if not row:
            continue
        row_clean = [_norm(c) for c in row if c is not None]
        matched = 0
        for keyword in keywords:
            kw = _norm(keyword)
            if any(kw in cell for cell in row_clean):
                matched += 1
        if matched >= min_matches:
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


DEFAULT_JUNK_KEYWORDS = ["NO ROLL", "METER AKHIR", "KG", "JAM", "SETING", "MESIN RUSAK", "TOTAL", "JUMLAH"]


def get_data_rows(rows, header_index, max_empty_streak=20, junk_keywords=DEFAULT_JUNK_KEYWORDS):
    """Ambil baris data setelah header; berhenti kalau ada 20 baris
    berturut-turut dengan kolom A-F kosong (dianggap akhir data sheet ini).

    junk_keywords: list kata kunci baris "sampah" (baris ringkasan/summary
    khas sheet mesin produksi) yang akan di-skip. Kosongkan (None atau [])
    untuk source yang datanya sendiri wajar mengandung kata-kata itu.

    Kolom apapun yang nama headernya mengandung "KETERANGAN" (mis.
    "KETERANGAN", "KETERANGAN_JADI") SELALU dikecualikan dari pengecekan
    junk_keywords, karena kolom itu teks bebas dan sering memuat kata
    seperti "KG"/"JAM"/"TOTAL"/"JUMLAH" secara wajar sebagai catatan --
    bukan indikasi baris ringkasan/sampah."""
    header_row = rows[header_index] if 0 <= header_index < len(rows) else []
    keterangan_cols = {i for i, h in enumerate(header_row) if "KETERANGAN" in str(h).strip().upper()}

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

        row_text = " ".join(
            str(c).strip().upper()
            for idx, c in enumerate(row)
            if idx not in keterangan_cols and c and not _is_cell_empty(c)
        )
        if "TGL/BLN/THN" in row_text or "SHIFT/OPERATOR" in row_text:
            continue
        if junk_keywords and any(k in row_text for k in junk_keywords):
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
    "RW_1", "REWIND",
    "DRY_1", "DRY_2", "DRY_3", "DRY_4", "DRY_5",
    "EX_1",
    "SF_1",
    "SL_1",
    "BAG_1",
    "JO_1",
    "LP_1",
    "VAL_1",
    "FORM_ST_1",
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
# UPDATE STOCK — beda dari source lain: 1 sheet sumber bisa berisi
# BEBERAPA "blok" kolom berdampingan (mis. A-J, lalu lompat 1 kolom
# gap, L-U, lompat lagi, W-AF, dst), masing2 blok lebar 10 kolom dan
# TANPA HEADER (baris 1 sudah data mentah). Semua blok dari semua
# sheet yang dicentang ditumpuk vertikal (blok 1 lalu blok 2 dst,
# sheet 1 lalu sheet 2 dst, sesuai urutan `sheets` di config), lalu
# ditulis ke SATU spreadsheet tujuan yang TERPISAH dari target_sheet_id
# utama -- HANYA kolom B..K (10 kolom) yang ditulis/dibersihkan, kolom
# A tujuan TIDAK PERNAH disentuh (isinya rumus manual di sheet tujuan
# yang mengacu ke kolom B, biar tetap jalan dia tidak boleh diusik).
#
# Beda lain dari source biasa: sheet tujuan TIDAK dihapus/dibuat ulang
# (_write_target biasa menghapus seluruh tab) -- di sini cuma range
# B2:K{n} yang di-clear lalu ditulis ulang, supaya kolom A & sheet itu
# sendiri tetap utuh.
# ============================================================

UPDATE_STOCK_BLOCK_WIDTH = 10
UPDATE_STOCK_BLOCK_GAP = 1


def _row_cell(row, idx):
    return row[idx] if idx < len(row) else ""


def _extract_stacked_blocks(rows, block_width=UPDATE_STOCK_BLOCK_WIDTH, gap=UPDATE_STOCK_BLOCK_GAP):
    """rows: hasil worksheet.get_all_values() APA ADANYA (baris 1 = data,
    TIDAK ADA header yang dilewati). Kembalikan list-of-list (tiap baris
    lebar `block_width`), gabungan SEMUA blok kolom di sheet ini yang
    ditumpuk vertikal berurutan (blok pertama dulu, baru blok kedua, dst).

    Blok ke-n (0-based) mulai di kolom index n*(block_width+gap):
      blok 0: A-J (idx 0-9), blok 1: L-U (idx 11-20), blok 2: W-AF (idx 22-31), ...
    Berhenti menambah blok baru begitu kolom PERTAMA blok itu (turun ke
    bawah, seluruh sheet) sama sekali tidak ada teks/nilai apapun --
    ini HANYA dipakai untuk tahu ada berapa blok di sheet ini (lompat
    ke kanan), BUKAN untuk berhenti membaca baris ke bawah.

    Di dalam satu blok TIDAK ADA cutoff "sekian baris kosong berturut2
    dianggap akhir data" -- baris kosong di tengah cukup dilewati
    (tidak ikut ditulis), tapi pembacaan tetap lanjut sampai baris
    terakhir sheet (rows habis), tidak pernah berhenti lebih awal
    karena ada gap kosong."""
    out = []
    block_start = 0
    step = block_width + gap
    while True:
        has_any = any(str(_row_cell(r, block_start)).strip() not in ("", "-") for r in rows)
        if not has_any:
            break

        for r in rows:
            cells = [str(_row_cell(r, block_start + i)) for i in range(block_width)]
            if all(c.strip() == "" for c in cells):
                continue  # baris kosong di tengah blok: lewati saja, tetap lanjut baca ke bawah
            out.append(cells)

        block_start += step
    return out


def run_update_stock_import(source_key="update_stock"):
    """Baca semua sheet yang dicentang di source `update_stock` (config.json),
    ekstrak blok-blok kolom lebar 10 dari tiap sheet (lihat
    _extract_stacked_blocks), tumpuk semua hasilnya, lalu tulis ke
    kolom B..K sheet tujuan (src["target_id"] / src["target_sheet"] di
    config.json) -- TANPA menghapus sheet tujuan & TANPA menyentuh
    kolom A (rumus manual di sana biarkan apa adanya).

    HANYA dipakai source "update_stock" (blok kolom lebar 10, TANPA
    header) -- BEDA dari "form_st_2"/"val_2" yang sumbernya ADA header
    (sama seperti form_st/val lama), lihat run_local_excel_import()
    (VARIAN 5) untuk dua source itu.

    Sheet-nya sendiri bisa dibaca dari DUA jenis sumber, tergantung
    src["input_mode"]:
      - "file" (baru, lewat kartu "Update Stock" -> upload Excel
        langsung, mirip Data Gudang) -- baca dari file yang tersimpan
        di UPDATE_STOCK_UPLOAD_DIR lewat _read_update_stock_excel_sheet().
      - "link" / tidak diisi (lama, dipertahankan untuk kompatibilitas
        kalau ada yang masih pakai cara lama) -- baca dari spreadsheet
        src["source_id"] lewat gspread, seperti sebelumnya."""
    cfg, src = get_source(source_key)

    target_id = src.get("target_id")
    target_sheet_name = src.get("target_sheet")
    if not target_id:
        raise ValueError(f"config.json: sources.{source_key}.target_id belum diisi (ID spreadsheet tujuan).")
    if not target_sheet_name:
        raise ValueError(f"config.json: sources.{source_key}.target_sheet belum diisi (nama tab tujuan).")

    input_mode = src.get("input_mode", "link")
    combined = []

    if input_mode == "file":
        filename = src.get("filename")
        card_label = LOCAL_EXCEL_CARD_LABELS.get(source_key, source_key)
        if not filename:
            raise ValueError(f"Source '{source_key}' belum ada file yang diupload (klik kartu \"{card_label}\" di halaman Input Data Produksi dulu).")
        for sheet_name in src.get("sheets", []):
            try:
                rows = _read_stacked_block_excel_sheet(source_key, filename, sheet_name)
            except ValueError as e:
                print(f"   ⚠️ {e}")
                continue
            blocks = _extract_stacked_blocks(rows)
            print(f"   Sheet '{sheet_name}': {len(blocks)} baris (semua blok digabung).")
            combined.extend(blocks)
        client = get_gspread_client()  # tetap dibutuhkan buat nulis ke target di bawah
    else:
        if not src.get("source_id"):
            raise ValueError("Source 'update_stock' belum terhubung ke spreadsheet manapun (klik Load dulu di halaman Input Data Produksi).")
        client = get_gspread_client()
        source_sp = client.open_by_key(src["source_id"])
        for sheet_name in src.get("sheets", []):
            try:
                ws = source_sp.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                print(f"   ⚠️ Sheet '{sheet_name}' tidak ditemukan di sumber, dilewati.")
                continue
            rows = _with_retry(ws.get_all_values, label=f"baca {sheet_name}")
            blocks = _extract_stacked_blocks(rows)
            print(f"   Sheet '{sheet_name}': {len(blocks)} baris (semua blok digabung).")
            combined.extend(blocks)

    target_sp = client.open_by_key(target_id)
    try:
        target_ws = target_sp.worksheet(target_sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError(f"Tab '{target_sheet_name}' tidak ditemukan di spreadsheet tujuan ({target_id}).")

    card_label = LOCAL_EXCEL_CARD_LABELS.get(source_key, source_key)
    existing = _with_retry(target_ws.get_all_values, label=f"baca sheet tujuan {card_label}")
    old_last_row = len(existing)  # termasuk baris 1 (walau baris 1 di sini bukan header, tetap ikut dihitung)
    new_last_row = 1 + len(combined)  # data mulai baris 2
    clear_last_row = max(old_last_row, new_last_row)

    end_col = gspread.utils.rowcol_to_a1(1, 1 + UPDATE_STOCK_BLOCK_WIDTH).rstrip("1")  # 'B'..+9 kolom -> 'K'

    if clear_last_row >= 2:
        _with_retry(
            target_ws.batch_clear, [f"B2:{end_col}{clear_last_row}"],
            label=f"hapus data lama {card_label} (kolom B-K)",
        )
    if combined:
        _with_retry(
            target_ws.update, f"B2:{end_col}{new_last_row}", combined,
            value_input_option="RAW", label=f"tulis data {card_label} (kolom B-K)",
        )

    print(f"   ✅ {len(combined)} baris ditulis ke kolom B-K sheet '{target_sheet_name}' (kolom A tidak disentuh).")
    return len(combined)


# ============================================================
# SYNC tab "UpdateStock" (spreadsheet UTAMA, cfg["target_sheet_id"])
# dari JO_1 -- BEDA dari run_update_stock_import() di atas, yang
# menulis ke spreadsheet EKSTERNAL "Monitor Bahan Baku" (source
# "update_stock" di config.json, dari sheet2 PET/PL/CPPM dst).
#
# Header tab "UpdateStock" (dikonfirmasi dari UPDATESTOCK_COLUMN_MAP
# di app.py + tampilan tabelnya di index.html):
#   A=JO, B=NAMA, C=ORDER, D=METER ORDER, E=METER VALIDASI,
#   F=LAPISAN ORDER, G=LAPISAN VALIDASI, H=KETERANGAN, I=ACC
#
# Kolom yang diisi fungsi ini dari JO_1:
#   - kolom A = JO_1 kolom F (kode JO)
#   - kolom B = JO_1 kolom G (KEMASAN / nama produk)
#   - kolom C (ORDER) = JO_1 kolom H (ORDER -- sama seperti JO_COL_ORDER
#     yang dipakai fitur Validasi)
#   - kolom D (METER ORDER) = JO_1 kolom M (METER) -- ditulis APA ADANYA
#     termasuk kalau isinya "-", TIDAK di-skip
#   - kolom F (LAPISAN ORDER) = gabungan JO_1 kolom P, Q, R
#     (TEXTJOIN ", ", nilai kosong ATAU "-" dilewati -- bukan ikut
#     jadi koma kosong di tengah)
#   - kolom H (KETERANGAN) = JO_1 kolom S
#
# Kolom E (METER VALIDASI) & G (LAPISAN VALIDASI) diisi BUKAN dari
# JO_1, tapi dari spreadsheet EKSTERNAL "Monitor Bahan Baku" (ID
# UPDATE_STOCK_MONITOR_SPREADSHEET_ID di bawah), tab
# UPDATE_STOCK_MONITOR_SHEET_NAME:
#   - Cocokkan kolom A tab UpdateStock (kode JO yang baru saja
#     dibangun di atas) dengan kolom B sheet tersebut.
#   - Kalau cocok (bisa lebih dari satu baris cocok untuk satu JO):
#       kolom E (METER VALIDASI)   = TEXTJOIN ", " dari kolom E baris2
#                                     yang cocok (nilai kosong dilewati)
#       kolom G (LAPISAN VALIDASI) = TEXTJOIN ", " dari kolom A baris2
#                                     yang cocok (nilai kosong dilewati)
#   - Kalau tidak ada yang cocok, kolom E/G untuk JO itu ditulis "".
#
# Baris JO_1 yang diproses: mulai dari baris PERTAMA yang kode JO-nya
# berformat JO/<UPDATE_STOCK_JO1_START_YEAR>/<...BULAN ROMAWI>/<...HARI>/<no>,
# lalu SEMUA baris JO_1 di bawahnya sampai akhir sheet ikut disertakan
# APA ADANYA -- tidak difilter tanggal lagi setelah titik itu (jadi JO
# bertanggal lebih lama yang muncul di bawah titik itu tetap ikut).
# Tidak ada baris yang di-skip gara-gara salah satu kolomnya "-".
# ============================================================

UPDATESTOCK_SHEET_NAME = "UpdateStock"

UPDATE_STOCK_JO1_START_YEAR = "26"            # "26" di JO/26/VIII/27/...
UPDATE_STOCK_JO1_START_MONTH_ROMAWI = "VIII"  # Agustus
UPDATE_STOCK_JO1_START_DAY = 27

# Spreadsheet EKSTERNAL "Monitor Bahan Baku" dipakai untuk isi kolom E
# (METER VALIDASI) & G (LAPISAN VALIDASI) tab UpdateStock -- BEDA dari
# spreadsheet "Monitor Bahan Baku" tujuan run_update_stock_import() di
# atas (yang itu diatur lewat config.json / target_id, ini di-hardcode
# karena sifatnya lookup tetap, bukan konfigurasi per-user).
UPDATE_STOCK_MONITOR_SPREADSHEET_ID = "1lSj54tQP8QKMR96HiAHBCd1Zm-fOxnkdt9x3gD1tuEM"
UPDATE_STOCK_MONITOR_SHEET_NAME = "UPDATE_STOCK"
UPDATE_STOCK_MONITOR_COL_JO = 1   # kolom B (0-based) -- dicocokkan dengan kolom A UpdateStock (kode JO)
UPDATE_STOCK_MONITOR_COL_E = 4    # kolom E (0-based) -> sumber UpdateStock kolom E (METER VALIDASI)
UPDATE_STOCK_MONITOR_COL_A = 0    # kolom A (0-based) -> sumber UpdateStock kolom G (LAPISAN VALIDASI)


def _build_update_stock_monitor_lookup(client):
    """Baca sheet UPDATE_STOCK_MONITOR_SHEET_NAME di spreadsheet
    UPDATE_STOCK_MONITOR_SPREADSHEET_ID (baris 1 = header, dilewati),
    kembalikan dict {kode_jo (kolom B, distrip): [baris, baris, ...]}
    -- satu kode JO bisa punya beberapa baris cocok (untuk textjoin)."""
    ws_monitor = _with_retry(
        client.open_by_key, UPDATE_STOCK_MONITOR_SPREADSHEET_ID,
        label="open Monitor Bahan Baku (kolom E/G UpdateStock)",
    ).worksheet(UPDATE_STOCK_MONITOR_SHEET_NAME)
    monitor_rows = _with_retry(ws_monitor.get_all_values, label="baca UPDATE_STOCK (Monitor Bahan Baku)")

    lookup = {}
    for row in monitor_rows[1:]:  # lewati header
        jo_val = str(_row_cell(row, UPDATE_STOCK_MONITOR_COL_JO)).strip()
        if not jo_val:
            continue
        lookup.setdefault(jo_val, []).append(row)
    return lookup


def _textjoin_monitor_col(matched_rows, col_idx):
    """TEXTJOIN ", " nilai kolom `col_idx` dari `matched_rows`, nilai
    kosong dilewati (bukan "-", cuma string kosong yang di-skip)."""
    parts = [str(_row_cell(r, col_idx)).strip() for r in matched_rows]
    return ", ".join(p for p in parts if p != "")

# Kolom (0-based) tambahan di JO_1 yang dipakai fitur ini (JO_COL_JO /
# JO_COL_NAMA / JO_COL_ORDER sudah didefinisikan di atas untuk fitur
# Validasi, index 5 / 6 / 7 -- dipakai ulang di sini karena sama
# persis: kolom F, G, H).
JO_COL_METER = 12       # kolom M
JO_COL_LAPISAN_1 = 15   # kolom P
JO_COL_LAPISAN_2 = 16   # kolom Q
JO_COL_LAPISAN_3 = 17   # kolom R
JO_COL_KETERANGAN = 18  # kolom S


def _jo1_row_matches_update_stock_start(jo_code):
    """True kalau kode JO (kolom F JO_1) persis berformat
    JO/<tahun>/<bulan romawi>/<tanggal>/<nomor> dengan tahun/bulan/
    tanggal sesuai UPDATE_STOCK_JO1_START_*."""
    if not jo_code:
        return False
    parts = str(jo_code).strip().split("/")
    if len(parts) < 4:
        return False
    yy, bulan, dd = parts[1].strip(), parts[2].strip().upper(), parts[3].strip()
    try:
        day_num = int(dd)
    except ValueError:
        return False
    return (
        yy == UPDATE_STOCK_JO1_START_YEAR
        and bulan == UPDATE_STOCK_JO1_START_MONTH_ROMAWI
        and day_num == UPDATE_STOCK_JO1_START_DAY
    )


def sync_update_stock_from_jo():
    """Isi kolom A/B/C/D/F/H tab 'UpdateStock' (spreadsheet utama) dari
    JO_1, dan kolom E/G dari lookup ke spreadsheet "Monitor Bahan
    Baku" (UPDATE_STOCK_MONITOR_SPREADSHEET_ID, tab UPDATE_STOCK) --
    lihat penjelasan lengkap di komentar blok di atas.

    Kolom ACC PERLU perhatian
    khusus: karena semua baris ditulis ulang dari nol tiap refresh
    (supaya baris JO yang sudah tidak relevan tidak nyangkut), status
    ACC yang sudah dikunci dibaca dulu SEBELUM clear -- per KODE JO
    DIGABUNG "sidik jari" datanya (NAMA/ORDER/METER ORDER/LAPISAN
    ORDER/KETERANGAN), bukan per posisi baris atau kode JO doang --
    lalu ditempel lagi ke baris JO yang sama setelah data baru ditulis
    HANYA KALAU fingerprint-nya masih identik. Ini supaya status
    "sudah di-ACC" tidak ketuker ke JO lain / hilang kalau urutan
    baris JO berubah antar refresh, TAPI kalau JO itu ternyata sudah
    direvisi (kode sama tapi meter/bahan berubah karena baris lama di
    JO_1 dihapus lalu diganti baris baru), ACC-nya otomatis direset
    ke kosong supaya wajib di-ACC ulang -- tidak ikut ke-ACC otomatis
    dari data lama yang sudah tidak berlaku."""
    cfg = load_config()
    target_id = cfg["target_sheet_id"]
    client = get_gspread_client()
    target_sp = _with_retry(client.open_by_key, target_id, label=f"open target {target_id}")

    ws_jo = target_sp.worksheet(JO_SOURCE_SHEET_NAME)
    ws_update = target_sp.worksheet(UPDATESTOCK_SHEET_NAME)

    # ---- 1. Baca JO_1, cari baris awal (JO tanggal start), ambil semua baris dari situ ke bawah ----
    jo_rows = _with_retry(ws_jo.get_all_values, label="baca JO_1 (sync UpdateStock)")
    max_col_needed = max(JO_COL_JO, JO_COL_NAMA, JO_COL_ORDER, JO_COL_METER, JO_COL_LAPISAN_3, JO_COL_KETERANGAN)

    start_idx = None
    for i, row in enumerate(jo_rows[1:], start=1):  # index asli di jo_rows, lewati header (index 0)
        if len(row) <= JO_COL_JO:
            continue
        if _jo1_row_matches_update_stock_start(row[JO_COL_JO]):
            start_idx = i
            break

    if start_idx is None:
        print(
            f"   ⚠️ Tidak ada JO di JO_1 berformat JO/{UPDATE_STOCK_JO1_START_YEAR}/"
            f"{UPDATE_STOCK_JO1_START_MONTH_ROMAWI}/{UPDATE_STOCK_JO1_START_DAY}/... "
            f"-- tidak ada yang disinkron ke UpdateStock."
        )
        return 0

    # ---- 2. Baca status ACC LAMA + "sidik jari" data LAMA (per kode JO) ----
    # PENTING: kode JO bisa dipakai ULANG kalau JO itu direvisi (baris lama
    # di JO_1 dihapus biar tidak dobel, baris baru dgn kode SAMA muncul lagi
    # tapi datanya beda -- mis. METER ORDER atau LAPISAN/bahan berubah).
    # Kalau ACC cuma dicocokkan lewat kode JO doang, revisi begini bakal
    # ke-ACC otomatis padahal belum pernah direview -- makanya di sini kita
    # simpan juga fingerprint (NAMA, ORDER, METER ORDER, LAPISAN ORDER,
    # KETERANGAN) dari data LAMA. Nanti pas rewrite, ACC lama HANYA
    # ditempel lagi kalau fingerprint baru == fingerprint lama (datanya
    # benar-benar belum berubah). Kalau beda dikit aja -> ACC direset "",
    # wajib di-ACC ulang.
    existing_rows = _with_retry(ws_update.get_all_values, label="baca UpdateStock (sebelum rewrite)")
    existing_header = existing_rows[0] if existing_rows else []
    try:
        col_acc_idx = existing_header.index("ACC")  # 0-based
    except ValueError:
        col_acc_idx = None
    # UpdateStock: A=JO(0) B=NAMA(1) C=ORDER(2) D=METER ORDER(3) E=METER VALIDASI(4)
    #              F=LAPISAN ORDER(5) G=LAPISAN VALIDASI(6) H=KETERANGAN(7)
    FINGERPRINT_COL_IDXS = (1, 2, 3, 5, 7)  # NAMA, ORDER, METER ORDER, LAPISAN ORDER, KETERANGAN
    acc_lookup = {}  # jo_code -> (acc_value_lama, fingerprint_lama)
    if col_acc_idx is not None:
        for row in existing_rows[1:]:
            if not row:
                continue
            jo_val = str(row[0]).strip() if len(row) > 0 else ""
            if not jo_val:
                continue
            acc_val = row[col_acc_idx] if len(row) > col_acc_idx else ""
            fingerprint_lama = tuple(
                str(row[idx]).strip() if len(row) > idx else "" for idx in FINGERPRINT_COL_IDXS
            )
            acc_lookup[jo_val] = (acc_val, fingerprint_lama)

    # ---- 2b. Baca lookup Monitor Bahan Baku (kolom E/G, cocokkan via kode JO) ----
    monitor_lookup = _build_update_stock_monitor_lookup(client)

    # ---- 3. Bangun baris baru kolom A, B, C, D, E, F, G, H (+ ACC dipertahankan
    #      per JO HANYA kalau datanya belum berubah -- lihat catatan di step 2) ----
    final_a, final_b, final_c, final_d, final_e, final_f, final_g, final_h, final_acc = (
        [], [], [], [], [], [], [], [], [],
    )
    revisi_ter_reset = []  # buat log: JO yang ACC-nya direset gara-gara datanya berubah
    for row in jo_rows[start_idx:]:
        if len(row) <= max_col_needed:
            continue
        jo_code = str(row[JO_COL_JO]).strip()
        if not jo_code:
            continue
        nama = row[JO_COL_NAMA]
        order = row[JO_COL_ORDER]
        meter = row[JO_COL_METER]  # ditulis apa adanya, "-" tetap ikut
        lapisan_parts = [
            str(row[idx]).strip()
            for idx in (JO_COL_LAPISAN_1, JO_COL_LAPISAN_2, JO_COL_LAPISAN_3)
            if str(row[idx]).strip() not in ("", "-")
        ]
        lapisan_gabung = ", ".join(lapisan_parts)
        keterangan = row[JO_COL_KETERANGAN]

        final_a.append([jo_code])
        final_b.append([nama])
        final_c.append([order])
        final_d.append([meter])
        matched_monitor_rows = monitor_lookup.get(jo_code, [])
        final_e.append([_textjoin_monitor_col(matched_monitor_rows, UPDATE_STOCK_MONITOR_COL_E)])
        final_f.append([lapisan_gabung])
        final_g.append([_textjoin_monitor_col(matched_monitor_rows, UPDATE_STOCK_MONITOR_COL_A)])
        final_h.append([keterangan])

        fingerprint_baru = (
            str(nama).strip(), str(order).strip(), str(meter).strip(),
            lapisan_gabung, str(keterangan).strip(),
        )
        acc_val_lama, fingerprint_lama = acc_lookup.get(jo_code, ("", None))
        if fingerprint_lama is not None and fingerprint_lama == fingerprint_baru:
            final_acc.append([acc_val_lama])  # data identik -> status ACC dipertahankan
        else:
            final_acc.append([""])  # baru pertama kali ATAU sudah direvisi -> wajib ACC ulang
            if fingerprint_lama is not None and acc_val_lama == "1":
                revisi_ter_reset.append(jo_code)

    if revisi_ter_reset:
        print(f"   ⚠️ ACC direset karena data berubah (revisi) untuk JO: {', '.join(revisi_ter_reset)}")

    old_last_row = len(existing_rows)
    new_last_row = 1 + len(final_a)
    clear_last_row = max(old_last_row, new_last_row)

    acc_col_letter = (
        gspread.utils.rowcol_to_a1(1, col_acc_idx + 1).rstrip("1") if col_acc_idx is not None else None
    )

    if clear_last_row >= 2:
        ranges_to_clear = [f"A2:H{clear_last_row}"]  # A-D, F, H (dari JO_1) + E, G (dari Monitor Bahan Baku)
        if acc_col_letter:
            ranges_to_clear.append(f"{acc_col_letter}2:{acc_col_letter}{clear_last_row}")
        _with_retry(
            ws_update.batch_clear, ranges_to_clear,
            label="hapus data lama UpdateStock (kolom A-H, ACC)",
        )

    if final_a:
        # Digabung jadi SATU batch_update (dulu 8 pemanggilan .update()
        # terpisah, satu per kolom A-H -- itu 8 write request sendiri-
        # sendiri ke Sheets API tiap kali refresh, gampang banget bikin
        # kena limit 'Read/Write requests per minute per user' apalagi
        # digabung sama semua pembacaan sebelumnya di fungsi ini +
        # run_update_stock_import(). Sekarang cuma 1 request buat kolom
        # A-H, jadi cost kuotanya jauh lebih kecil).
        data = [
            {"range": f"A2:A{new_last_row}", "values": final_a},
            {"range": f"B2:B{new_last_row}", "values": final_b},
            {"range": f"C2:C{new_last_row}", "values": final_c},
            {"range": f"D2:D{new_last_row}", "values": final_d},
            {"range": f"E2:E{new_last_row}", "values": final_e},
            {"range": f"F2:F{new_last_row}", "values": final_f},
            {"range": f"G2:G{new_last_row}", "values": final_g},
            {"range": f"H2:H{new_last_row}", "values": final_h},
        ]
        if acc_col_letter:
            data.append({
                "range": f"{acc_col_letter}2:{acc_col_letter}{new_last_row}",
                "values": final_acc,
            })
        _with_retry(
            ws_update.batch_update, data, value_input_option="RAW",
            label="tulis UpdateStock kolom A-H (+ACC) sekaligus",
        )

    print(
        f"   ✅ {len(final_a)} baris JO (mulai dari JO/{UPDATE_STOCK_JO1_START_YEAR}/"
        f"{UPDATE_STOCK_JO1_START_MONTH_ROMAWI}/{UPDATE_STOCK_JO1_START_DAY}) disinkron ke UpdateStock."
    )
    return len(final_a)


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
                           target_headers, header_keywords, junk_keywords=DEFAULT_JUNK_KEYWORDS,
                           header_min_matches=1):
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
        header_idx = find_header_row(rows, header_keywords, min_matches=header_min_matches)
        if header_idx is None:
            print("   ❌ Tidak ditemukan baris header, dilewati.")
            continue
        header_row = rows[header_idx]
        print(f"   ✅ Header ditemukan di baris {header_idx+1}")
        mapping = get_column_mapping(header_row, target_headers)
        data_rows = get_data_rows(rows, header_idx, junk_keywords=junk_keywords)
        print(f"   📊 Jumlah baris data valid: {len(data_rows)}")
        for row in data_rows:
            all_rows.append(_map_row(row, target_headers, mapping, _sanitize_cell_gsheet))

    print(f"\n📝 Menulis ke sheet tujuan '{target_sheet_name}'...")
    return _write_target(target_sp, target_sheet_name, all_rows, target_headers)


def run_gsheet_import(source_key, target_sheet_name, target_headers, header_keywords,
                       junk_keywords=DEFAULT_JUNK_KEYWORDS, header_min_matches=1,
                       target_id=None):
    """Dipanggil dari tiap import_printing_X.py / import_rw.py / import_sl.py / import_sf.py / import_jo.py / import_lp.py.

    junk_keywords: teruskan [] (list kosong) untuk source yang tidak perlu
    filter baris "sampah" sama sekali (lihat get_data_rows).

    header_min_matches: berapa banyak header_keywords yang harus ketemu
    BARENGAN dalam 1 baris yang sama supaya baris itu dianggap baris
    header (lihat find_header_row). Default 1 (perilaku lama: cukup 1
    kata kunci cocok) -- naikkan (mis. sejumlah len(header_keywords))
    untuk source yang kata kuncinya cukup umum dan berisiko kebetulan
    cocok di baris data (bukan baris header beneran).

    target_id: override spreadsheet TUJUAN tulis. Kalau None (default),
    pakai target_sheet_id global di config.json seperti biasa -- tapi
    beberapa source (mis. "rewind_kecil") sengaja punya spreadsheet
    tujuan SENDIRI yang beda dari warehouse utama (lihat
    import_rewind_kecil.py), jadi harus dikirim eksplisit di sini,
    BUKAN ditaruh di config.json target_sheet_id (yang dipakai bareng
    oleh semua source lain)."""
    cfg, src = get_source(source_key)
    source_id = src.get("source_id")
    sheets = src.get("sheets") or []
    target_id = target_id or cfg["target_sheet_id"]
    if not source_id:
        raise RuntimeError(f"'{source_key}': belum ada link spreadsheet sumber (isi lewat halaman Input Data).")
    if not sheets:
        raise RuntimeError(f"'{source_key}': belum ada sheet yang dicentang (isi lewat halaman Input Data).")
    try:
        rows_written = import_sheets_aligned(source_id, target_id, sheets, target_sheet_name, target_headers,
                                              header_keywords, junk_keywords=junk_keywords,
                                              header_min_matches=header_min_matches)
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


def run_excel_import(source_key, target_sheet_name, target_headers, header_keywords, target_id=None):
    """Dipanggil dari tiap import_dry_X.py. target_id: lihat penjelasan
    di run_gsheet_import() -- override spreadsheet tujuan kalau source ini
    punya tujuan sendiri, beda dari target_sheet_id global."""
    cfg, src = get_source(source_key)
    source_id = src.get("source_id")
    sheets = src.get("sheets") or []
    target_id = target_id or cfg["target_sheet_id"]
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


# ============================================================
# VARIAN 4: IMPORT GUDANG DARI FILE EXCEL/WPS LOKAL (folder network).
# BEDA dari varian 1 & 2 di atas dalam DUA hal:
#   a) sumbernya file lokal/jaringan, bukan link Drive/Sheets;
#   b) pemilihan folder->file->sheet dan EKSEKUSI import-nya terpisah
#      di dua halaman berbeda di frontend:
#        - Kartu "Data Gudang" di halaman "Input Data Produksi": HANYA
#          untuk UPLOAD file lalu pilih sheet, lalu menyimpan pilihan itu
#          ke config.json (save_gudang_upload() + save_gudang_selection()).
#          Kartu ini SENGAJA TIDAK ikut tombol "Refresh Semua" (run_all.py)
#          di halaman itu.
#        - Halaman "Data Gudang BJB" & "Data Gudang BJL" (grup "Gudang"
#          di sidebar): tidak ada input apa pun di sini, cuma tombol
#          Refresh sendiri-sendiri yang MENJALANKAN import
#          (run_gudang_import()) memakai file/sheet yang SUDAH disimpan
#          lewat kartu di atas. Klik Refresh di salah satu halaman ini
#          menjalankan proses yang sama persis (satu sumber, satu
#          tujuan) -- keduanya cuma dua pintu masuk ke tombol Refresh
#          yang sama.
#
# Kenapa upload, bukan baca folder lokal (versi lama)? Karena app.py bisa
# di-deploy online (mis. Render) -- begitu itu terjadi, server jalan di
# komputer LAIN, bukan komputer user, jadi server tidak akan pernah bisa
# baca path folder di laptop user (mis. "C:\Users\...\Downloads\..."
# atau "Z:\...") sama sekali. Solusinya: file-nya yang dikirim (upload)
# dari browser ke server, baru server baca dari disknya sendiri.
#
# Alur lengkap (lihat endpoint /api/gudang/* di app.py):
#   1. Di kartu "Data Gudang": user pilih file lewat <input type="file">
#      -> klik "Upload & Pilih Sheet" -> file dikirim ke server lewat
#      /api/gudang/upload -> save_gudang_upload() SIMPAN file itu di
#      server (GUDANG_UPLOAD_DIR) lalu balikin daftar nama sheet/tab di
#      dalamnya (muncul di modal pilih sheet).
#   2. User pilih 1 sheet -> save_gudang_selection() menyimpan
#      filename+sheet_name ke config.json (source_key "gudang").
#   3. Di halaman Data Gudang BJB *atau* BJL, user klik Refresh ->
#      run_gudang_import() -- baca semua value (bukan formula) dari
#      sheet yang tersimpan itu (calamine, lebih toleran ke file hasil
#      WPS dibanding openpyxl), lalu TIMPA tab "API" (GUDANG_TARGET_SHEET)
#      di spreadsheet GUDANG_TARGET_ID: baris 1 = header baku
#      (GUDANG_HEADER, kolom R sengaja kosong), baris 2 dst = data. Satu
#      tab "API" ini isinya gabungan BJB+BJL, dibedakan lewat kolom
#      AREA -- halaman Data Gudang BJB/BJL di frontend menampilkan hasil
#      filter dari tab yang sama ini, bukan dari tab terpisah lagi.
# ============================================================

GUDANG_TARGET_ID = "1-ZyKSwXLzZaA6uNYRcpJNQZWX_ssYzvX45Z51xERipI"

# Satu-satunya source key gudang di config.json "sources" (dulu ada dua:
# 'gudang_bjb' & 'gudang_bjl', masing2 dengan tab tujuan sendiri --
# sekarang digabung jadi satu sumber, satu tab tujuan "API").
GUDANG_SOURCE_KEY = "gudang"
GUDANG_LABEL = "Data Gudang"
GUDANG_TARGET_SHEET = "API"

# Dipertahankan dalam bentuk dict (bukan cuma konstanta) supaya kode di
# app.py yang sudah terbiasa lookup lewat dict ini tidak perlu berubah
# bentuk, walau isinya sekarang cuma satu entry.
GUDANG_SOURCES = {
    GUDANG_SOURCE_KEY: {"label": GUDANG_LABEL, "target_sheet": GUDANG_TARGET_SHEET},
}

# Header baku yang ditulis di baris 1 tiap tab tujuan (kolom A s/d V).
# Kolom R (index ke-18) sengaja dikosongkan sebagai pemisah.
GUDANG_HEADER = [
    "NO_INV",
    "AREA",
    "CUSTOMER",
    "UKURAN_PRODUK",
    "PRODUK",
    "SISA_STOCK_AWAL",
    "BARANG_MASUK_U/KIRIM",
    "BARANG_MASUK_U/STOCK",
    "BARANG_KELUAR_U/KIRIM",
    "BARANG_KELUAR_SISA_STOCK",
    "SISA_STOCK_AKHIR",
    "BERAT_ROLL",
    "BERAT_TOTAL",
    "JO_DAN_STATUS",
    "STATUS_BARANG",
    "JENIS_BARANG",
    "KETERANGAN",
    "",  # kolom R sengaja dikosongkan
    "JO",
    "STATUS",
    "TANGGAL_MASUK",
    "FINISH/PARSIAL",
]


def _get_calamine():
    """Import python_calamine secara lazy (baru dipakai kalau fitur Gudang
    benar-benar dipanggil), supaya modul ini tetap bisa di-import di
    lingkungan yang belum/tidak install library itu."""
    try:
        from python_calamine import CalamineWorkbook
    except ImportError as e:
        raise RuntimeError(
            "Library 'python_calamine' belum terinstall. Jalankan: "
            "pip install python-calamine"
        ) from e
    return CalamineWorkbook


def _sanitize_cell_gudang(cell):
    """Sama seperti sanitasi_nilai() di script import_gudang.py lama --
    ubah tipe data yang tidak bisa langsung dikirim ke Google Sheets API
    (mis. objek tanggal/waktu) jadi string biasa."""
    if cell is None:
        return ""
    if isinstance(cell, (datetime.datetime, datetime.date, datetime.time)):
        return cell.isoformat()
    return cell


def save_gudang_upload(file_storage, original_filename):
    """Dipanggil dari endpoint /api/gudang/upload. Terima file yang
    dikirim browser (werkzeug FileStorage, dari request.files), SIMPAN ke
    GUDANG_UPLOAD_DIR di server (bukan baca dari path lokal user lagi),
    lalu balikin daftar nama sheet di dalamnya supaya user bisa langsung
    pilih sheet tanpa upload ulang.

    Nama file di server SENGAJA dibuat tetap ("current" + ekstensi asli),
    bukan nama asli file -- supaya tiap kali user upload ulang (misal file
    minggu ini menggantikan minggu lalu), file lama otomatis tertimpa dan
    tidak menumpuk di disk server."""
    if file_storage is None:
        raise ValueError("Tidak ada file yang diupload.")

    original_filename = (original_filename or "").strip()
    ext = Path(original_filename).suffix.lower() or ".xlsx"
    if ext not in (".xlsx", ".xls"):
        raise ValueError("Format file harus .xlsx atau .xls.")

    GUDANG_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    saved_filename = f"current{ext}"
    filepath = GUDANG_UPLOAD_DIR / saved_filename
    file_storage.save(str(filepath))

    CalamineWorkbook = _get_calamine()
    wb = CalamineWorkbook.from_path(str(filepath))
    sheet_names = list(wb.sheet_names)

    update_source(
        GUDANG_SOURCE_KEY,
        filename=saved_filename,
        original_filename=original_filename,
        sheet_name=None,  # direset -- user harus pilih ulang sheet dari file baru ini
        uploaded_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    return sheet_names


def save_gudang_selection(filename, sheet_name):
    """Dipanggil dari kartu "Data Gudang" di halaman Input Data Produksi
    setelah user selesai upload file (save_gudang_upload()) lalu pilih
    sheet dari daftar yang dibalikin. HANYA menyimpan pilihan sheet ke
    config.json -- TIDAK menjalankan import. Eksekusi import-nya baru
    terjadi saat tombol Refresh di halaman Data Gudang BJB/BJL diklik
    (lihat run_gudang_import())."""
    filename = (filename or "").strip()
    sheet_name = (sheet_name or "").strip()
    if not filename or not sheet_name:
        raise ValueError("File dan sheet wajib dipilih.")

    _, src = get_source(GUDANG_SOURCE_KEY)
    if src.get("filename") != filename:
        raise ValueError(
            "File yang dipilih sudah tidak sesuai dengan file yang terakhir "
            "diupload. Upload ulang filenya."
        )

    return update_source(
        GUDANG_SOURCE_KEY,
        sheet_name=sheet_name,
        last_connected=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )


def get_gudang_selection():
    """Baca filename/sheet_name yang sudah tersimpan (hasil
    save_gudang_upload() + save_gudang_selection()). Dipakai
    run_gudang_import() dan endpoint status di app.py."""
    _, src = get_source(GUDANG_SOURCE_KEY)
    return {
        "filename": src.get("filename"),
        "sheet_name": src.get("sheet_name"),
    }


def _read_gudang_excel(filename, sheet_name):
    CalamineWorkbook = _get_calamine()
    filepath = GUDANG_UPLOAD_DIR / filename
    if not filepath.is_file():
        raise FileNotFoundError(
            f"File '{filename}' tidak ditemukan di server. Upload ulang lewat "
            "kartu \"Data Gudang\" di Input Data Produksi."
        )
    wb = CalamineWorkbook.from_path(str(filepath))
    if sheet_name not in wb.sheet_names:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan di file ini. Sheet yang tersedia: {wb.sheet_names}"
        )
    ws = wb.get_sheet_by_name(sheet_name)
    rows = ws.to_python()

    data = []
    for row in rows:
        if any(cell is not None and str(cell).strip() != "" for cell in row):
            data.append([_sanitize_cell_gudang(c) for c in row])
    return data


def _ratakan_kolom_gudang(row, jumlah_kolom):
    """Samakan panjang satu baris data dengan jumlah kolom header. Kurang
    -> ditambal string kosong. Lebih -> dipotong."""
    row = list(row)
    if len(row) < jumlah_kolom:
        row = row + [""] * (jumlah_kolom - len(row))
    elif len(row) > jumlah_kolom:
        row = row[:jumlah_kolom]
    return row


def _run_gudang_classification(sh):
    """Jalankan classify_gudang_sheets.classify_sheet() untuk tiap sheet di
    classify_gudang_sheets.SOURCE_SHEETS (["BJB", "BJL"]), pakai spreadsheet
    handle `sh` yang SUDAH TERBUKA (bukan buka koneksi/kredensial baru
    lewat gspread.service_account() seperti di classify_gudang_sheets.main())
    -- ini valid karena classify_gudang_sheets.SPREADSHEET_ID sama persis
    dengan GUDANG_TARGET_ID, jadi memang satu spreadsheet yang sama dengan
    yang barusan ditulis tab "API"-nya.

    Tiap sheet diproses independen: kalau salah satu gagal (mis. sheet
    "BJB"/"BJL" belum ada), errornya dikumpulkan & dilaporkan, TIDAK
    menghentikan sheet yang satunya ataupun menggagalkan hasil penulisan
    tab "API" yang sudah berhasil sebelumnya.

    Return: list string error (kosong kalau semua sheet berhasil)."""
    errors = []
    for source_title in classify_gudang_sheets.SOURCE_SHEETS:  # ["BJB", "BJL"]
        try:
            classify_gudang_sheets.classify_sheet(sh, source_title)
        except Exception as e:
            errors.append(f"Klasifikasi '{source_title}': {e}")
    return errors


def run_gudang_import():
    """Baca sheet yang SUDAH DIPILIH user sebelumnya (lewat kartu "Data
    Gudang" di halaman Input Data Produksi -> save_gudang_selection()),
    lalu TIMPA tab "API" (GUDANG_TARGET_SHEET) di spreadsheet
    GUDANG_TARGET_ID: baris 1 = GUDANG_HEADER, baris 2 dst = data. SESUDAH
    itu, langsung lanjut jalankan klasifikasi BJB/BJL
    (_run_gudang_classification() -> classify_gudang_sheets.classify_sheet())
    supaya sheet "BJB_KATEGORI" / "BJL_KATEGORI" ikut ter-update tiap kali
    tombol Refresh ini diklik -- tidak perlu jalan manual terpisah lagi.
    Dipanggil dari tombol Refresh di halaman Data Gudang BJB *atau* BJL
    -- keduanya memicu proses yang sama persis.
    Beda dari run_gsheet_import/run_excel_import: sheet tujuan di-CLEAR
    lalu ditimpa (bukan delete+recreate), biar formatting tab yang
    sudah ada di spreadsheet tidak hilang.

    Return: dict {"rows_written": int, "classification_errors": [str, ...]}
    -- rows_written tetap dianggap sukses walau classification_errors
    tidak kosong (tab "API" sudah berhasil ditulis; klasifikasi BJB/BJL
    dianggap langkah tambahan, bukan syarat sukses/gagalnya refresh)."""
    sel = get_gudang_selection()
    filename, sheet_name = sel["filename"], sel["sheet_name"]
    if not filename or not sheet_name:
        raise ValueError(
            "Belum ada file/sheet yang dipilih. Upload dulu file & pilih sheet "
            "lewat kartu \"Data Gudang\" di halaman Input Data Produksi."
        )

    try:
        data = _read_gudang_excel(filename, sheet_name)
        data_rata = [_ratakan_kolom_gudang(row, len(GUDANG_HEADER)) for row in data]

        client = get_gspread_client()
        sh = _with_retry(client.open_by_key, GUDANG_TARGET_ID, label="open gudang target")

        try:
            worksheet = sh.worksheet(GUDANG_TARGET_SHEET)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(
                title=GUDANG_TARGET_SHEET, rows=max(len(data_rata) + 10, 100), cols=26
            )

        worksheet.clear()
        semua_baris = [GUDANG_HEADER] + data_rata
        worksheet.update(range_name="A1", values=semua_baris, value_input_option="RAW")

        rows_written = len(data_rata)

        # ---- Klasifikasi BJB/BJL (dipanggil otomatis, bagian dari 1x klik Refresh) ----
        classification_errors = _run_gudang_classification(sh)

        error_msg = " | ".join(classification_errors) if classification_errors else None
        set_import_result(GUDANG_SOURCE_KEY, "OK", rows_written=rows_written, error=error_msg)
        return {"rows_written": rows_written, "classification_errors": classification_errors}
    except Exception as e:
        set_import_result(GUDANG_SOURCE_KEY, "ERROR", rows_written=None, error=str(e))
        raise


# ============================================================
# VARIAN 5: IMPORT HEADER-BASED DARI FILE EXCEL UPLOAD LOKAL
# (form_st_2, val_2 -- kartu "Form Serah Terima 2" & "Validasi 2")
# ============================================================
# BEDA dari VARIAN 4 "UPDATE STOCK — SUMBER FILE" di bawah (yang baca blok
# kolom lebar 10 TANPA header): dua source ini header-nya SAMA PERSIS
# dengan sumber lama yang dipakai import_form_st.py & import_val.py
# (lihat TARGET_HEADERS & HEADER_KEYWORDS di kedua file itu) -- cuma
# sumbernya diganti dari link Google Sheet jadi file Excel yang DIUPLOAD
# lewat browser (mirip Data Gudang/Update Stock), bukan didownload dari
# Google Drive seperti VARIAN 2 (import_dry_X.py).
#
# Makanya logikanya di sini SENGAJA DISAMAKAN dengan
# import_excel_from_drive() (VARIAN 2): cari baris header pakai
# find_header_row(), petakan kolom pakai get_column_mapping(), ambil baris
# data pakai get_data_rows(), lalu tulis lewat _write_target() yang sama
# (hapus tab lama -> buat baru -> isi -> rapikan urutan tab) -- BEDANYA
# cuma file-nya dibaca dari LOCAL_EXCEL_UPLOAD_DIRS (upload lokal, lewat
# python_calamine) alih-alih di-download dari Google Drive (pandas).
# ============================================================

def _sanitize_cell_local_excel(cell):
    """Sama prinsipnya dengan _sanitize_cell_excel (VARIAN 2), disesuaikan
    dengan tipe data yang dibalikin python_calamine (bukan pandas)."""
    if cell is None or (isinstance(cell, str) and cell.strip() == ""):
        return ""
    if isinstance(cell, (datetime.datetime, datetime.date)):
        return cell.strftime("%d-%m-%Y")
    if isinstance(cell, datetime.time):
        return cell.strftime("%H:%M:%S")
    if isinstance(cell, datetime.timedelta):
        total_seconds = int(cell.total_seconds())
        h, m, s = total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60
        return f"{h:02d}:{m:02d}:{s:02d}"
    return cell


def import_local_excel_upload(source_key, target_id, sheets_to_import, target_sheet_name,
                               target_headers, header_keywords,
                               junk_keywords=DEFAULT_JUNK_KEYWORDS, header_min_matches=1):
    """Versi VARIAN 5 dari import_excel_from_drive() (VARIAN 2) -- baca
    dari file yang diupload lokal (LOCAL_EXCEL_UPLOAD_DIRS[source_key]),
    bukan didownload dari Google Drive."""
    upload_dir = LOCAL_EXCEL_UPLOAD_DIRS.get(source_key)
    if upload_dir is None:
        raise ValueError(f"source_key '{source_key}' tidak dikenal untuk upload Excel lokal.")

    _, src = get_source(source_key)
    filename = src.get("filename")
    card_label = LOCAL_EXCEL_CARD_LABELS.get(source_key, source_key)
    if not filename:
        raise ValueError(f"Source '{source_key}' belum ada file yang diupload (klik kartu \"{card_label}\" di halaman Input Data Produksi dulu).")

    filepath = upload_dir / filename
    if not filepath.is_file():
        raise FileNotFoundError(
            f"File '{filename}' tidak ditemukan di server. Upload ulang lewat "
            f"kartu \"{card_label}\" di halaman Input Data Produksi."
        )

    CalamineWorkbook = _get_calamine()
    wb = CalamineWorkbook.from_path(str(filepath))
    available_sheets = wb.sheet_names

    client = get_gspread_client()
    target_sp = _with_retry(client.open_by_key, target_id, label=f"open target {target_id}")

    # Urutkan per bulan (bukan urutan centang user), supaya hasil akhir
    # di sheet target selalu kronologis JAN -> DES -- sama seperti
    # import_excel_from_drive() (VARIAN 2).
    sheets_to_import = sorted(sheets_to_import, key=_sheet_month_sort_key)
    print("   🗓️ Urutan proses (per bulan): " + ", ".join(sheets_to_import))

    all_rows = [target_headers]
    for sheet_name in sheets_to_import:
        if sheet_name not in available_sheets:
            print(f"⚠️ Sheet '{sheet_name}' tidak ditemukan di file yang diupload, dilewati.")
            continue
        print(f"\n🔍 Memproses sheet: {sheet_name}")
        ws = wb.get_sheet_by_name(sheet_name)
        rows = ws.to_python()
        if not rows:
            print("   Sheet kosong, dilewati.")
            continue
        header_idx = find_header_row(rows, header_keywords, min_matches=header_min_matches)
        if header_idx is None:
            print("   ❌ Tidak ditemukan baris header, dilewati.")
            continue
        header_row = rows[header_idx]
        print(f"   ✅ Header ditemukan di baris {header_idx+1}")
        mapping = get_column_mapping(header_row, target_headers)
        data_rows = get_data_rows(rows, header_idx, junk_keywords=junk_keywords)
        print(f"   📊 Jumlah baris data valid: {len(data_rows)}")
        for row in data_rows:
            all_rows.append(_map_row(row, target_headers, mapping, _sanitize_cell_local_excel))

    print(f"\n📝 Menulis ke sheet tujuan '{target_sheet_name}'...")
    return _write_target(target_sp, target_sheet_name, all_rows, target_headers)


def run_local_excel_import(source_key, target_sheet_name, target_headers, header_keywords,
                            junk_keywords=DEFAULT_JUNK_KEYWORDS, header_min_matches=1):
    """Dipanggil dari import_form_st_2.py & import_val_2.py (juga dipanggil
    langsung dari gudang_refresh() di app.py, bagian dari 1x klik Refresh
    di halaman Data Gudang BJB/BJL -- lihat komentar di sana).

    BEDA dari run_excel_import (VARIAN 2, source_id = ID file di Google
    Drive): source di sini TIDAK punya source_id sama sekali, filenya
    sudah tersimpan LOKAL di server (LOCAL_EXCEL_UPLOAD_DIRS) hasil upload
    lewat kartu di halaman Input Data Produksi. target_id-nya juga dibaca
    dari config.json per-source (src['target_id']), BUKAN
    cfg['target_sheet_id'] (spreadsheet utama) -- form_st_2 & val_2
    menulis ke spreadsheet Monitor Bahan Baku yang terpisah."""
    cfg, src = get_source(source_key)
    sheets = src.get("sheets") or []
    target_id = src.get("target_id")
    card_label = LOCAL_EXCEL_CARD_LABELS.get(source_key, source_key)
    if not target_id:
        raise ValueError(f"config.json: sources.{source_key}.target_id belum diisi (ID spreadsheet tujuan).")
    if not sheets:
        raise RuntimeError(f"'{source_key}': belum ada sheet yang dicentang (klik kartu \"{card_label}\" di halaman Input Data Produksi).")
    try:
        rows_written = import_local_excel_upload(
            source_key, target_id, sheets, target_sheet_name, target_headers, header_keywords,
            junk_keywords=junk_keywords, header_min_matches=header_min_matches,
        )
        set_import_result(source_key, "OK", rows_written=rows_written)
        return rows_written
    except Exception as e:
        set_import_result(source_key, "ERROR", error=str(e))
        raise


# ============================================================
# UPDATE STOCK — SUMBER FILE (upload Excel langsung), ALTERNATIF dari
# sumber link spreadsheet lama yang dipakai run_update_stock_import()
# di atas. Polanya sengaja dibuat MIRIP Data Gudang (VARIAN 4 di atas):
# upload file lewat browser -> server simpan ke UPDATE_STOCK_UPLOAD_DIR
# -> deteksi nama sheet -> user centang sheet mana yang mau dipakai.
#
# Beda dari Data Gudang: sheet yang dicentang BISA LEBIH DARI SATU
# (kartu "Update Stock" makai modal "Pilih Sheet" generik yang sama
# dengan source link lain, checkbox multi-select -- lihat
# openSheetModal()/saveSheetSelection() & endpoint /api/produksi/sheets
# di app.py), makanya field-nya "sheets" (list), BUKAN "sheet_name"
# (single) seperti Data Gudang.
# ============================================================

def save_local_excel_upload(source_key, file_storage, original_filename):
    """Versi GENERIK dari save_update_stock_upload() lama -- dipakai
    ketiga kartu upload-file "blok kolom lebar 10" (Update Stock, Form
    Serah Terima 2, Validasi 2). `source_key` harus salah satu key di
    LOCAL_EXCEL_UPLOAD_DIRS supaya tahu folder mana yang dipakai.

    Terima file yang dikirim browser (werkzeug FileStorage), simpan ke
    folder upload source ini di server (nama file dibuat tetap,
    "current"+ekstensi asli, supaya upload berikutnya otomatis menimpa
    file lama -- sama seperti save_gudang_upload()), lalu balikin daftar
    nama sheet di dalamnya supaya user bisa langsung centang lewat
    modal "Pilih Sheet", tanpa upload ulang."""
    if source_key not in LOCAL_EXCEL_UPLOAD_DIRS:
        raise ValueError(f"source_key '{source_key}' tidak dikenal untuk upload blok kolom.")
    if file_storage is None:
        raise ValueError("Tidak ada file yang diupload.")

    original_filename = (original_filename or "").strip()
    ext = Path(original_filename).suffix.lower() or ".xlsx"
    if ext not in (".xlsx", ".xls"):
        raise ValueError("Format file harus .xlsx atau .xls.")

    upload_dir = LOCAL_EXCEL_UPLOAD_DIRS[source_key]
    upload_dir.mkdir(parents=True, exist_ok=True)
    saved_filename = f"current{ext}"
    filepath = upload_dir / saved_filename
    file_storage.save(str(filepath))

    CalamineWorkbook = _get_calamine()
    wb = CalamineWorkbook.from_path(str(filepath))
    sheet_names = list(wb.sheet_names)

    update_source(
        source_key,
        input_mode="file",
        filename=saved_filename,
        original_filename=original_filename,
        source_id=None,  # bukan lagi mode link, kosongkan biar status lama tidak menyesatkan
        source_name=original_filename,
        sheets=[],  # direset -- user harus centang ulang dari file baru ini
        uploaded_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    return sheet_names


def save_update_stock_upload(file_storage, original_filename):
    """Alias lama, dipertahankan untuk kompatibilitas -- lihat
    save_local_excel_upload()."""
    return save_local_excel_upload("update_stock", file_storage, original_filename)


def _stringify_excel_cell(cell):
    """Ubah 1 sel hasil python_calamine jadi teks dengan format yang
    SAMA seperti kalau dibaca dari Google Sheets (worksheet.get_all_values()
    selalu balikin string), supaya pindah dari sumber link ke sumber file
    upload TIDAK mengubah isi data yang ditulis ke tujuan: angka bulat
    ditulis tanpa ".0" (mis. "27", bukan "27.0"), sel kosong jadi string
    kosong. _extract_stacked_blocks() sendiri juga bungkus tiap sel
    dengan str(...), jadi di sini cukup pastikan hasil str()-nya sudah
    benar duluan."""
    cell = _sanitize_cell_gudang(cell)
    if cell is None:
        return ""
    if isinstance(cell, float) and cell.is_integer():
        return str(int(cell))
    return str(cell)


LOCAL_EXCEL_CARD_LABELS = {
    "update_stock": "Update Stock",
    "form_st_2": "Form Serah Terima 2",
    "val_2": "Validasi 2",
}


def _read_stacked_block_excel_sheet(source_key, filename, sheet_name):
    """Versi GENERIK dari _read_update_stock_excel_sheet() lama. Baca satu
    sheet dari file yang diupload lewat save_local_excel_upload(source_key, ...),
    balikin list-of-list string per baris -- format yang sama seperti
    worksheet.get_all_values() dari gspread, supaya bisa langsung dilempar
    ke _extract_stacked_blocks() tanpa perlu perlakuan berbeda dari mode
    link lama."""
    upload_dir = LOCAL_EXCEL_UPLOAD_DIRS.get(source_key)
    if upload_dir is None:
        raise ValueError(f"source_key '{source_key}' tidak dikenal untuk upload blok kolom.")

    CalamineWorkbook = _get_calamine()
    filepath = upload_dir / filename
    if not filepath.is_file():
        card_label = LOCAL_EXCEL_CARD_LABELS.get(source_key, source_key)
        raise FileNotFoundError(
            f"File '{filename}' tidak ditemukan di server. Upload ulang lewat "
            f"kartu \"{card_label}\" di halaman Input Data Produksi."
        )
    wb = CalamineWorkbook.from_path(str(filepath))
    if sheet_name not in wb.sheet_names:
        raise ValueError(
            f"Sheet '{sheet_name}' tidak ditemukan di file ini. Sheet yang tersedia: {wb.sheet_names}"
        )
    ws = wb.get_sheet_by_name(sheet_name)
    return [[_stringify_excel_cell(c) for c in row] for row in ws.to_python()]


def _read_update_stock_excel_sheet(filename, sheet_name):
    """Alias lama, dipertahankan untuk kompatibilitas -- lihat
    _read_stacked_block_excel_sheet()."""
    return _read_stacked_block_excel_sheet("update_stock", filename, sheet_name)


# ============================================================
# VARIAN 3: SINKRON KOLOM TANGGAL / JO / NAMA / HASIL SLITTING
# DI SHEET "Validasi"
# ============================================================
# BEDA dari 2 varian di atas: ini bukan import dari spreadsheet luar.
# Sumbernya adalah 2 tab yang SUDAH ADA di spreadsheet tujuan sendiri
# (SL_1 & JO_1, hasil run_all.py), jadi tidak baca config.json "sources"
# sama sekali. Dipanggil dari refresh_validasi_jo.py (refresh ke-3 di
# halaman Data Validasi, bareng import_val.py & import_form_st.py).
#
# Alur:
#   1. Sheet SL_1: kolom TANGGAL, SPK/JO (mis. "2346/2309"), HASIL_ROL,
#      METER/ROL. Ambil baris yang TANGGAL-nya >= tanggal 1 bulan
#      berjalan (dijalankan tiap kali tombol Refresh dipencet, jadi
#      otomatis ambil semua data bulan ini sampai hari ini).
#   2. Kelompokkan baris2 itu per (TANGGAL, angka BELAKANG kode JO) --
#      1 SPK/JO biasanya punya banyak baris ROLL di SL_1.
#   3. Cari NAMA: angka belakang kode JO itu dicocokkan ke sheet JO_1
#      kolom F (kode multi-segmen, mis. "JO/26/I/5/41" -> dicocokkan
#      juga cuma bagian belakangnya, "41"). Kalau ketemu, ambil sheet
#      JO_1 kolom G sebagai NAMA.
#   4. Hitung HASIL SLITTING dari kumpulan (HASIL_ROL, METER/ROL) di
#      kelompok tsb:
#        - Sub-kelompokkan per nilai METER/ROL ("panjang"), jumlahkan
#          HASIL_ROL di tiap sub-kelompok.
#        - MODUS = panjang dengan frekuensi (jumlah baris) TERBANYAK,
#          syaratnya frekuensi > 1 dan tidak seri/tie sama sub-kelompok
#          lain. Kalau ada modus: "{jumlah_modus} + {jumlah_lain}@{panjang_lain} + ..."
#          Kalau TIDAK ada modus (semua panjang cuma muncul 1x, atau
#          seri): semua ditulis "{jumlah}@{panjang} + ..." (tanpa ada
#          yang "polos" tanpa @).
#        - Kolom HASIL_ROL & METER/ROL boleh berupa teks atau angka,
#          koma maupun titik desimal -- diparse fleksibel.
#   5. Baris yang (TANGGAL, JO) SUDAH ADA di "Validasi": kolom E (HASIL
#      SLITTING) di-update/refresh, kolom lain TIDAK disentuh.
#      Baris yang BELUM ADA: ditambahkan (append) dengan kolom A/B/C/E
#      terisi, kolom D (ORDER) dikosongkan.
#   6. Kolom I (MATCH VAL_1): angka belakang kode JO (suffix, sama kayak
#      dipakai buat NAMA/ORDER/POTONGAN) dicocokkan ke sheet VAL_1 kolom
#      G (isinya juga kode JO multi-segmen, mis. "JO/26/VI/20/2326").
#      Suffix VAL_1 kolom G bisa nyangkut huruf di belakang angka (mis.
#      "456A"), yang dipakai cuma bagian angkanya ("456").
#      SEMUA baris VAL_1 yang suffix-nya cocok dikumpulkan (bukan cuma
#      baris terakhir) -- nilai di kolom I & J tiap baris itu (boleh
#      berupa angka polos, mis. "62", atau format gabungan macam HASIL
#      SLITTING, mis. "68+2@530M") dipecah jadi bagian2, lalu:
#        - semua bagian ANGKA POLOS (bukan "count@panjang") dijumlah
#          langsung jadi SATU total.
#        - bagian "count@panjang" dikelompokkan per panjang, count-nya
#          dijumlah per kelompok, ditulis "{jumlah}@{panjang}".
#        - huruf satuan yang nyangkut di belakang angka (mis. "M" pada
#          "530M") dibuang duluan sebelum diparse jadi angka.
#      Hasil akhir: total (kalau ada) diikuti tiap kelompok "N@panjang",
#      digabung " + ", ditulis ke kolom I "Validasi". Kalau suffix tidak
#      ketemu di VAL_1 sama sekali (atau semua baris kosong/"-"), kolom
#      I "Validasi" dikosongkan.
#   7. Kolom J (MATCH FORM_ST_1, header "FORM SERAH TERIMA"): sama
#      persis logikanya kayak poin 6, tapi sumbernya sheet FORM_ST_1:
#      suffix JO dicocokkan lewat kolom D (bukan kolom G), nilai yang
#      dikumpulkan & dijumlah dari kolom K (bukan kolom I & J).
# ============================================================

VALIDASI_SHEET_NAME = "Validasi"
SL_SOURCE_SHEET_NAME = "SL_1"
JO_SOURCE_SHEET_NAME = "JO_1"

# Batas awal tanggal SL_1 yang menentukan JO mana yang ikut disinkron
# ke Validasi (dipakai buat nentuin baris/JO apa saja yang muncul --
# lihat poin 3 di sync_validasi_header()). Kalau tidak ada SATU PUN
# baris SL_1 untuk suffix JO tersebut yang tanggalnya >= tanggal ini,
# JO itu tidak akan dimasukkan/diupdate ke Validasi sama sekali.
#
# PENTING: ini HANYA menyaring JO mana yang ikut diproses. Begitu
# suatu JO lolos syarat ini, HASIL_SLITTING/HASIL_SLIT_QTY (kolom
# E/G) tetap dihitung dari SEMUA baris SL_1 untuk JO itu di seluruh
# sheet (tidak dibatasi ke tanggal ini saja) -- sama seperti kolom
# I/J/K (VAL_1/FORM_ST_1/TOTAL) yang memang sudah dari dulu
# menjumlah seluruh sheet tanpa filter tanggal.
VALIDASI_SL_START_DATE = datetime.date(2026, 7, 10)

# Kolom A sheet Validasi = "TANGGAL" (dipakai buat update tanggal terbaru
# saat 1 JO yang sama muncul di lebih dari satu tanggal di SL_1).
#
# CATATAN: sejak sync_validasi_header() ditulis ulang jadi "clear + tulis
# semua baris dari nol" tiap refresh (bukan update_cells per-kolom lagi),
# konstanta VALIDASI_COL_* di bawah ini murni DOKUMENTASI urutan kolom
# (dipakai buat baca kode ini), tidak lagi dipakai langsung sebagai
# argumen ke gspread.Cell(...).
VALIDASI_COL_TANGGAL = 1  # 1-based, buat update_cells
# Kolom D sheet Validasi = "ORDER" (nilai diambil dari JO_1 kolom H,
# di-index by suffix JO yang sama dengan lookup NAMA/POTONGAN).
VALIDASI_COL_ORDER = 4  # 1-based, buat update_cells
# Kolom F sheet Validasi = "POTONGAN" (nilai diambil dari JO_1 kolom K,
# di-index by suffix JO yang sama dengan lookup NAMA).
VALIDASI_COL_POTONGAN = 6  # 1-based, buat update_cells
# Kolom G sheet Validasi = "HASIL SLIT (QTY)" -- angka utuh hasil konversi
# HASIL_SLITTING (kolom E) memakai POTONGAN (kolom F). Lihat
# _compute_hasil_slit_qty().
VALIDASI_COL_HASIL_SLIT_QTY = 7  # 1-based, buat update_cells
# Kolom E sheet Validasi = "HASIL SLITTING" (lihat VALIDASI_COLUMN_MAP di app.py).
VALIDASI_COL_HASIL_SLITTING = 5  # 1-based, buat update_cells
# Kolom H sheet Validasi = "HASIL BAG" -- diisi dari jumlah BAG_1 kolom L
# (hanya kalau POTONGAN kolom F == 1, lihat sync_validasi_header()).
VALIDASI_COL_HASIL_BAG = 8  # 1-based, buat update_cells

BAG_SOURCE_SHEET_NAME = "BAG_1"
BAG_COL_JO = 4      # kolom E, format 'xxxx/25/XII/23/5273' -- suffix terakhir
BAG_COL_HASIL = 11  # kolom L -- nilai yang dijumlahkan

# Kolom I sheet Validasi = hasil jumlah kolom I & J dari SEMUA baris
# VAL_1 yang suffix JO-nya cocok (bukan cuma 1 baris). Lihat poin 6 di
# komentar blok sync_validasi_header() di atas.
VALIDASI_COL_MATCH_VAL1 = 9  # 1-based, buat update_cells

VAL1_SOURCE_SHEET_NAME = "VAL_1"
VAL1_COL_JO = 6  # kolom G -- kode JO, dicocokkan lewat suffix (angka belakang)
VAL1_COL_I = 8   # kolom I -- dikumpulkan lalu dijumlah ke kolom I Validasi
VAL1_COL_J = 9   # kolom J -- dikumpulkan lalu dijumlah ke kolom I Validasi

# Kolom J sheet Validasi ("FORM SERAH TERIMA") = hasil jumlah kolom K
# dari SEMUA baris FORM_ST_1 yang suffix JO-nya cocok (logika sama
# persis kayak kolom I / VAL_1 di atas). Lihat poin 7 di komentar blok
# sync_validasi_header() di atas.
VALIDASI_COL_MATCH_FORM_ST1 = 10  # 1-based, buat update_cells

# Kolom K sheet Validasi ("TOTAL") = penjumlahan kolom I + kolom J.
# Kolom kosong dianggap 0. Angka polos dijumlah apa adanya; bagian
# 'count@panjang' dikonversi dulu jadi (count*panjang)/POTONGAN
# (kolom F) baru ditambahkan -- KECUALI potongan == 1, di situ bagian
# 'count@panjang' diabaikan total (tidak dihitung sama sekali).
# Lihat _compute_kolom_k_total().
VALIDASI_COL_TOTAL = 11  # 1-based, buat update_cells

# Kolom M sheet Validasi ("STATUS") = dicari OTOMATIS tiap refresh dari
# sheet "Revisi_Manual": kalau JO baris ini ketemu di sana (artinya sudah
# pernah diklik tombol "Status OK (Manual)" di halaman Rekap), kolom M
# ditulis "OK". Kalau tidak ketemu, dikosongkan. Sengaja TIDAK disimpan
# sebagai nilai statis di kolom M supaya tidak salah baris kalau posisi
# baris JO berubah -- sumber kebenarannya selalu sheet Revisi_Manual,
# dicocokkan ulang setiap kali sync_validasi_header() jalan.
VALIDASI_COL_STATUS = 13  # 1-based, buat update_cells

REVISI_MANUAL_SHEET_NAME = "Revisi_Manual"
REVISI_MANUAL_COL_JO = 1  # 0-based -- kolom B sheet Revisi_Manual

FORM_ST1_SOURCE_SHEET_NAME = "FORM_ST_1"
FORM_ST1_COL_JO = 3   # kolom D -- suffix JO (angka), dicocokkan ke kolom B Validasi
FORM_ST1_COL_K = 10   # kolom K -- dikumpulkan lalu dijumlah ke kolom J Validasi

# Posisi kolom (0-based) fallback di sheet SL_1, dipakai kalau nama
# headernya tidak ketemu persis (lihat _find_col_index di bawah).
SL_COL_TANGGAL_FALLBACK = 0    # kolom A: TANGGAL
SL_COL_JO_FALLBACK = 3         # kolom D: SPK/JO
SL_COL_HASIL_ROL_FALLBACK = 10  # kolom K: HASIL_ROL
SL_COL_METER_ROL_FALLBACK = 14  # kolom O: METER/ROL

# Posisi kolom (0-based) di sheet JO_1 -- sesuai TARGET_HEADERS di
# import_jo.py (index 5 = "JO", index 6 = "KEMASAN", index 7 = "ORDER",
# index 10 = "POTONGAN").
JO_COL_JO = 5        # kolom F
JO_COL_NAMA = 6      # kolom G
JO_COL_ORDER = 7     # kolom H
JO_COL_POTONGAN = 10  # kolom K


def _last_segment(code):
    """Ambil bagian PALING BELAKANG (setelah '/' terakhir) dari kode JO,
    mis. '2346/2309' -> '2309', 'JO/26/I/5/41' -> '41'."""
    if not code:
        return ""
    return str(code).strip().split("/")[-1].strip()


def _numeric_key(text):
    """Normalisasi angka buat pencocokan (biar '0456' dianggap sama
    dengan '456'). Kalau bukan angka murni, dipakai apa adanya sebagai
    fallback (dibandingkan sebagai teks, huruf besar)."""
    text = (text or "").strip()
    if text.isdigit():
        return int(text)
    return text.upper()


def _build_revisi_manual_keys(ws_revisi):
    """Baca semua JO (kolom B) di sheet 'Revisi_Manual', kembalikan set
    suffix key (_numeric_key(_last_segment(...)), sama seperti pencocokan
    JO di seluruh script ini) -- dipakai buat isi kolom M (STATUS) di
    Validasi: kalau suffix JO baris Validasi ketemu di set ini, STATUS
    ditulis 'OK'."""
    rows = _with_retry(ws_revisi.get_all_values, label="baca Revisi_Manual")
    keys = set()
    for row in rows[1:]:  # lewati header
        if len(row) <= REVISI_MANUAL_COL_JO:
            continue
        key = _numeric_key(_last_segment(row[REVISI_MANUAL_COL_JO]))
        if key != "":
            keys.add(key)
    return keys


def _numeric_key_prefix(text):
    """Sama seperti _numeric_key(), tapi toleran kalau ada huruf
    nyangkut DI BELAKANG angka pada suffix (mis. sheet VAL_1 kolom G
    yang segmen terakhirnya bisa '456A', bukan '456' murni) -- fokus
    cuma ke angka di depan, huruf di belakangnya diabaikan ('456A' jadi
    key yang sama dengan '456'). Kalau segmennya sama sekali tidak
    diawali angka, fallback ke _numeric_key() biasa (dibandingkan
    sebagai teks)."""
    text = (text or "").strip()
    m = re.match(r"\d+", text)
    if m:
        return int(m.group(0))
    return _numeric_key(text)


def _parse_date_flexible(text):
    """Parse tanggal dari berbagai kemungkinan format ('1-8-2026',
    '01-08-2026', '1/8/2026', '2026-08-01', dll) -> datetime.date.
    Return None kalau tidak berhasil di-parse sama sekali."""
    text = (text or "").strip()
    if not text or text == "-":
        return None
    for sep in ("-", "/"):
        parts = text.split(sep)
        if len(parts) == 3:
            try:
                a, b, c = parts
                if len(a) == 4:  # YYYY-MM-DD
                    y, m, d = int(a), int(b), int(c)
                else:  # D-M-YYYY / DD-MM-YYYY
                    d, m, y = int(a), int(b), int(c)
                    if y < 100:
                        y += 2000
                return datetime.date(y, m, d)
            except (ValueError, TypeError):
                continue
    return None


def _find_col_index(header_row, target_name):
    """Cari index kolom (0-based) di header_row yang namanya cocok
    (dinormalisasi lewat _norm) dengan target_name. None kalau tidak ketemu."""
    target_norm = _norm(target_name)
    for i, h in enumerate(header_row):
        if _norm(h) == target_norm:
            return i
    return None


def _parse_flexible_number(text):
    """Parse angka dari cell yang bisa berupa teks/angka, format Indonesia
    (titik = pemisah ribuan, koma = desimal, mis. '1.600' -> 1600,
    '37,5' -> 37.5, '150.000' -> 150000). Kosong atau '-' -> None
    (dianggap tidak ada data, bukan 0)."""
    if text is None:
        return None
    text = str(text).strip()
    if text == "" or text == "-":
        return None
    text = text.replace(" ", "")
    if "," in text and "." in text:
        # titik = pemisah ribuan, koma = desimal
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    elif "." in text:
        # tidak ada koma sama sekali -> titik yang ada pasti pemisah
        # ribuan (format id), BUKAN desimal.
        text = text.replace(".", "")
    try:
        return float(text)
    except ValueError:
        return None


def _format_number(value):
    """96.0 -> '96', 13.8 -> '13,8' (format angka Indonesia, tanpa
    desimal kalau bilangan bulat)."""
    if value is None:
        return "0"
    if float(value).is_integer():
        return str(int(value))
    formatted = f"{value:.2f}".rstrip("0").rstrip(".")
    return formatted.replace(".", ",")


def _format_number_precise(value, max_decimals=6):
    """Sama seperti _format_number, tapi TIDAK dibulatkan ke 2 desimal --
    dipakai buat HASIL SLIT (QTY) yang hasil baginya (count*panjang/potongan)
    harus tetap presisi apa adanya (mis. 44,7 bukan dibulatkan jadi 45)."""
    if value is None:
        return "0"
    if float(value).is_integer():
        return str(int(value))
    formatted = f"{value:.{max_decimals}f}".rstrip("0").rstrip(".")
    return formatted.replace(".", ",")


def _resolve_potongan(raw_text):
    """Nilai POTONGAN (dari JO_1 kolom K). Kosong/tidak bisa di-parse/
    hasilnya < 1 -> dianggap 1. Return (nilai_numerik, teks_ditampilkan)."""
    val = _parse_flexible_number(raw_text)
    if val is None or val < 1:
        val = 1.0
    return val, _format_number(val)


def _build_val1_terms_lookup(val1_rows):
    """Baca semua baris VAL_1, kembalikan dict suffix_key -> list teks
    mentah dari kolom I & J (SEMUA baris yang suffix-nya cocok
    dikumpulkan, bukan cuma baris terakhir). Sel kosong/"-" dilewati.
    Suffix-nya dinormalisasi pakai _numeric_key_prefix() (bukan
    _numeric_key() biasa) karena kolom G VAL_1 tidak menutup
    kemungkinan segmen terakhirnya kenapa huruf nyangkut di belakang
    angka (mis. '456A' tetap dicocokkan sebagai suffix '456')."""
    lookup = {}
    for row in val1_rows[1:]:  # lewati header
        jo_code = row[VAL1_COL_JO] if VAL1_COL_JO < len(row) else ""
        if not jo_code or not str(jo_code).strip() or str(jo_code).strip() == "-":
            continue
        suffix_key = _numeric_key_prefix(_last_segment(jo_code))
        if suffix_key == "":
            continue
        terms = lookup.setdefault(suffix_key, [])
        for col in (VAL1_COL_I, VAL1_COL_J):
            val = row[col] if col < len(row) else ""
            val = str(val).strip()
            if val and val != "-":
                terms.append(val)
    return lookup


def _build_form_st1_terms_lookup(form_st1_rows):
    """Baca semua baris FORM_ST_1, kembalikan dict suffix_key -> list
    teks mentah dari kolom K (SEMUA baris yang suffix-nya cocok
    dikumpulkan, bukan cuma 1 baris). Sel kosong/"-" dilewati.

    Beda dari VAL_1: suffix di sini diambil dari kolom D, yang
    (menurut definisi user) sudah berisi ANGKA suffix JO itu sendiri
    (bukan kode JO lengkap kayak 'JO/26/VII/30/2779') -- tapi tetap
    dilewatkan _last_segment() dulu (aman/no-op kalau tidak ada '/'
    sama sekali) baru _numeric_key_prefix() (toleran huruf nyangkut di
    belakang angka, mis. '2779A' -> suffix '2779'), biar konsisten
    sama cara VAL_1 dicocokkan."""
    lookup = {}
    for row in form_st1_rows[1:]:  # lewati header
        jo_code = row[FORM_ST1_COL_JO] if FORM_ST1_COL_JO < len(row) else ""
        if not jo_code or not str(jo_code).strip() or str(jo_code).strip() == "-":
            continue
        suffix_key = _numeric_key_prefix(_last_segment(jo_code))
        if suffix_key == "":
            continue
        terms = lookup.setdefault(suffix_key, [])
        val = row[FORM_ST1_COL_K] if FORM_ST1_COL_K < len(row) else ""
        val = str(val).strip()
        if val and val != "-":
            terms.append(val)
    return lookup


_UNIT_SUFFIX_RE = re.compile(r"[\d.,]+")


def _strip_unit_suffix(text):
    """Buang huruf satuan yang nyangkut di belakang angka, mis. '530M'
    -> '530', '37,5 M' -> '37,5'. Kalau tidak ada angka di depan sama
    sekali, dikembalikan apa adanya (dibersihkan spasi saja)."""
    text = (text or "").strip()
    m = _UNIT_SUFFIX_RE.match(text)
    return m.group(0) if m else text


def _combine_number_terms(terms):
    """terms: list teks mentah (dari kolom I & J VAL_1, atau kolom K
    FORM_ST_1 -- lihat _build_val1_terms_lookup() /
    _build_form_st1_terms_lookup()), boleh berisi angka polos ('62')
    atau format gabungan macam HASIL SLITTING ('68+2@530M'). Tiap teks
    dipecah dulu per '+', lalu:
      - bagian ANGKA POLOS dijumlah langsung jadi satu total.
      - bagian 'count@panjang' dikelompokkan per panjang (huruf satuan
        di belakang angka panjang dibuang dulu, mis. '530M' -> '530'),
        count-nya dijumlah per kelompok.
    Hasil: total (kalau ada & bukan nol) diikuti tiap kelompok
    '{jumlah}@{panjang}', digabung ' + '. String kosong kalau tidak
    ada satupun bagian yang berhasil diparse."""
    total = 0.0
    have_total = False
    length_groups = OrderedDict()  # panjang_text (bersih) -> sum count

    for raw in terms:
        for part in str(raw).split("+"):
            part = part.strip()
            if not part:
                continue
            if "@" in part:
                count_text, _, panjang_text = part.partition("@")
                count_val = _parse_flexible_number(_strip_unit_suffix(count_text))
                panjang_clean = _strip_unit_suffix(panjang_text) or panjang_text.strip()
                if count_val is None:
                    continue
                length_groups[panjang_clean] = length_groups.get(panjang_clean, 0.0) + count_val
            else:
                num_val = _parse_flexible_number(_strip_unit_suffix(part))
                if num_val is None:
                    continue
                total += num_val
                have_total = True

    parts_out = []
    if have_total:
        parts_out.append(_format_number(total))
    for panjang_text, sum_val in length_groups.items():
        parts_out.append(f"{_format_number(sum_val)}@{panjang_text}")

    return " + ".join(parts_out)


def _group_sl_matches(sl_matches):
    """Kelompokkan sl_matches per panjang_text, dan tentukan modus_key
    (kelompok frekuensi tertinggi, syarat >1 dan tidak seri). Dipakai
    bareng oleh _compute_hasil_slitting() dan _compute_hasil_slit_qty()
    supaya definisi "modus" selalu konsisten di kedua kolom."""
    groups = OrderedDict()  # panjang_text -> {"sum": float, "count": int}
    for k_val, panjang_text in sl_matches:
        g = groups.setdefault(panjang_text, {"sum": 0.0, "count": 0})
        g["sum"] += k_val
        g["count"] += 1

    modus_key = None
    if groups:
        max_count = max(g["count"] for g in groups.values())
        if max_count > 1:
            top = [k for k, g in groups.items() if g["count"] == max_count]
            if len(top) == 1:
                modus_key = top[0]
    return groups, modus_key


def _compute_hasil_slitting(sl_matches):
    """sl_matches: list berisi (k_value: float, panjang_text: str) dari
    semua baris SL_1 yang JO-nya cocok. Return string HASIL SLITTING
    sesuai aturan modus/non-modus (lihat komentar blok di atas)."""
    if not sl_matches:
        return "-"

    groups, modus_key = _group_sl_matches(sl_matches)
    if not groups:
        return "-"

    parts = []
    if modus_key is not None:
        parts.append(_format_number(groups[modus_key]["sum"]))
        for panjang_text, g in groups.items():
            if panjang_text == modus_key:
                continue
            parts.append(f"{_format_number(g['sum'])}@{panjang_text}")
    else:
        for panjang_text, g in groups.items():
            parts.append(f"{_format_number(g['sum'])}@{panjang_text}")

    return " + ".join(parts)


def _compute_hasil_slit_qty(sl_matches, potongan_num):
    """HASIL SLIT (QTY) kolom G: konversi string HASIL_SLITTING jadi 1
    angka utuh, pakai POTONGAN (kolom F) sebagai pembagi.

    - Bagian MODUS (angka polos di HASIL_SLITTING, mis. '42') dijumlah
      apa adanya, TIDAK dibagi potongan.
    - Tiap bagian 'count@panjang' (mis. '3@900') dikonversi jadi
      (count * panjang) / potongan, baru dijumlahkan ke bagian modus.
    - Hasil akhir TIDAK dibulatkan (lihat _format_number_precise).
    - potongan_num == 1 ditangani di pemanggil (langsung kosongkan kolom
      G), fungsi ini hanya dipanggil kalau potongan_num > 1.
    """
    if not sl_matches:
        return None

    groups, modus_key = _group_sl_matches(sl_matches)
    if not groups:
        return None

    total = 0.0
    for panjang_text, g in groups.items():
        if panjang_text == modus_key:
            total += g["sum"]
        else:
            panjang_val = _parse_flexible_number(panjang_text) or 0.0
            total += (g["sum"] * panjang_val) / potongan_num
    return total


def _compute_kolom_k_total(text_i, text_j, potongan_num):
    """Kolom K sheet Validasi ('TOTAL') = penjumlahan kolom I (MATCH
    VAL_1) + kolom J (FORM SERAH TERIMA), dikembalikan sebagai SATU
    angka (float), bukan string gabungan macam _combine_number_terms().

    text_i / text_j: string mentah dari kolom I & J Validasi (bisa
    kosong/'-', angka polos '334', atau gabungan '5+4@500').

    Aturan tiap teks dipecah per '+':
      - Kolom kosong / '-' dianggap 0 (tidak menyumbang apa-apa).
      - Bagian ANGKA POLOS (mis. '334', '18') dijumlah apa adanya.
      - Bagian 'count@panjang' (mis. '4@500') dikonversi dulu jadi
        (count * panjang) / potongan, baru ditambahkan ke total --
        KECUALI potongan_num == 1: di situ bagian 'count@panjang' ini
        diabaikan TOTAL (bukan dihitung jadi 0 lewat pembagian, tapi
        memang tidak diikutkan sama sekali ke total).

    Return None kalau kedua teks kosong / tidak ada satupun angka
    valid yang berhasil diparse (biar kolom K bisa dikosongkan oleh
    pemanggil, bukan ditulis '0')."""
    total = 0.0
    have_any = False
    for text in (text_i, text_j):
        text = (text or "").strip()
        if not text or text == "-":
            continue
        for part in text.split("+"):
            part = part.strip()
            if not part:
                continue
            if "@" in part:
                if potongan_num == 1:
                    # Potongan 1 -> bagian count@panjang diabaikan total.
                    continue
                count_text, _, panjang_text = part.partition("@")
                count_val = _parse_flexible_number(_strip_unit_suffix(count_text))
                panjang_val = _parse_flexible_number(_strip_unit_suffix(panjang_text))
                if count_val is None or panjang_val is None:
                    continue
                total += (count_val * panjang_val) / potongan_num
                have_any = True
            else:
                num_val = _parse_flexible_number(_strip_unit_suffix(part))
                if num_val is None:
                    continue
                total += num_val
                have_any = True
    return total if have_any else None


def sync_validasi_header():
    """Lihat penjelasan alur lengkap di komentar blok di atas."""
    cfg = load_config()
    target_id = cfg["target_sheet_id"]
    client = get_gspread_client()
    target_sp = _with_retry(client.open_by_key, target_id, label=f"open target {target_id}")

    ws_sl = target_sp.worksheet(SL_SOURCE_SHEET_NAME)
    ws_jo = target_sp.worksheet(JO_SOURCE_SHEET_NAME)
    ws_bag = target_sp.worksheet(BAG_SOURCE_SHEET_NAME)
    ws_val1 = target_sp.worksheet(VAL1_SOURCE_SHEET_NAME)
    ws_form_st1 = target_sp.worksheet(FORM_ST1_SOURCE_SHEET_NAME)
    ws_validasi = target_sp.worksheet(VALIDASI_SHEET_NAME)
    ws_revisi = target_sp.worksheet(REVISI_MANUAL_SHEET_NAME)

    # ---- 1. Bangun lookup NAMA, ORDER & POTONGAN dari JO_1 (angka belakang -> KEMASAN / ORDER / POTONGAN) ----
    jo_rows = _with_retry(ws_jo.get_all_values, label="baca JO_1")
    nama_lookup = {}
    order_lookup = {}
    potongan_lookup = {}
    max_col_jo1 = max(JO_COL_JO, JO_COL_NAMA, JO_COL_ORDER, JO_COL_POTONGAN)
    for row in jo_rows[1:]:  # lewati header
        if len(row) <= max_col_jo1:
            continue
        key = _numeric_key(_last_segment(row[JO_COL_JO]))
        if key == "":
            continue
        nama_val = row[JO_COL_NAMA]
        existing_nama = nama_lookup.get(key)
        # Simpan kecocokan PERTAMA yang punya isi. Kalau match pertama yang
        # ketemu untuk suffix ini kebetulan KEMASAN-nya kosong (tersimpan
        # sebagai "-"), jangan dikunci di situ -- biarkan match berikutnya
        # yang beneran ada isinya mengisi lookup, supaya NAMA tidak
        # kosong padahal datanya ada di baris lain dengan JO yang sama.
        if existing_nama is None or existing_nama in ("", "-"):
            nama_lookup[key] = nama_val

        # ORDER: sama persis logic-nya kayak NAMA -- simpan kecocokan
        # pertama yang punya isi, jangan dikunci ke value kosong/"-".
        order_val = row[JO_COL_ORDER]
        existing_order = order_lookup.get(key)
        if existing_order is None or existing_order in ("", "-"):
            order_lookup[key] = order_val

        # POTONGAN: sama, simpan kecocokan pertama yang punya angka valid
        # (>= 1 setelah diformat); kalau belum ada/masih kosong, biarkan
        # match berikutnya mengisi.
        if key not in potongan_lookup:
            potongan_lookup[key] = row[JO_COL_POTONGAN]
        else:
            existing_val = _parse_flexible_number(potongan_lookup[key])
            if existing_val is None or existing_val < 1:
                new_val = _parse_flexible_number(row[JO_COL_POTONGAN])
                if new_val is not None and new_val >= 1:
                    potongan_lookup[key] = row[JO_COL_POTONGAN]

    # ---- 1b. Bangun lookup HASIL BAG dari BAG_1 (angka belakang -> jumlah kolom L) ----
    # Beda dari NAMA/ORDER/POTONGAN (ambil kecocokan PERTAMA): di sini
    # semua baris BAG_1 yang suffix JO-nya cocok DIJUMLAHKAN.
    bag_rows = _with_retry(ws_bag.get_all_values, label="baca BAG_1")
    bag_lookup = {}  # suffix_key -> total (float)
    max_col_bag = max(BAG_COL_JO, BAG_COL_HASIL)
    for row in bag_rows[1:]:  # lewati header
        if len(row) <= max_col_bag:
            continue
        key = _numeric_key(_last_segment(row[BAG_COL_JO]))
        if key == "":
            continue
        val = _parse_flexible_number(row[BAG_COL_HASIL])
        if val is None:
            continue
        bag_lookup[key] = bag_lookup.get(key, 0.0) + val

    # ---- 1c. Bangun lookup kolom I (MATCH VAL_1) dari sheet VAL_1: suffix
    # kolom G -> list nilai kolom I & J dari SEMUA baris yang cocok ----
    val1_rows = _with_retry(ws_val1.get_all_values, label="baca VAL_1")
    val1_lookup = _build_val1_terms_lookup(val1_rows)

    # ---- 1d. Bangun lookup kolom J (MATCH FORM_ST_1) dari sheet FORM_ST_1:
    # suffix kolom D -> list nilai kolom K dari SEMUA baris yang cocok ----
    form_st1_rows = _with_retry(ws_form_st1.get_all_values, label="baca FORM_ST_1")
    form_st1_lookup = _build_form_st1_terms_lookup(form_st1_rows)

    # ---- 2. Baca SL_1, cari posisi kolom yang dibutuhkan lewat nama header ----
    sl_rows = _with_retry(ws_sl.get_all_values, label="baca SL_1")
    if not sl_rows:
        print("   Sheet SL_1 kosong.")
        return 0
    header = sl_rows[0]
    col_tanggal = _find_col_index(header, "TANGGAL")
    col_jo = _find_col_index(header, "SPK/JO")
    col_hasil_rol = _find_col_index(header, "HASIL_ROL")
    col_meter_rol = _find_col_index(header, "METER/ROL")
    if col_tanggal is None:
        col_tanggal = SL_COL_TANGGAL_FALLBACK
    if col_jo is None:
        col_jo = SL_COL_JO_FALLBACK
    if col_hasil_rol is None:
        col_hasil_rol = SL_COL_HASIL_ROL_FALLBACK
    if col_meter_rol is None:
        col_meter_rol = SL_COL_METER_ROL_FALLBACK
    max_col_needed = max(col_tanggal, col_jo, col_hasil_rol, col_meter_rol)

    # ---- 3. Kelompokkan SEMUA baris SL_1 (seluruh sheet, tidak dibatasi
    # tanggal) per suffix JO SAJA (bukan per tanggal) -- 1 JO yang sama
    # bisa muncul di beberapa tanggal/bulan di SL_1 (mis. sisa produksi
    # lanjut ke bulan berikutnya); semua baris itu digabung jadi SATU
    # baris di Validasi, ditampilkan di tanggal PALING BARU, dengan
    # HASIL_SLITTING dihitung dari gabungan SEMUA baris (seluruh sheet,
    # BUKAN dibatasi VALIDASI_SL_START_DATE) untuk JO tersebut.
    #
    # VALIDASI_SL_START_DATE hanya dipakai buat nentuin JO MANA SAJA
    # yang ikut diproses/ditampilkan (harus punya minimal 1 baris SL_1
    # dengan tanggal >= tanggal itu) -- bukan buat membatasi baris mana
    # yang ikut dijumlah ke HASIL_SLITTING/HASIL_SLIT_QTY.
    groups = OrderedDict()  # suffix_key -> {"jo_text":.., "sl_matches": [...], "latest_date": date}
    included_keys = set()  # suffix_key yang punya >=1 baris SL_1 tanggal >= VALIDASI_SL_START_DATE
    for row in sl_rows[1:]:
        if len(row) <= max_col_needed:
            continue
        jo_code = row[col_jo]
        if not jo_code or not str(jo_code).strip() or str(jo_code).strip() == "-":
            continue
        tgl_parsed = _parse_date_flexible(row[col_tanggal])
        if tgl_parsed is None:
            continue
        suffix_key = _numeric_key(_last_segment(jo_code))
        if suffix_key == "":
            continue

        g = groups.setdefault(suffix_key, {
            "jo_text": str(jo_code).strip(),
            "sl_matches": [],
            "latest_date": tgl_parsed,
        })
        if tgl_parsed >= g["latest_date"]:
            g["latest_date"] = tgl_parsed
            g["jo_text"] = str(jo_code).strip()  # pakai penulisan JO dari baris tanggal terbaru

        k_val = _parse_flexible_number(row[col_hasil_rol])
        o_val = _parse_flexible_number(row[col_meter_rol])
        if k_val is not None and o_val is not None:
            g["sl_matches"].append((k_val, _format_number(o_val)))

        if tgl_parsed >= VALIDASI_SL_START_DATE:
            included_keys.add(suffix_key)

    # Buang suffix_key yang SAMA SEKALI tidak punya baris >= tanggal
    # mulai -- tapi sl_matches yang sudah terkumpul di atas tetap dari
    # SELURUH sheet buat suffix_key yang lolos.
    groups = OrderedDict((k, v) for k, v in groups.items() if k in included_keys)

    if not groups:
        print("   Tidak ada baris SL_1 bulan ini yang perlu disinkron ke Validasi.")
        return 0

    # ---- 4. Baca baris Validasi yang sudah ada (cuma buat tahu berapa
    # banyak baris lama yang perlu dibersihkan -- lihat poin 5) ----
    existing_rows = _with_retry(ws_validasi.get_all_values, label="baca Validasi")
    old_last_row = len(existing_rows)  # termasuk header; row 1 = header

    # ---- 4b. Cari JO mana saja yang sudah "OK" manual (ada di sheet
    # Revisi_Manual) -- dipakai buat isi kolom M di langkah 5 ----
    revisi_keys = _build_revisi_manual_keys(ws_revisi)

    # ---- 5. Tulis ULANG semua baris dari nol: kolom A-K dan M dibersihkan
    # dulu (baris 2 ke bawah, header baris 1 tidak disentuh), lalu ditulis
    # fresh dari hasil hitung di atas. Kolom L (SELISIH) SENGAJA TIDAK
    # disentuh sama sekali karena isinya formula Google Sheets, bukan nilai
    # dari script ini -- kalau ikut dibersihkan, formulanya hilang.
    #
    # Ini beda dari versi lama (update_cells utk baris lama + append_rows
    # utk baris baru): sekarang SEMUA baris ditulis ulang tiap refresh,
    # jadi JO yang sudah tidak relevan lagi (misal ada revisi data di
    # sumbernya) tidak nyangkut sebagai baris basi. Kolom M juga jadi
    # SELALU dihitung ulang dari Revisi_Manual di sini, bukan disimpan
    # statis -- jadi tidak akan salah baris walau posisi baris JO
    # berubah antar-refresh.
    final_rows = []      # utk kolom A-K, 1 baris per JO
    status_values = []   # utk kolom M, sejajar index-nya dgn final_rows
    for suffix_key, g in groups.items():
        nama = nama_lookup.get(suffix_key, "-")
        order_val = order_lookup.get(suffix_key, "-")
        hasil_slitting = _compute_hasil_slitting(g["sl_matches"])
        tgl_text = g["latest_date"].strftime("%d/%m/%Y")

        match_val1 = _combine_number_terms(val1_lookup.get(suffix_key, []))
        match_form_st1 = _combine_number_terms(form_st1_lookup.get(suffix_key, []))

        potongan_num, potongan_text = _resolve_potongan(potongan_lookup.get(suffix_key))
        if potongan_num == 1:
            # Potongan 1 -> tidak perlu konversi apa-apa, kosongkan kolom G.
            hasil_slit_qty = ""
            # Potongan 1 -> ambil HASIL BAG dari jumlah BAG_1 kolom L.
            bag_sum = bag_lookup.get(suffix_key)
            hasil_bag = "" if bag_sum is None else _format_number(bag_sum)
        else:
            qty_val = _compute_hasil_slit_qty(g["sl_matches"], potongan_num)
            hasil_slit_qty = "" if qty_val is None else _format_number_precise(qty_val)
            # Potongan > 1 -> kolom H dikosongkan.
            hasil_bag = ""

        kolom_k_val = _compute_kolom_k_total(match_val1, match_form_st1, potongan_num)
        kolom_k_total = "" if kolom_k_val is None else _format_number_precise(kolom_k_val)

        final_rows.append([
            tgl_text, g["jo_text"], nama, order_val, hasil_slitting,
            potongan_text, hasil_slit_qty, hasil_bag, match_val1, match_form_st1,
            kolom_k_total,
        ])
        status_values.append(["OK" if suffix_key in revisi_keys else ""])

    new_last_row = 1 + len(final_rows)  # baris terakhir SETELAH ditulis ulang
    clear_last_row = max(old_last_row, new_last_row)

    if clear_last_row >= 2:
        _with_retry(
            ws_validasi.batch_clear,
            [f"A2:K{clear_last_row}", f"M2:M{clear_last_row}"],
            label="hapus isi lama Validasi (A-K, M)",
        )

    if final_rows:
        _with_retry(
            ws_validasi.update, f"A2:K{new_last_row}", final_rows,
            value_input_option="RAW", label="tulis ulang Validasi (A-K)",
        )
        _with_retry(
            ws_validasi.update, f"M2:M{new_last_row}", status_values,
            value_input_option="RAW", label="tulis ulang Validasi (M/STATUS)",
        )

    total = len(final_rows)
    ok_count = sum(1 for v in status_values if v[0] == "OK")
    print(f"   ✅ {total} baris ditulis ulang di Validasi ({ok_count} di antaranya berstatus OK dari Revisi_Manual).")
    return total
