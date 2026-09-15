"""
KSA System - Backend API
Menggantikan google.script.run (Apps Script) dengan Flask + gspread,
supaya frontend (index.html) bisa baca-tulis ke Google Sheets lewat
REST API biasa.

Struktur Spreadsheet yang diharapkan (1 spreadsheet, beberapa sheet/tab).
Nama TAB harus persis (huruf besar/kecil ikut dicek), nama KOLOM di baris 1 juga harus persis:

  - Tab "Login"
      USERNAME | PASSWORD | NAMA | ROLE

  - Tab "Validasi"
      TANGGAL | JO | NAMA | ORDER | HASIL SLITTING | HASIL SLIT(QTY) |
      HASIL BAG | VALIDASI | FORM SERAH TERIMA | TOTAL | SELISIH | STATUS | POTONGAN

  - Tab "UpdateStock"
      JO | NAMA | ORDER | METER ORDER | METER VALIDASI |
      LAPISAN ORDER | LAPISAN VALIDASI | ACC

  - Tab "StockBahan"
      TANGGAL | USER | JO | NAMA | ORDER | METER ORDER | METER VALIDASI |
      LAPISAN ORDER | LAPISAN VALIDASI

  - Tab "PIC"
      NAMA | NOMOR

Kalau header di sheet kamu beda, cukup ubah nilai di *_COLUMN_MAP di bawah
(bagian kiri = nama kolom asli di sheet, bagian kanan = nama field yang
dipakai kode/frontend, jangan diubah bagian kanannya).

--------------------------------------------------------------------------
BAGIAN "INPUT DATA PRODUKSI" (baru)
--------------------------------------------------------------------------
Endpoint /api/produksi/* di bawah menggantikan cara lama isi
SOURCE_SHEET_ID / SHEETS_TO_IMPORT manual di tiap import_*.py. Sekarang:

  1. User paste link spreadsheet di kartu source (mis. "Printing 2") lalu
     klik Load -> /api/produksi/load -> deteksi ID + nama semua sheet/tab.
  2. User klik "Pilih Sheet" -> centang beberapa sheet dari hasil deteksi
     -> /api/produksi/sheets -> disimpan ke config.json.
  3. User klik satu tombol "Refresh Semua" -> /api/produksi/run-all ->
     menjalankan run_all.py di background thread (semua script import
     baca config.json sendiri-sendiri) -> frontend polling
     /api/produksi/run-status untuk lihat progress live.

Semua penyimpanan konfigurasi ada di config.json (lihat import_engine.py).

--------------------------------------------------------------------------
BAGIAN "FSTL — LAMPIRAN WASTE" (baru)
--------------------------------------------------------------------------
Endpoint /api/fstl/* pakai spreadsheet TERPISAH (FSTL_SPREADSHEET_ID, lihat
env var / default di bawah), bukan SPREADSHEET_ID utama:

  - /api/fstl/cek-jo      : cocokkan JO input ke sheet "JO_1" kolom F
                            (suffix setelah "/" terakhir, huruf nyangkut di
                            belakang angka diabaikan), balikin produk dari
                            kolom G.
  - /api/fstl/keterangan  : buat tiap proses yang dicentang (Printing, Dry
                            Laminasi, Slitting, Rewinding, Extrusi, Bag
                            Making), textjoin semua KETERANGAN yang JO-nya
                            cocok dari sheet sumbernya (lihat
                            FSTL_PROCESS_SOURCES), ditambah hasil dari sheet
                            "LP_1" (difilter kolom KLASIFIKASI) sebagai
                            "... LAPORAN PROD: ...".
  - /api/fstl/save        : simpan catatan waste ke sheet "{USERNAME}_Kitir"
                            (dibuat otomatis kalau belum ada, TANPA hapus
                            sheet lama), sebagai satu "kartu" mulai kolom B
                            baris 2 (kolom A & baris 1 TIDAK disentuh):
                              baris 1 (hijau) : SPK/JO | Produk | Waste Besar
                                                Proses : <daftar proses>
                              baris 2 (biru)  : label kolom (Waste Besar
                                                Proses/Keterangan/Action
                                                Plan/Status)
                              baris 3..N (hijau) : satu baris per proses yang
                                                dicentang
                            Kartu baru selalu disisipkan tepat di baris 2,
                            jadi kartu-kartu lama otomatis ikut turun tanpa
                            baris kosong pemisah.
"""

import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import gspread
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from google.oauth2.service_account import Credentials

import import_engine
import chatbot_engine

# Baca file .env (SPREADSHEET_ID, GOOGLE_CREDENTIALS_FILE, PORT,
# DEEPSEEK_API_KEY, dst) dan masukkan ke environment variable proses ini,
# supaya os.environ.get(...) di bawah bisa nemu nilainya. Kalau file .env
# nggak ada, ini nggak error -- cuma dianggap kosong (masih bisa jalan
# kalau env var-nya sudah di-set manual lewat "set" di CMD).
load_dotenv()

# --------------------------------------------------------------------------
# KONFIGURASI
# --------------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")  # isi di file .env

BASE_DIR = Path(__file__).resolve().parent
RUN_ALL_PATH = BASE_DIR / "run_all.py"

app = Flask(__name__)
CORS(app)  # izinkan dipanggil dari frontend berbeda origin (mis. Figma / GitHub Pages)


# --------------------------------------------------------------------------
# SERVE FRONTEND (index.html) -- tanpa ini, buka domain Render langsung
# bakal muncul "Not Found" 404 karena app.py aslinya cuma nyediain
# route /api/... saja (index.html sebelumnya dibuka manual dari komputer,
# bukan lewat server, jadi ini nggak ketahuan sampai di-deploy ke Render).
# --------------------------------------------------------------------------
@app.route("/")
def serve_index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/<path:filename>")
def serve_static_asset(filename):
    """Buat file pendukung frontend lain (css/js/gambar) kalau ada, yang
    ditaruh sejajar index.html dan direferensikan pakai path relatif."""
    return send_from_directory(BASE_DIR, filename)


import os as _os_diag  # noqa: E402  (cuma buat print PID di bawah, nggak ganggu import 'os' yang di atas)
print(f"=== SERVER STARTED (PID={_os_diag.getpid()}) — kalau baris ini muncul LAGI di tengah-tengah kamu testing, artinya server abis restart otomatis (cache ke-reset) ===", flush=True)


_gspread_client = None
_gspread_client_lock = threading.Lock()


def get_client():
    """Login ke Google (baca credentials.json + otorisasi) itu operasi yang
    lumayan berat kalau diulang tiap request. Sebelumnya dipanggil dari nol
    di SETIAP endpoint yang butuh Sheets -- termasuk /api/fstl/keterangan,
    jadi tiap klik "Ambil Keterangan" (JO sama ATAUPUN beda) selalu kena
    biaya login ulang ini duluan, sebelum sempat manfaatin cache sheet di
    bawah. Client login cuma dibuat SEKALI lalu dipakai ulang terus --
    aman, karena Credentials dari google-auth otomatis refresh token-nya
    sendiri kalau kadaluarsa, tanpa perlu login dari awal lagi."""
    global _gspread_client
    with _gspread_client_lock:
        if _gspread_client is None:
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            _gspread_client = gspread.authorize(creds)
        return _gspread_client


_spreadsheet_handle = {"sh": None}
_spreadsheet_lock = threading.Lock()


def _is_quota_error(e):
    """True kalau exception ini gspread.exceptions.APIError status 429
    (limit 'requests per minute per user' Google Sheets API kelampauan).
    Beda dari error lain (sheet nggak ada, credential salah, dst) yang
    memang harus langsung gagal -- 429 ini murni soal kebanyakan request
    dalam satu menit, jadi wajar buat dicoba lagi setelah nunggu sebentar."""
    if not isinstance(e, gspread.exceptions.APIError):
        return False
    try:
        status = e.response.status_code
    except AttributeError:
        status = None
    return status == 429 or "RESOURCE_EXHAUSTED" in str(e) or "Quota exceeded" in str(e)


def get_sheet(sheet_name, attempts=3):
    """Sama kayak _fstl_spreadsheet() di bawah -- open_by_key() itu
    panggilan ke Google (fetch metadata spreadsheet), dan ID-nya nggak
    pernah berubah selama app jalan. SEBELUMNYA dipanggil dari nol di
    SETIAP get_sheet(), jadi tool baru search_produk (yang buka hampir
    SEMUA sheet buat cari nama produk lintas grup, bisa belasan sheet
    sekaligus) jadi buka spreadsheet dari nol belasan kali cuma buat satu
    pertanyaan -- ini yang bikin request lambat/timeout ('Failed to
    fetch') waktu user tanya berdasarkan nama produk. Sekarang handle
    spreadsheet-nya dibuka sekali lalu dipakai ulang terus.

    Ditambah retry khusus buat 429 ('Quota exceeded ... per minute') --
    SEBELUMNYA sekali kena 429 (misalnya persis setelah tombol Refresh
    dipencet dan kuota per-menit abis) endpoint manapun yang manggil
    get_sheet() langsung crash jadi 500 tanpa dicoba ulang. Sekarang
    ditunggu sebentar dulu (kuotanya reset per menit) lalu dicoba lagi
    sebelum benar-benar dianggap gagal."""
    if not SPREADSHEET_ID:
        raise RuntimeError("SPREADSHEET_ID belum diset (lihat file .env)")
    with _spreadsheet_lock:
        sh = _spreadsheet_handle["sh"]
    if sh is None:
        client = get_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        with _spreadsheet_lock:
            _spreadsheet_handle["sh"] = sh

    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return sh.worksheet(sheet_name)
        except gspread.exceptions.APIError as e:
            if not _is_quota_error(e) or attempt == attempts:
                raise
            last_exc = e
            wait = 15 * attempt
            print(f"   ⚠️ Kena limit kuota Google Sheets (get_sheet '{sheet_name}'): {e}. Coba lagi dalam {wait}s...")
            time.sleep(wait)
    raise last_exc


# --------------------------------------------------------------------------
# 1. LOGIN
# --------------------------------------------------------------------------
LOGIN_COLUMN_MAP = {
    "USERNAME": "username",
    "PASSWORD": "password",
    "NAMA": "nama",
    "ROLE": "role",
}


def normalize_row(row, column_map):
    """Ubah key dari header asli sheet jadi key yang dipakai frontend,
    sekaligus tetap simpan key aslinya kalau-kalau dibutuhkan."""
    out = dict(row)  # simpan versi asli juga
    for original_key, new_key in column_map.items():
        if original_key in row:
            out[new_key] = row[original_key]
    return out


@app.route("/api/login", methods=["POST"])
def login():
    body = request.get_json(force=True) or {}
    username = str(body.get("username", "")).strip()
    password = str(body.get("password", "")).strip()

    ws = get_sheet("Login")
    raw = ws.get_all_records()  # baris pertama dianggap header
    rows = [normalize_row(r, LOGIN_COLUMN_MAP) for r in raw]

    for r in rows:
        if str(r.get("username", "")).strip() == username and str(r.get("password", "")).strip() == password:
            return jsonify({
                "status": "SUKSES",
                "nama": r.get("nama", ""),
                "role": r.get("role", ""),
            })

    return jsonify({"status": "GAGAL", "pesan": "Username atau password salah"}), 401


# --------------------------------------------------------------------------
# 2. DATA VALIDASI
# --------------------------------------------------------------------------
# Nama kolom ASLI di sheet -> nama field yang dipakai frontend.
# Sesuaikan bagian kiri kalau header di sheet kamu berubah.
VALIDASI_COLUMN_MAP = {
    "TANGGAL": "tanggal",
    "JO": "jo",
    "NAMA": "nama",
    "ORDER": "order",
    "HASIL SLITTING": "slitting",
    "HASIL SLIT\n(QTY)": "qtySlit",
    "HASIL BAG": "hasilBag",
    "VALIDASI": "validasi",
    "FORM SERAH TERIMA": "serahTerima",
    "TOTAL": "total",
    "SELISIH": "selisih",
    "STATUS": "status",
    "POTONGAN": "potongan",
}


@app.route("/api/validasi", methods=["GET"])
def get_validasi():
    ws = get_sheet("Validasi")
    # numericise_ignore=['all']: JANGAN biarkan gspread otomatis mengubah
    # cell yang keliatan seperti angka jadi int/float. Kita pakai format
    # angka Indonesia (titik = ribuan, koma = desimal) -- kalau dibiarkan,
    # gspread nganggep titik itu desimal (konvensi US) dan "160.000" jadi
    # kebaca 160.0, ditampilkan "160" di frontend. Ambil apa adanya (string).
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, VALIDASI_COLUMN_MAP) for r in raw]
    return jsonify(data)



# ---- Refresh sumber Data Validasi (VAL + Form Serah Terima + sinkron JO) ----
# Beda dari "Refresh Semua" di halaman Input Data Produksi (run_all.py,
# 15 script Printing/Rw/Sl/Dry/Sf/Ex/Bag/JO): ini cuma 3 script kecil, jadi
# dijalankan langsung (blocking) di request ini, tanpa background thread.
# - import_val.py & import_form_st.py : import dari spreadsheet luar (VAL_1 & FORM_ST_1)
# - refresh_validasi_jo.py            : sinkron internal SL_1 + JO_1 -> kolom
#                                        TANGGAL/JO/NAMA di tab "Validasi"
VALIDASI_IMPORT_SCRIPTS = ["import_val.py", "import_form_st.py", "refresh_validasi_jo.py"]


@app.route("/api/validasi/refresh-import", methods=["POST"])
def refresh_import_validasi():
    """Dipanggil tombol Refresh di halaman Data Validasi: jalankan
    import_val.py, import_form_st.py, lalu refresh_validasi_jo.py,
    supaya tab VAL_1 & FORM_ST_1 ter-update dan tab "Validasi" (kolom
    TANGGAL/JO/NAMA) tersinkron, sebelum data ditarik ulang lewat
    GET /api/validasi."""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    results = []
    for script_name in VALIDASI_IMPORT_SCRIPTS:
        script_path = BASE_DIR / script_name
        if not script_path.exists():
            results.append({"script": script_name, "status": "NOT FOUND", "log": ""})
            continue
        try:
            proc = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            ok = proc.returncode == 0
            results.append({
                "script": script_name,
                "status": "OK" if ok else f"FAILED (exit {proc.returncode})",
                "log": (proc.stdout or "") + (proc.stderr or ""),
            })
        except Exception as e:
            results.append({"script": script_name, "status": f"FAILED ({e})", "log": ""})

    success = all(r["status"] == "OK" for r in results)
    return jsonify({"success": success, "results": results})


@app.route("/api/validasi/status", methods=["POST"])
def update_status_rekap():
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    status = body.get("status", "")

    ws = get_sheet("Validasi")
    cell = ws.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan"}), 404

    header = ws.row_values(1)
    if "STATUS" not in header:
        return jsonify({"success": False, "message": "Kolom 'STATUS' tidak ada di sheet Validasi"}), 400

    col_status = header.index("STATUS") + 1
    ws.update_cell(cell.row, col_status, status)
    return jsonify({"success": True})


# ---- Tombol "Status OK (Manual)" di halaman Rekap ----
@app.route("/api/validasi/ok-manual", methods=["POST"])
def ok_manual_validasi():
    """Dipanggil tombol 'Status OK (Manual)' di tabel Rekap Data Selisih:
    1. Salin 1 baris JO dari sheet "Validasi" ke sheet "Revisi_Manual"
       (kalau JO itu belum pernah tersimpan di sana, supaya tidak dobel).
       Sheet "Revisi_Manual" inilah yang jadi SUMBER KEBENARAN status OK --
       tiap kali refresh_validasi_jo.py jalan (lihat sync_validasi_header()
       di import_engine.py), kolom STATUS di seluruh sheet Validasi
       dihitung ULANG dari sini (dicocokkan lewat JO, bukan nomor baris),
       jadi tidak akan salah baris walau posisi baris JO berubah antar-refresh.
    2. Set juga kolom STATUS di baris JO ini langsung jadi "OK", supaya
       Rekap langsung update seketika tanpa perlu tunggu tombol Refresh
       (loadRekap() di frontend sudah filter status != "OK")."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    if not jo:
        return jsonify({"success": False, "message": "JO wajib diisi"}), 400

    ws_val = get_sheet("Validasi")
    cell = ws_val.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan di Validasi"}), 404

    header = ws_val.row_values(1)
    row_values = ws_val.row_values(cell.row)

    ws_revisi = get_sheet("Revisi_Manual")
    if not ws_revisi.find(jo):
        ws_revisi.append_row(row_values)

    if "STATUS" in header:
        col_status = header.index("STATUS") + 1
        ws_val.update_cell(cell.row, col_status, "OK")

    return jsonify({"success": True})


# --------------------------------------------------------------------------
# 3. PIC (untuk dropdown kirim WA)
# --------------------------------------------------------------------------
PIC_COLUMN_MAP = {
    "NAMA": "nama",
    "NOMOR": "nomor",
}


@app.route("/api/pic", methods=["GET"])
def get_pic_list():
    ws = get_sheet("PIC")
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, PIC_COLUMN_MAP) for r in raw]
    return jsonify(data)  # [{"nama": ..., "nomor": ...}, ...]


# --------------------------------------------------------------------------
# 4. UPDATE STOCK (monitor bahan baku - butuh ACC)
# --------------------------------------------------------------------------
UPDATESTOCK_COLUMN_MAP = {
    "JO": "jo",
    "NAMA": "nama",
    "ORDER": "order",
    "METER ORDER": "meterOrder",
    "METER VALIDASI": "meterValidasi",
    "LAPISAN ORDER": "lapisanOrder",
    "LAPISAN VALIDASI": "lapisanValidasi",
    "KETERANGAN": "keterangan",
    "ACC": "acc",
}


@app.route("/api/update-stock/refresh-import", methods=["POST"])
def refresh_import_update_stock():
    """Dipanggil tombol Refresh di halaman 'Update Stock Bahan Baku'.
    Menjalankan DUA sinkronisasi terpisah (target spreadsheet beda,
    tidak saling menimpa):
      1. run_update_stock_import() -- baca sheet2 PL/PET/CPPM dst yang
         dicentang di kartu 'Update Stock' -> tulis ulang kolom B-K di
         tab tujuan pada spreadsheet EKSTERNAL Monitor Bahan Baku
         (kolom A di sana tidak disentuh).
      2. sync_update_stock_from_jo() -- baca JO_1 (spreadsheet utama)
         -> tulis kolom A/B/C/F tab 'UpdateStock' (spreadsheet utama)
         yang dipakai halaman ini sendiri lewat GET /api/update-stock.
    Kalau salah satu gagal, tetap coba jalankan yang satunya (supaya
    satu bagian yang error tidak ikut menggagalkan bagian lain), lalu
    laporkan errornya di response."""
    errors = []

    try:
        rows_written = import_engine.run_update_stock_import("update_stock")
        import_engine.set_import_result("update_stock", "OK", rows_written=rows_written, error=None)
    except Exception as e:
        rows_written = None
        import_engine.set_import_result("update_stock", "ERROR", rows_written=None, error=str(e))
        errors.append(f"run_update_stock_import: {e}")

    # Jeda sebelum lanjut ke sync_update_stock_from_jo() -- fungsi itu
    # LANGSUNG buka lagi spreadsheet eksternal Monitor Bahan Baku
    # (1lSj54tQP8QKMR96HiAHBCd1Zm-fOxnkdt9x3gD1tuEM) yang barusan ditulis
    # di atas (sheet UPDATE_STOCK, kolom B-K), buat baca ulang isinya
    # (_build_update_stock_monitor_lookup, kolom E/G). Tulis besar
    # (batch_clear + update) langsung disusul baca lagi ke spreadsheet
    # YANG SAMA dalam hitungan detik itu yang bikin gampang numpuk kena
    # limit "requests per minute" -- kasih jeda dulu di sini biar kuotanya
    # sempat longgar sebelum dipakai lagi.
    time.sleep(20)

    try:
        jo_rows_synced = import_engine.sync_update_stock_from_jo()
    except Exception as e:
        jo_rows_synced = None
        errors.append(f"sync_update_stock_from_jo: {e}")

    if errors:
        return jsonify({
            "success": False,
            "message": " | ".join(errors),
            "rows_written": rows_written,
            "jo_rows_synced": jo_rows_synced,
        }), 400

    return jsonify({
        "success": True,
        "rows_written": rows_written,
        "jo_rows_synced": jo_rows_synced,
    })


@app.route("/api/update-stock", methods=["GET"])
def get_update_stock():
    ws = get_sheet("UpdateStock")
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, UPDATESTOCK_COLUMN_MAP) for r in raw]
    for r in data:
        r["isLocked"] = str(r.get("acc", "0")) == "1"
    return jsonify(data)


@app.route("/api/update-stock/acc", methods=["POST"])
def acc_update_stock():
    """Tombol 'ACC & Kirim': kunci baris + salin data ke sheet StockBahan."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()

    ws_update = get_sheet("UpdateStock")
    cell = ws_update.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan"}), 404

    header = ws_update.row_values(1)
    col_acc = header.index("ACC") + 1 if "ACC" in header else None
    if col_acc:
        ws_update.update_cell(cell.row, col_acc, "1")

    ws_stock = get_sheet("StockBahan")
    ws_stock.append_row([
        datetime.now().strftime("%d-%m-%Y %H:%M"),
        body.get("user", "Tidak Diketahui"),
        body.get("jo", ""),
        body.get("nama", ""),
        body.get("order", ""),
        body.get("meterOrder", ""),
        body.get("meterValidasi", ""),
        body.get("lapisanOrder", ""),
        body.get("lapisanValidasi", ""),
    ])
    return jsonify({"success": True})


@app.route("/api/update-stock/unlock", methods=["POST"])
def unlock_update_stock():
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()

    ws = get_sheet("UpdateStock")
    cell = ws.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan"}), 404

    header = ws.row_values(1)
    if "ACC" in header:
        ws.update_cell(cell.row, header.index("ACC") + 1, "0")
    return jsonify({"success": True})


# --------------------------------------------------------------------------
# 5. STOCK BAHAN BAKU (hasil ACC)
# --------------------------------------------------------------------------
STOCKBAHAN_COLUMN_MAP = {
    "TANGGAL": "tanggal",
    "USER": "user",
    "JO": "jo",
    "NAMA": "nama",
    "ORDER": "order",
    "METER ORDER": "meterOrder",
    "METER VALIDASI": "meterValidasi",
    "LAPISAN ORDER": "lapisanOrder",
    "LAPISAN VALIDASI": "lapisanValidasi",
}


@app.route("/api/stock-bahan", methods=["GET"])
def get_stock_bahan():
    ws = get_sheet("StockBahan")
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, STOCKBAHAN_COLUMN_MAP) for r in raw]
    return jsonify(data)


# --------------------------------------------------------------------------
# 6. INPUT DATA PRODUKSI — Load link / Pilih Sheet / Refresh (Run All)
# --------------------------------------------------------------------------

@app.route("/api/produksi/sources", methods=["GET"])
def produksi_sources():
    """Daftar semua source (Printing 2..5, RW, SL, SF, Dry 1..5) beserta
    status koneksi & sheet yang sudah dicentang — dipakai untuk render kartu
    generik (link + checklist sheet). Source "gudang", "update_stock",
    "form_st_2", & "val_2" SENGAJA DIKECUALIKAN di sini karena alurnya beda
    (upload file, bukan link) dan punya kartu hardcoded sendiri di frontend
    (lihat index.html, kartu "Data Gudang", "Update Stock", "Form Serah
    Terima 2", "Validasi 2")."""
    cfg = import_engine.load_config()
    sources = cfg.get("sources", {})
    sources = {
        k: v for k, v in sources.items()
        if k not in import_engine.GUDANG_SOURCES and k not in import_engine.FILE_UPLOAD_EXTRA_SOURCES
    }
    return jsonify(sources)


@app.route("/api/produksi/load", methods=["POST"])
def produksi_load():
    """Body: {source_key, link}
    Ekstrak ID dari link, coba connect, deteksi nama semua sheet/tab,
    simpan source_id ke config.json. Sheet yang sudah pernah dicentang
    sebelumnya TIDAK dihapus otomatis, biar user bisa cocokkan ulang."""
    body = request.get_json(force=True) or {}
    source_key = str(body.get("source_key", "")).strip()
    link = str(body.get("link", "")).strip()

    if not source_key:
        return jsonify({"success": False, "message": "source_key wajib diisi"}), 400
    if not link:
        return jsonify({"success": False, "message": "Link/ID spreadsheet wajib diisi"}), 400

    try:
        import_engine.get_source(source_key)
    except KeyError as e:
        return jsonify({"success": False, "message": str(e)}), 404

    try:
        source_id = import_engine.extract_id_from_link(link)
    except ValueError as e:
        return jsonify({"success": False, "message": str(e)}), 400

    # Jenis sumber (Google Sheets asli vs file Excel/WPS di Drive) dideteksi
    # OTOMATIS lewat mimeType-nya di Drive API -- user tidak perlu pilih
    # manual lagi (dulu ada dropdown khusus di kartu "Rewind Kecil").
    try:
        source_type = import_engine.detect_source_type(source_id)
        detected_sheets, file_name = import_engine.detect_sheets(source_id, source_type)
    except Exception as e:
        return jsonify({"success": False, "message": f"Gagal connect ke spreadsheet: {e}"}), 400

    updated = import_engine.update_source(
        source_key,
        type=source_type,
        source_id=source_id,
        source_name=file_name,
        last_connected=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    # Simpan juga link mentah yang dipaste user ke sheet "ListMesin" (kolom B),
    # pada baris yang cocok dengan nama mesin source ini (kolom A). Kalau ini
    # gagal (mis. sheet ListMesin belum ada / nama mesin tidak match), jangan
    # sampai menggagalkan proses Load utama — cukup diabaikan.
    try:
        import_engine.update_list_mesin_link(source_key, link)
    except Exception:
        pass

    return jsonify({
        "success": True,
        "source_id": source_id,
        "file_name": file_name,
        "detected_sheets": detected_sheets,
        "selected_sheets": updated.get("sheets", []),
    })


@app.route("/api/produksi/sheets", methods=["POST"])
def produksi_sheets():
    """Body: {source_key, sheets: [...]}
    Simpan daftar sheet yang dicentang user untuk source ini -> menggantikan
    SHEETS_TO_IMPORT yang dulu hardcoded di tiap script."""
    body = request.get_json(force=True) or {}
    source_key = str(body.get("source_key", "")).strip()
    sheets = body.get("sheets")

    if not source_key:
        return jsonify({"success": False, "message": "source_key wajib diisi"}), 400
    if not isinstance(sheets, list):
        return jsonify({"success": False, "message": "sheets harus berupa list"}), 400

    try:
        import_engine.update_source(source_key, sheets=sheets)
    except KeyError as e:
        return jsonify({"success": False, "message": str(e)}), 404

    return jsonify({"success": True, "sheets": sheets})


# ---- Refresh / Run All (background, supaya 1 tombol tapi tidak nge-block) ----
RUN_STATE_LOCK = threading.Lock()
RUN_STATE = {
    "running": False,
    "log": "",
    "started_at": None,
    "finished_at": None,
    "returncode": None,
}


def _run_all_worker():
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        proc = subprocess.Popen(
            [sys.executable, str(RUN_ALL_PATH)],
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            bufsize=1,
        )
        for line in proc.stdout:
            with RUN_STATE_LOCK:
                RUN_STATE["log"] += line
        proc.wait()
        with RUN_STATE_LOCK:
            RUN_STATE["returncode"] = proc.returncode
    except Exception as e:
        with RUN_STATE_LOCK:
            RUN_STATE["log"] += f"\n[GAGAL MENJALANKAN run_all.py] {e}\n"
            RUN_STATE["returncode"] = -1
    finally:
        with RUN_STATE_LOCK:
            RUN_STATE["running"] = False
            RUN_STATE["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@app.route("/api/produksi/run-all", methods=["POST"])
def produksi_run_all():
    """Tombol Refresh tunggal: jalankan run_all.py di background thread.
    Frontend lalu polling /api/produksi/run-status untuk lihat progress."""
    with RUN_STATE_LOCK:
        if RUN_STATE["running"]:
            return jsonify({"success": False, "message": "Sedang berjalan, tunggu sampai selesai."}), 409
        RUN_STATE["running"] = True
        RUN_STATE["log"] = ""
        RUN_STATE["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        RUN_STATE["finished_at"] = None
        RUN_STATE["returncode"] = None

    thread = threading.Thread(target=_run_all_worker, daemon=True)
    thread.start()
    return jsonify({"success": True, "message": "run_all.py mulai dijalankan."})


@app.route("/api/produksi/run-status", methods=["GET"])
def produksi_run_status():
    with RUN_STATE_LOCK:
        return jsonify(dict(RUN_STATE))



# --------------------------------------------------------------------------
# 6a.5 REWIND KECIL — refresh TERPISAH, TIDAK ikut /api/produksi/run-all
# --------------------------------------------------------------------------
REWIND_KECIL_SCRIPT = BASE_DIR / "import_rewind_kecil.py"
REWIND_KECIL_RUN_STATE_LOCK = threading.Lock()
REWIND_KECIL_RUN_STATE = {
    "running": False,
    "log": "",
    "started_at": None,
    "finished_at": None,
    "returncode": None,
    "rows_written": None,
    "error": None,
}


def _run_rewind_kecil_worker():
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        proc = subprocess.Popen(
            [sys.executable, str(REWIND_KECIL_SCRIPT)],
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            bufsize=1,
        )
        for line in proc.stdout:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += line
        proc.wait()

        rows_written = None
        try:
            _, src_after = import_engine.get_source("rewind_kecil")
            rows_written = src_after.get("last_rows")
            err = src_after.get("last_error")
        except Exception:
            err = None

        with REWIND_KECIL_RUN_STATE_LOCK:
            REWIND_KECIL_RUN_STATE["returncode"] = proc.returncode
            REWIND_KECIL_RUN_STATE["rows_written"] = rows_written
            REWIND_KECIL_RUN_STATE["error"] = err
    except Exception as e:
        with REWIND_KECIL_RUN_STATE_LOCK:
            REWIND_KECIL_RUN_STATE["log"] += f"\n[GAGAL MENJALANKAN import_rewind_kecil.py] {e}\n"
            REWIND_KECIL_RUN_STATE["returncode"] = -1
            REWIND_KECIL_RUN_STATE["error"] = str(e)
    finally:
        with REWIND_KECIL_RUN_STATE_LOCK:
            REWIND_KECIL_RUN_STATE["running"] = False
            REWIND_KECIL_RUN_STATE["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@app.route("/api/produksi/rewind-kecil/run", methods=["POST"])
def produksi_run_rewind_kecil():
    """Refresh hanya Rewind Kecil. Tidak memanggil run_all.py."""
    with REWIND_KECIL_RUN_STATE_LOCK:
        if REWIND_KECIL_RUN_STATE["running"]:
            return jsonify({
                "success": False,
                "message": "Refresh Rewind Kecil sedang berjalan."
            }), 409

        REWIND_KECIL_RUN_STATE["running"] = True
        REWIND_KECIL_RUN_STATE["log"] = ""
        REWIND_KECIL_RUN_STATE["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        REWIND_KECIL_RUN_STATE["finished_at"] = None
        REWIND_KECIL_RUN_STATE["returncode"] = None
        REWIND_KECIL_RUN_STATE["rows_written"] = None
        REWIND_KECIL_RUN_STATE["error"] = None

    thread = threading.Thread(target=_run_rewind_kecil_worker, daemon=True)
    thread.start()

    return jsonify({
        "success": True,
        "message": "import_rewind_kecil.py mulai dijalankan."
    })


@app.route("/api/produksi/rewind-kecil/status", methods=["GET"])
def produksi_status_rewind_kecil():
    with REWIND_KECIL_RUN_STATE_LOCK:
        return jsonify(dict(REWIND_KECIL_RUN_STATE))


# --------------------------------------------------------------------------
# 6b. DATA GUDANG — file DIUPLOAD & sheet DIPILIH di kartu "Data Gudang"
# pada halaman Input Data Produksi (upload, save-selection), tapi
# DIEKSEKUSI (refresh) dari halaman Data Gudang BJB/BJL, TIDAK ikut
# run_all.py / tombol "Refresh Semua". Lihat import_engine.py bagian
# "VARIAN 4".
# --------------------------------------------------------------------------

@app.route("/api/gudang/sources", methods=["GET"])
def gudang_sources():
    """Status & pilihan folder/file/sheet yang tersimpan untuk source
    "Data Gudang" -- dipakai oleh kartu di Input Data Produksi (buat
    tahu apa yang sudah tersimpan) dan halaman Data Gudang BJB/BJL
    (buat nampilin status terakhir + sumber yang lagi aktif)."""
    cfg = import_engine.load_config()
    sources = cfg.get("sources", {})
    return jsonify({key: sources.get(key, {}) for key in import_engine.GUDANG_SOURCES})


@app.route("/api/gudang/upload", methods=["POST"])
def gudang_upload():
    """Multipart/form-data, field 'file'. Ganti dari cara lama (browse
    folder lokal di komputer user) -- server sekarang TIDAK PERNAH baca
    filesystem komputer user (tidak bisa, apalagi setelah di-deploy
    online), jadi browser yang kirim file-nya langsung lewat upload, baru
    server simpan & baca dari disknya sendiri. Dipanggil dari kartu "Data
    Gudang" di halaman Input Data Produksi begitu user pilih file lewat
    <input type="file">."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_gudang_upload(file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "filename": f"current{Path(file_storage.filename).suffix.lower() or '.xlsx'}", "sheets": sheets})


@app.route("/api/gudang/save-selection", methods=["POST"])
def gudang_save_selection():
    """Body: {filename, sheet}. Dipanggil setelah user selesai pilih
    sheet di kartu "Data Gudang" (Input Data Produksi), setelah file-nya
    diupload lewat /api/gudang/upload. HANYA menyimpan pilihan sheet ke
    config.json -- TIDAK menjalankan import."""
    body = request.get_json(force=True) or {}
    filename = str(body.get("filename", "")).strip()
    sheet = str(body.get("sheet", "")).strip()
    try:
        src = import_engine.save_gudang_selection(filename, sheet)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "source": src})


GUDANG_EXTRA_IMPORT_SCRIPTS = [
    ("form_st_2", "import_form_st_2.py"),
    ("val_2", "import_val_2.py"),
]


@app.route("/api/gudang/refresh", methods=["POST"])
def gudang_refresh():
    """Tombol Refresh di halaman Data Gudang BJB *atau* BJL -- keduanya
    memanggil endpoint yang sama ini. Tidak perlu body: folder/file/sheet
    dibaca dari config.json (hasil save-selection di kartu "Data Gudang",
    dan hasil upload di kartu "Form Serah Terima 2" / "Validasi 2").

    SATU klik di sini sekarang menjalankan TIGA proses (kalau salah satu
    gagal, yang lain TETAP dicoba -- sama polanya dengan
    refresh_import_validasi() di atas):
      1. run_gudang_import() -- tulis tab "API" (Monitor Bahan Baku) +
         klasifikasi BJB/BJL (classify_gudang_sheets.py). Dijalankan
         LANGSUNG (in-process, bukan subprocess) -- kalau ini gagal,
         seluruh refresh dianggap gagal (data Gudang BJB/BJL sendiri
         yang mau ditampilkan tidak ke-update).
      2. import_form_st_2.py -- kartu "Form Serah Terima 2" -> tab
         "FORM_ST_2" di spreadsheet Monitor Bahan Baku.
      3. import_val_2.py -- kartu "Validasi 2" -> tab "VAL_2" di
         spreadsheet yang sama.
    (2) & (3) dijalankan sebagai SUBPROCESS terpisah (python
    import_form_st_2.py / import_val_2.py) -- sama seperti
    VALIDASI_IMPORT_SCRIPTS di atas -- karena keduanya header-based
    (TARGET_HEADERS/HEADER_KEYWORDS ada di script masing2, lihat
    run_local_excel_import() di import_engine.py), bukan generik lewat
    satu fungsi seperti run_gudang_import()/run_update_stock_import().

    Kalau (1) Gudang gagal, response success=False. Kalau (1) berhasil
    tapi (2)/(3) ada yang gagal, response TETAP success=True (supaya
    tampilan Gudang BJB/BJL tetap ke-update) tapi errornya dikirim lewat
    'extra_errors' & detail per-script di 'extra_results'."""
    try:
        result = import_engine.run_gudang_import()
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    extra_results = []
    for source_key, script_name in GUDANG_EXTRA_IMPORT_SCRIPTS:
        script_path = BASE_DIR / script_name
        if not script_path.exists():
            extra_results.append({"source_key": source_key, "script": script_name, "status": "NOT FOUND", "rows_written": None, "log": ""})
            continue
        try:
            proc = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            ok = proc.returncode == 0
        except Exception as e:
            extra_results.append({"source_key": source_key, "script": script_name, "status": f"FAILED ({e})", "rows_written": None, "log": ""})
            continue

        # Baca status/rows_written TERBARU dari config.json (ditulis oleh
        # set_import_result() di dalam run_local_excel_import()) -- lebih
        # akurat daripada parsing stdout script.
        try:
            _, src_after = import_engine.get_source(source_key)
        except KeyError:
            src_after = {}

        extra_results.append({
            "source_key": source_key,
            "script": script_name,
            "status": "OK" if ok else f"FAILED (exit {proc.returncode})",
            "rows_written": src_after.get("last_rows"),
            "last_error": src_after.get("last_error") if not ok else None,
            "log": (proc.stdout or "") + (proc.stderr or ""),
        })

    extra_errors = [
        f"{r['source_key']}: {r.get('last_error') or r['status']}"
        for r in extra_results if r["status"] != "OK"
    ]

    return jsonify({
        "success": True,
        "rows_written": result["rows_written"],
        "classification_errors": result["classification_errors"],
        "extra_results": extra_results,
        "extra_errors": extra_errors,
    })


# --------------------------------------------------------------------------
# 6c. UPDATE STOCK — SUMBER FILE (upload Excel langsung di kartu "Update
# Stock" pada halaman Input Data Produksi, ALTERNATIF dari cara lama
# paste link spreadsheet). Alurnya mirip Data Gudang di atas, bedanya
# sheet yang dicentang bisa lebih dari satu -- makanya SETELAH upload,
# pemilihan sheet-nya lewat modal "Pilih Sheet" GENERIK yang sama dengan
# source link lain (endpoint /api/produksi/sheets, TIDAK butuh endpoint
# "save-selection" terpisah seperti Data Gudang).
# --------------------------------------------------------------------------

@app.route("/api/update-stock-source", methods=["GET"])
def update_stock_source():
    """Status & pilihan file/sheet yang tersimpan untuk source
    'update_stock' -- dipakai kartu "Update Stock" (mode upload file) di
    halaman Input Data Produksi. Dipisah dari /api/produksi/sources
    karena source ini dikecualikan dari daftar situ (lihat komentar di
    produksi_sources())."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}).get("update_stock", {}))


@app.route("/api/update-stock-source/upload", methods=["POST"])
def update_stock_source_upload():
    """Multipart/form-data, field 'file'. Sama alurnya seperti
    /api/gudang/upload, tapi untuk source 'update_stock' -- file
    disimpan ke server lalu dibalikin daftar nama sheet di dalamnya,
    supaya user bisa langsung centang sheet mana yang mau dipakai lewat
    modal "Pilih Sheet" (checkbox multi-select, sama dengan source link
    lain)."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_update_stock_upload(file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "sheets": sheets})


# --------------------------------------------------------------------------
# 6d. FORM SERAH TERIMA 2 & VALIDASI 2 — SUMBER FILE, alurnya PERSIS sama
# dengan "Update Stock" di 6c (upload Excel -> centang sheet lewat modal
# "Pilih Sheet" generik -> run_update_stock_import(source_key) ekstrak
# blok kolom lebar 10 & tumpuk). BEDA dari Update Stock: dua source ini
# TIDAK ikut "Refresh Semua" -- keduanya dijalankan otomatis bareng
# import Gudang tiap kali tombol Refresh di halaman Data Gudang BJB/BJL
# diklik (lihat gudang_refresh() di 6b, 3 proses sekali klik). Target
# tulisnya juga BUKAN spreadsheet Monitor Bahan Baku, tapi spreadsheet
# terpisah (lihat config.json: sources.form_st_2 / sources.val_2 ->
# target_id "1-ZyKSwXLzZaA6uNYRcpJNQZWX_ssYzvX45Z51xERipI", target_sheet
# "FORM_ST_2" / "VAL_2").
# --------------------------------------------------------------------------

@app.route("/api/form-st2-source", methods=["GET"])
def form_st2_source():
    """Status & pilihan file/sheet yang tersimpan untuk source 'form_st_2'
    -- dipakai kartu "Form Serah Terima 2" (mode upload file) di halaman
    Input Data Produksi."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}).get("form_st_2", {}))


@app.route("/api/form-st2-source/upload", methods=["POST"])
def form_st2_source_upload():
    """Multipart/form-data, field 'file'. Sama alurnya seperti
    /api/update-stock-source/upload, tapi untuk source 'form_st_2'."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_local_excel_upload("form_st_2", file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "sheets": sheets})


@app.route("/api/val2-source", methods=["GET"])
def val2_source():
    """Status & pilihan file/sheet yang tersimpan untuk source 'val_2'
    -- dipakai kartu "Validasi 2" (mode upload file) di halaman Input
    Data Produksi."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}).get("val_2", {}))


@app.route("/api/val2-source/upload", methods=["POST"])
def val2_source_upload():
    """Multipart/form-data, field 'file'. Sama alurnya seperti
    /api/update-stock-source/upload, tapi untuk source 'val_2'."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_local_excel_upload("val_2", file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "sheets": sheets})



# --------------------------------------------------------------------------
# 6d. WASTE REWIND — tampilan data dari spreadsheet REWIND_PY
# --------------------------------------------------------------------------
WASTE_REWIND_SPREADSHEET_ID = "1DnXtcMPkRdoadgO7ML7y7M9s4injmMTsKqEQ4BPrJxc"
WASTE_REWIND_SHEET_NAME = "REWIND_PY"

WASTE_REWIND_VISIBLE_HEADERS = [
    "JO",
    "SPK",
    "NO_JO",
    "Nama_Produk",
    "Planning_Order",
    "Planning_Meter",
    "Bahan_Awal_Printing_(Meter)",
    "Meter_Hilang_Rewind",
    "Hasil_Slitting_(Rol)",
    "Hasil_Slitting_(Meter)",
    "Waste_Slitting_Meter",
    "Persentase_Waste_(%)",
    "Waste_Slitting_After_Rewind_Meter",
    "Waste_Slitting_After_Rewind_Presentase",
]

_waste_rewind_spreadsheet_handle = {"sh": None}
_waste_rewind_spreadsheet_lock = threading.Lock()
_waste_rewind_cache = {"ts": 0.0, "headers": [], "rows": []}
_waste_rewind_cache_lock = threading.Lock()
_WASTE_REWIND_CACHE_TTL = int(os.environ.get("WASTE_REWIND_CACHE_TTL_SECONDS", "60"))


def _waste_rewind_spreadsheet():
    with _waste_rewind_spreadsheet_lock:
        if _waste_rewind_spreadsheet_handle["sh"] is not None:
            return _waste_rewind_spreadsheet_handle["sh"]

    client = get_client()
    sh = client.open_by_key(WASTE_REWIND_SPREADSHEET_ID)

    with _waste_rewind_spreadsheet_lock:
        _waste_rewind_spreadsheet_handle["sh"] = sh
    return sh


def _read_waste_rewind(force=False):
    now = time.time()
    with _waste_rewind_cache_lock:
        cached = dict(_waste_rewind_cache)
    if (
        not force
        and cached["headers"]
        and now - cached["ts"] < _WASTE_REWIND_CACHE_TTL
    ):
        return cached["headers"], cached["rows"]

    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()

    if not values:
        headers, rows = [], []
    else:
        headers = [str(x).strip() for x in values[0]]
        rows = []
        for row in values[1:]:
            padded = list(row) + [""] * max(0, len(headers) - len(row))
            row_obj = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                row_obj[header] = padded[i]
            if any(str(v).strip() for v in row_obj.values()):
                rows.append(row_obj)

    with _waste_rewind_cache_lock:
        _waste_rewind_cache["ts"] = now
        _waste_rewind_cache["headers"] = headers
        _waste_rewind_cache["rows"] = rows

    return headers, rows


@app.route("/api/waste-rewind", methods=["GET"])
def get_waste_rewind():
    """Ambil data REWIND_PY. Main table frontend boleh memakai
    WASTE_REWIND_VISIBLE_HEADERS, modal memakai seluruh headers."""
    force = str(request.args.get("refresh", "")).strip().lower() in ("1", "true", "yes")
    try:
        headers, rows = _read_waste_rewind(force=force)
        return jsonify({
            "success": True,
            "sheet": WASTE_REWIND_SHEET_NAME,
            "headers": headers,
            "visible_headers": WASTE_REWIND_VISIBLE_HEADERS,
            "rows": rows,
            "count": len(rows),
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"Gagal membaca {WASTE_REWIND_SHEET_NAME}: {e}"
        }), 500


# --------------------------------------------------------------------------
# 7. FSTL — LAMPIRAN WASTE
# --------------------------------------------------------------------------
# DUA spreadsheet berbeda dipakai di sini, jangan ketuker:
#
#   FSTL_SPREADSHEET_ID       -- spreadsheet SUMBER DATA (JO_1, LP_1,
#                                 PRINTING_2..5, DRY_1..5, SL_1, RW_1, EX_1,
#                                 BAG_1, dst -- lihat config.json). Dipakai
#                                 buat /api/fstl/cek-jo & /api/fstl/keterangan
#                                 (baik saat input JO baru MAUPUN saat user
#                                 klik "Ambil Keterangan" lagi pas revisi --
#                                 dua-duanya butuh data sumber yang sama).
#
#   FSTL_KITIR_SPREADSHEET_ID -- spreadsheet TEMPAT NYIMPEN KITIR user
#                                 (tab "{USERNAME}_Kitir"). Dipakai buat
#                                 /api/fstl/save, /api/fstl/list, dan
#                                 /api/fstl/revisi. SENGAJA dipisah dari
#                                 spreadsheet sumber di atas.
#
# Keduanya bisa dioverride lewat env var kalau suatu saat pindah lagi.
FSTL_SPREADSHEET_ID = os.environ.get(
    "FSTL_SPREADSHEET_ID", "1FRWpza_fa65jt8-n1-rN4rFFrfNBLixRxOLS_uUgYYU"
)
FSTL_KITIR_SPREADSHEET_ID = os.environ.get(
    "FSTL_KITIR_SPREADSHEET_ID", "1goadL7s6y2F38Zqgx65Z9Vy9mSVHkovEZDfDb16AgD0"
)

FSTL_HEADER = ["TANGGAL", "USER", "JO", "PRODUK", "PROSES", "KETERANGAN", "ACTION PLAN", "STATUS"]

# Nama proses (dari frontend, HARUS dibandingkan case-insensitive) -> sheet
# sumber yang dicari buat ambil KETERANGAN-nya.
#   prefixes = sheet yang NAMANYA DIAWALI salah satu prefix ini ikut dicari
#   exact    = sheet dengan nama PERSIS ini ikut dicari
FSTL_PROCESS_SOURCES = {
    "PRINTING": {"prefixes": ["PRINTING_"], "exact": []},
    "DRY LAMINASI": {"prefixes": ["DRY_"], "exact": ["SF_1"]},
    "SLITTING": {"prefixes": [], "exact": ["SL_1"]},
    "REWINDING": {"prefixes": [], "exact": ["RW_1"]},
    "EXTRUSI": {"prefixes": [], "exact": ["EX_1"]},
    "BAG MAKING": {"prefixes": [], "exact": ["BAG_1"]},
}

FSTL_COLOR_TITLE_LABEL = "#9bc2e6"  # biru — baris judul kartu (SPK/JO..) & baris label kolom
FSTL_COLOR_PROCESS = "#a9d08e"      # hijau — cuma sel nama proses di tiap baris data
FSTL_COLOR_WHITE = "#ffffff"        # putih — sel Keterangan/Action Plan/Status di baris data


def _fstl_hex_to_rgb01(hex_color):
    """'#a9d08e' -> {'red':.., 'green':.., 'blue':..} skala 0-1, format yang
    dipakai Google Sheets API buat backgroundColor lewat Worksheet.format()."""
    hex_color = hex_color.lstrip("#")
    return {
        "red": int(hex_color[0:2], 16) / 255,
        "green": int(hex_color[2:4], 16) / 255,
        "blue": int(hex_color[4:6], 16) / 255,
    }


FSTL_JO1_SHEET = "JO_1"
FSTL_JO1_COL_JO = 5      # kolom F (0-based)
FSTL_JO1_COL_PRODUK = 6  # kolom G (0-based)
FSTL_LP1_SHEET = "LP_1"


_fstl_spreadsheet_handle = {"sh": None}
_fstl_spreadsheet_lock = threading.Lock()

_fstl_kitir_spreadsheet_handle = {"sh": None}
_fstl_kitir_spreadsheet_lock = threading.Lock()


def _fstl_spreadsheet():
    """Handle ke spreadsheet SUMBER DATA (FSTL_SPREADSHEET_ID) -- open_by_key()
    juga panggilan ke Google, dan ID-nya nggak pernah berubah selama app
    jalan, jadi cukup dibuka sekali lalu handle-nya dipakai ulang terus,
    bukan dibuka lagi dari nol tiap ada request."""
    if not FSTL_SPREADSHEET_ID:
        raise RuntimeError("FSTL_SPREADSHEET_ID belum diset")
    with _fstl_spreadsheet_lock:
        if _fstl_spreadsheet_handle["sh"] is not None:
            return _fstl_spreadsheet_handle["sh"]
    client = get_client()
    sh = client.open_by_key(FSTL_SPREADSHEET_ID)
    with _fstl_spreadsheet_lock:
        _fstl_spreadsheet_handle["sh"] = sh
    return sh


def _fstl_kitir_spreadsheet():
    """Sama seperti _fstl_spreadsheet(), tapi buat spreadsheet TEMPAT NYIMPEN
    KITIR (FSTL_KITIR_SPREADSHEET_ID) -- spreadsheet BEDA dari sumber data.
    Handle-nya dipisah total (dict + lock sendiri) dari _fstl_spreadsheet()
    supaya nggak ketuker antara baca sumber data vs baca/tulis kitir user.

    CATATAN: service account yang dipakai get_client() harus sudah di-share
    (Editor) ke spreadsheet ini dulu, kalau belum open_by_key() bakal error
    permission (403)."""
    if not FSTL_KITIR_SPREADSHEET_ID:
        raise RuntimeError("FSTL_KITIR_SPREADSHEET_ID belum diset")
    with _fstl_kitir_spreadsheet_lock:
        if _fstl_kitir_spreadsheet_handle["sh"] is not None:
            return _fstl_kitir_spreadsheet_handle["sh"]
    client = get_client()
    sh = client.open_by_key(FSTL_KITIR_SPREADSHEET_ID)
    with _fstl_kitir_spreadsheet_lock:
        _fstl_kitir_spreadsheet_handle["sh"] = sh
    return sh


# --------------------------------------------------------------------------
# CACHE RINGAN buat endpoint /api/fstl/keterangan.
#
# Sheet sumber waste (PRINTING_2, SLITTING, dst) & LP_1 bisa ribuan-puluhan
# ribu baris (lihat config.json), dan sebelumnya di-fetch ULANG dari nol
# (sh.worksheets() + ws.get_all_values()) untuk SETIAP proses yang dicentang
# dalam satu request yang sama -- padahal daftar worksheet & isi LP_1 itu
# sama persis buat semua proses. Dua cache di bawah ini:
#   1. _fstl_worksheet_titles_cache : daftar nama tab (buat _fstl_matching_sheet_names)
#   2. _fstl_sheet_values_cache     : isi get_all_values() per nama sheet
# Keduanya di-share dalam SATU request (lewat parameter get_rows/titles yang
# dioper ke fungsi-fungsi di bawah), DAN juga disimpan lintas-request dengan
# TTL pendek supaya klik "Ambil Keterangan" berkali-kali dalam waktu dekat
# nggak perlu narik ulang sheet gede dari Google Sheets API. Data sumbernya
# sendiri cuma di-refresh berkala oleh proses import terpisah (lihat
# last_import di config.json), jadi cache basi beberapa menit aman -- durasinya
# bisa diatur lewat env var FSTL_CACHE_TTL_SECONDS tanpa ubah kode.
# --------------------------------------------------------------------------
_FSTL_CACHE_TTL_SECONDS = int(os.environ.get("FSTL_CACHE_TTL_SECONDS", "3600"))

# Sheet JO_1 dipakai buat "Periksa" JO (cek nama produk) waktu input JO baru.
# Sama kayak sheet sumber "Ambil Keterangan" (data yang dicari cuma ~1 bulan
# terakhir / data lama), jadi dikasih TTL 1 jam juga -- biar nggak keseringan
# loading ke Google Sheets. Ditulis eksplisit di override (bukan cuma ngandelin
# default di atas) supaya kalau nanti default umum di-ubah lagi lewat env var,
# JO_1 tetap punya nilainya sendiri yang bisa diatur terpisah.
_FSTL_JO1_CACHE_TTL_SECONDS = int(os.environ.get("FSTL_JO1_CACHE_TTL_SECONDS", "3600"))
_FSTL_SHEET_TTL_OVERRIDES = {FSTL_JO1_SHEET: _FSTL_JO1_CACHE_TTL_SECONDS}

_fstl_worksheet_titles_cache = {"ts": 0.0, "titles": None}
_fstl_sheet_values_cache = {}  # sheet_name -> (timestamp, rows)
_fstl_cache_lock = threading.Lock()


def _fstl_get_worksheet_titles(sh):
    """Daftar nama semua tab di spreadsheet FSTL, di-cache TTL pendek."""
    now = time.time()
    with _fstl_cache_lock:
        cached = _fstl_worksheet_titles_cache
        if cached["titles"] is not None and now - cached["ts"] < _FSTL_CACHE_TTL_SECONDS:
            return cached["titles"]
    titles = [w.title for w in sh.worksheets()]
    with _fstl_cache_lock:
        _fstl_worksheet_titles_cache["ts"] = now
        _fstl_worksheet_titles_cache["titles"] = titles
    return titles


def _fstl_get_sheet_values(sh, sheet_name):
    """get_all_values() satu sheet, di-cache TTL pendek per nama sheet.
    Dipanggil lewat helper ini supaya dalam SATU request /api/fstl/keterangan,
    sheet yang sama (mis. LP_1) cuma benar-benar di-fetch sekali walau
    dipakai berkali-kali (sekali per proses yang dicentang).

    TTL per-sheet bisa berbeda -- lihat _FSTL_SHEET_TTL_OVERRIDES (mis. JO_1
    dikasih TTL 1 jam karena datanya jauh lebih jarang berubah)."""
    ttl = _FSTL_SHEET_TTL_OVERRIDES.get(sheet_name, _FSTL_CACHE_TTL_SECONDS)
    now = time.time()
    with _fstl_cache_lock:
        cached = _fstl_sheet_values_cache.get(sheet_name)
        if cached is not None and now - cached[0] < ttl:
            return cached[1]
    try:
        rows = sh.worksheet(sheet_name).get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        rows = []
    with _fstl_cache_lock:
        _fstl_sheet_values_cache[sheet_name] = (now, rows)
    return rows


def _fstl_invalidate_cache():
    """Panggil ini setelah nulis data baru ke spreadsheet FSTL (mis. abis
    fstl_save nambah worksheet baru "{user}_Kitir"), biar cache worksheet
    titles nggak ketinggalan tab yang baru dibuat."""
    with _fstl_cache_lock:
        _fstl_worksheet_titles_cache["ts"] = 0.0
        _fstl_worksheet_titles_cache["titles"] = None


def _fstl_batch_prefetch_sheets(sh, sheet_names):
    """Ambil isi BANYAK sheet SEKALIGUS lewat SATU panggilan batch ke Google
    Sheets API (values.batchGet), taruh semuanya ke cache -- dipanggil di
    awal /api/fstl/keterangan SEBELUM proses satu-satu jalan.

    Ini beda dari cache TTL biasa di atas. Cache TTL cuma nyegah baca ULANG
    sheet yang SAMA berkali-kali. Tapi kalau user centang banyak proses
    sekaligus (mis. semua 6 proses), itu bisa nyentuh 10-15 sheet yang
    BEDA-BEDA, dan kalau belum ada satupun yang ke-cache (pertama kali buka
    app, atau cache-nya udah lewat batas TTL), sebelumnya tiap sheet itu
    dibaca lewat request TERPISAH ke Google Sheets API satu-satu secara
    berurutan. Selain kena latency jaringan berkali-kali, Google Sheets API
    juga punya BATAS JUMLAH REQUEST per menit per akun -- kalau kena limit
    itu, request berikutnya otomatis ditunda/di-retry, dan itu yang paling
    mungkin bikin kerasa lelet sampai hitungan menit walau data per sheet-nya
    sendiri nggak segede itu. Gabungin semua sheet yang dibutuhkan jadi SATU
    request besar (bukan banyak request kecil) menghindari masalah ini.

    valueRenderOption=UNFORMATTED_VALUE dipakai juga karena Google Sheets
    butuh waktu ekstra buat "merender" tampilan berformat (format tanggal,
    angka, dsb) tiap baca -- kolom yang kita pakai (JO/KETERANGAN/
    KLASIFIKASI) isinya teks biasa, jadi aman dilewatin proses render itu
    dan hasilnya sama, cuma lebih cepat didapat dari sisi server Google-nya."""
    now = time.time()
    with _fstl_cache_lock:
        to_fetch = [
            name for name in sheet_names
            if name not in _fstl_sheet_values_cache
            or now - _fstl_sheet_values_cache[name][0] >= _FSTL_SHEET_TTL_OVERRIDES.get(name, _FSTL_CACHE_TTL_SECONDS)
        ]
    if not to_fetch:
        return
    print(f"[FSTL] Cache MISS, ambil ulang dari Google Sheets: {to_fetch}", flush=True)
    ranges = [f"'{name}'" for name in to_fetch]
    try:
        resp = sh.values_batch_get(ranges, params={"valueRenderOption": "UNFORMATTED_VALUE"})
    except Exception:
        # Batch gagal (mis. salah satu range bermasalah) -> jangan sampai bikin
        # seluruh request error. Biarin _fstl_get_sheet_values() di bawah baca
        # satu-satu seperti biasa sebagai fallback (lebih lambat, tapi tetap jalan).
        return
    value_ranges = resp.get("valueRanges", []) if resp else []
    now = time.time()
    with _fstl_cache_lock:
        for name, vr in zip(to_fetch, value_ranges):
            values = vr.get("values", [])
            _fstl_sheet_values_cache[name] = (now, [[str(c) for c in row] for row in values])


def _fstl_suffix_key(jo_text):
    """Cocokkan cara ambil suffix JO sama persis kayak di import_engine:
    ambil segmen paling belakang setelah '/' (mis. '123/456A' -> '456A'),
    lalu ambil angka di depannya saja (huruf nyangkut di belakang
    diabaikan, jadi '456A' == '456')."""
    return import_engine._numeric_key_prefix(import_engine._last_segment(jo_text))


def _fstl_find_col(header_row, *keywords):
    """Cari index kolom (0-based) di header_row yang cocok sama salah satu
    keyword. Prioritas: EXACT MATCH keyword pertama di SEMUA kolom dulu,
    baru exact match keyword berikutnya, dst -- baru kalau tidak ada satupun
    exact match, fallback ke substring match dengan urutan prioritas yang
    sama. Ini penting karena beberapa sheet (mis. PRINTING_5) punya kolom
    "SPK" DAN "JO" terpisah -- kalau asal ambil kolom pertama yang
    mengandung salah satu keyword, bisa kepilih kolom yang salah (SPK
    kepilih duluan padahal yang dimaksud kolom JO)."""

    def _norm(text):
        return str(text or "").strip().upper().replace("_", "").replace(" ", "")

    normed_keywords = [_norm(k) for k in keywords if k]
    normed_header = [_norm(cell) for cell in header_row]

    for kw in normed_keywords:
        for idx, cell_norm in enumerate(normed_header):
            if cell_norm and cell_norm == kw:
                return idx

    for kw in normed_keywords:
        for idx, cell_norm in enumerate(normed_header):
            if cell_norm and kw in cell_norm:
                return idx

    return None


def fstl_lookup_produk(jo_raw):
    """Cocokkan JO input ke sheet JO_1 kolom F, kembalikan (produk, error).
    produk None + error None artinya JO tidak ketemu (bukan error sistem)."""
    sh = _fstl_spreadsheet()
    try:
        ws = sh.worksheet(FSTL_JO1_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        return None, f"Sheet '{FSTL_JO1_SHEET}' tidak ditemukan di spreadsheet FSTL"
    target_key = _fstl_suffix_key(jo_raw)
    if target_key == "":
        return None, "Format JO tidak valid"
    rows = _fstl_get_sheet_values(sh, FSTL_JO1_SHEET)
    for row in rows[1:]:
        jo_cell = row[FSTL_JO1_COL_JO] if len(row) > FSTL_JO1_COL_JO else ""
        if not str(jo_cell).strip():
            continue
        if _fstl_suffix_key(jo_cell) == target_key:
            produk = row[FSTL_JO1_COL_PRODUK] if len(row) > FSTL_JO1_COL_PRODUK else ""
            return (str(produk).strip() or None), None
    return None, None


def _fstl_matching_sheet_names(sh, prefixes, exact):
    names = _fstl_get_worksheet_titles(sh)
    matched = []
    for name in names:
        if name in exact:
            matched.append(name)
            continue
        if any(name.upper().startswith(p.upper()) for p in prefixes):
            matched.append(name)
    return matched


def _fstl_join_terms(terms):
    """textjoin semua keterangan yang cocok, buang kosong/'-', buang duplikat."""
    cleaned = []
    for t in terms:
        t = str(t).strip()
        if t and t != "-" and t not in cleaned:
            cleaned.append(t)
    return " | ".join(cleaned)


def _fstl_keterangan_rows(sh, sheet_name, target_key):
    rows = _fstl_get_sheet_values(sh, sheet_name)
    if not rows:
        return []
    header = rows[0]
    col_jo = _fstl_find_col(header, "JO", "SPK")
    col_ket = _fstl_find_col(header, "KETERANGAN")
    if col_jo is None or col_ket is None:
        return []
    terms = []
    for row in rows[1:]:
        jo_cell = row[col_jo] if len(row) > col_jo else ""
        if not str(jo_cell).strip():
            continue
        if _fstl_suffix_key(jo_cell) == target_key:
            terms.append(row[col_ket] if len(row) > col_ket else "")
    return terms


def _fstl_lp1_rows(sh, target_key, process_name):
    """LP_1 ambil pakai cara yang sama (suffix JO), tapi ditambah filter
    kolom KLASIFIKASI harus cocok sama proses yang lagi dicek.
    Kolom keterangan di sheet LP_1 headernya "Faktor_Penyebab_Waste"
    (BUKAN "KETERANGAN" seperti di sheet-sheet sumber proses) -- makanya
    dicari duluan, dengan "KETERANGAN" jadi fallback kalau ada versi LP_1
    lama yang headernya beda."""
    rows = _fstl_get_sheet_values(sh, FSTL_LP1_SHEET)
    if not rows:
        return []
    header = rows[0]
    col_jo = _fstl_find_col(header, "JO", "SPK")
    col_klas = _fstl_find_col(header, "KLASIFIKASI")
    col_ket = _fstl_find_col(header, "FAKTOR_PENYEBAB_WASTE", "KETERANGAN")
    if col_jo is None or col_ket is None:
        return []
    proc_norm = process_name.strip().upper()
    terms = []
    for row in rows[1:]:
        jo_cell = row[col_jo] if len(row) > col_jo else ""
        if not str(jo_cell).strip():
            continue
        if _fstl_suffix_key(jo_cell) != target_key:
            continue
        if col_klas is not None:
            klas_cell = str(row[col_klas] if len(row) > col_klas else "").strip().upper()
            if proc_norm not in klas_cell and klas_cell not in proc_norm:
                continue
        terms.append(row[col_ket] if len(row) > col_ket else "")
    return terms


def fstl_keterangan_for_process(sh, jo_raw, process_name):
    """Gabung keterangan dari sheet sumber proses + LP_1 (filter klasifikasi),
    jadi satu kalimat: '<keterangan proses> LAPORAN PROD: <keterangan LP_1>'."""
    target_key = _fstl_suffix_key(jo_raw)
    proc_key = process_name.strip().upper()
    proc_text = ""
    cfg = FSTL_PROCESS_SOURCES.get(proc_key)
    if cfg is not None:
        all_terms = []
        for name in _fstl_matching_sheet_names(sh, cfg["prefixes"], cfg["exact"]):
            all_terms.extend(_fstl_keterangan_rows(sh, name, target_key))
        proc_text = _fstl_join_terms(all_terms)
    lp1_text = _fstl_join_terms(_fstl_lp1_rows(sh, target_key, proc_key))
    if proc_text and lp1_text:
        return f"{proc_text} LAPORAN PROD: {lp1_text}"
    if lp1_text:
        return f"LAPORAN PROD: {lp1_text}"
    return proc_text


@app.route("/api/fstl/cek-jo", methods=["POST"])
def fstl_cek_jo():
    """Body: {jo}. Ambil JO belakang (abaikan huruf nyangkut), cocokkan ke
    sheet JO_1 kolom F, kembalikan produk dari kolom G."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    if not jo:
        return jsonify({"error": "JO wajib diisi"}), 400
    try:
        produk, err = fstl_lookup_produk(jo)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    if err:
        return jsonify({"error": err}), 404
    if not produk:
        return jsonify({"error": f"JO '{jo}' tidak ditemukan di sheet {FSTL_JO1_SHEET}"}), 404
    return jsonify({"produk": produk})


@app.route("/api/fstl/inspect-sheet", methods=["GET"])
def fstl_inspect_sheet():
    """DIAGNOSTIK — GET /api/fstl/inspect-sheet?sheet=PRINTING_5&jo=123/2254
    Balikin header sheet, index kolom JO/KETERANGAN yang berhasil dideteksi,
    dan contoh isi kolom JO (mentah + hasil parsing suffix-nya) biar gampang
    ketauan kenapa pencarian keterangan gak ketemu (nama header meleset,
    format JO beda, dll). Buka aja URL-nya lewat browser buat lihat hasilnya."""
    sheet_name = request.args.get("sheet", "").strip()
    jo = request.args.get("jo", "").strip()
    if not sheet_name:
        return jsonify({"error": "parameter 'sheet' wajib diisi, mis. ?sheet=PRINTING_5"}), 400
    try:
        sh = _fstl_spreadsheet()
        ws = sh.worksheet(sheet_name)
        rows = ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        return jsonify({"error": f"Sheet '{sheet_name}' tidak ditemukan di spreadsheet FSTL"}), 404
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    if not rows:
        return jsonify({"sheet": sheet_name, "error": "Sheet kosong (tidak ada baris sama sekali)"})

    header = rows[0]
    col_jo = _fstl_find_col(header, "JO", "SPK")
    col_ket = _fstl_find_col(header, "KETERANGAN")
    target_key = _fstl_suffix_key(jo) if jo else None

    sample_jo_values = []
    matched_rows = 0
    if col_jo is not None:
        for row in rows[1:]:
            cell = row[col_jo] if len(row) > col_jo else ""
            if not str(cell).strip():
                continue
            suffix_key = _fstl_suffix_key(cell)
            if len(sample_jo_values) < 8:
                sample_jo_values.append({
                    "raw": cell,
                    "last_segment": import_engine._last_segment(cell),
                    "suffix_key": suffix_key,
                })
            if target_key is not None and suffix_key == target_key:
                matched_rows += 1

    return jsonify({
        "sheet": sheet_name,
        "header_row": header,
        "col_jo_terdeteksi": {"index": col_jo, "nama_header": header[col_jo] if col_jo is not None else None},
        "col_keterangan_terdeteksi": {"index": col_ket, "nama_header": header[col_ket] if col_ket is not None else None},
        "total_baris_data": len(rows) - 1,
        "jo_yang_dicari": jo or None,
        "suffix_key_yang_dicari": target_key,
        "jumlah_baris_cocok": matched_rows,
        "contoh_isi_kolom_jo": sample_jo_values,
    })


@app.route("/api/fstl/keterangan", methods=["POST"])
def fstl_keterangan():
    """Body: {jo, processes:[nama_proses,...]}. Buat tiap proses, textjoin
    keterangan dari sheet sumbernya + LP_1 (LAPORAN PROD)."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    processes = body.get("processes") or []
    if not jo:
        return jsonify({"error": "JO wajib diisi"}), 400
    if not processes:
        return jsonify({"error": "Pilih minimal satu proses"}), 400
    try:
        sh = _fstl_spreadsheet()
        # Kumpulin dulu SEMUA nama sheet yang bakal dibutuhin buat SEMUA
        # proses yang dicentang (bisa 10-15 sheet kalau user centang banyak
        # proses), baru ambil sekaligus dalam SATU batch request -- lihat
        # penjelasan lengkap di docstring _fstl_batch_prefetch_sheets().
        needed_sheets = {FSTL_LP1_SHEET}
        for name in processes:
            cfg = FSTL_PROCESS_SOURCES.get(name.strip().upper())
            if cfg is not None:
                needed_sheets.update(_fstl_matching_sheet_names(sh, cfg["prefixes"], cfg["exact"]))
        _fstl_batch_prefetch_sheets(sh, needed_sheets)
        results = {name: fstl_keterangan_for_process(sh, jo, name) for name in processes}
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify({"results": results})


def _fstl_get_or_create_user_sheet(sh, safe_username):
    """Kalau username XX -> sheet 'XX_Kitir'. Kalau sudah ada, dipakai apa
    adanya (TIDAK menghapus sheet lama). Kalau belum ada, dibuat baru KOSONG
    -- baris 1 sengaja TIDAK ditulisi header apa pun, biar tetap kosong buat
    dipakai user sendiri (mis. baris filter Google Sheets), sama seperti pola
    di sheet contoh (kartu-kartu waste mulai dari baris 2).

    `sh` di sini SELALU handle spreadsheet KITIR (_fstl_kitir_spreadsheet()),
    BUKAN spreadsheet sumber data -- tab '{USER}_Kitir' hidup di spreadsheet
    kitir. Nggak perlu _fstl_invalidate_cache() lagi di sini: cache
    worksheet-titles (_fstl_worksheet_titles_cache) itu punya spreadsheet
    sumber data, jadi bikin tab baru di spreadsheet kitir nggak bikin cache
    itu basi."""
    sheet_name = f"{safe_username}_Kitir"
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=200, cols=6)
    return ws, sheet_name


def _fstl_next_card_start_row(all_values):
    """Tentukan baris awal buat kartu baru, dengan aturan jeda 1 baris kosong
    di antara kartu-kartu (bukan lagi nempel/mepet seperti sebelumnya):

      - Sheet masih kosong sama sekali (belum ada kartu)  -> mulai baris 2
        (baris 1 tetap dibiarkan kosong seperti biasa, ini BUKAN "jeda antar
        kartu" jadi tidak perlu ditambah baris kosong lagi).
      - Baris terakhir yang kepakai LANGSUNG berisi data (0 baris kosong
        di bawahnya)         -> selipkan 1 baris kosong, baru mulai kartu.
      - Sudah ada TEPAT 1 baris kosong di bawah baris terakhir yang kepakai
                              -> lanjut langsung, jeda itu sudah cukup.
      - Sudah ada 2 baris kosong atau lebih                -> lanjut langsung
        juga (jangan nambah baris kosong lagi), jeda yang sudah ada dipakai
        apa adanya.

    `all_values` = hasil ws.get_all_values() (list of list of str)."""
    last_filled = 0  # nomor baris (1-based) terakhir yang punya isi
    for idx, row in enumerate(all_values, start=1):
        if any(str(cell).strip() for cell in row):
            last_filled = idx
    if last_filled == 0:
        return 2  # sheet kosong total, belum pernah ada kartu
    trailing_blank = len(all_values) - last_filled
    if trailing_blank == 0:
        return last_filled + 2  # tidak ada jeda -> selipkan 1 baris kosong
    return last_filled + 1  # sudah ada >=1 baris kosong -> lanjut langsung


@app.route("/api/fstl/save", methods=["POST"])
def fstl_save():
    """Body: {username, jo, produk, processes:[{name,keterangan,actionPlan,status}]}.
    Simpan ke sheet '{USERNAME}_Kitir' (dibuat kalau belum ada, tanpa hapus
    sheet lama) sebagai SATU "kartu" yang di-APPEND ke bagian PALING BAWAH
    sheet (bukan disisipkan di atas lagi) -- kolom A & baris 1 TIDAK pernah
    ditulisi apa pun, data mulai kolom B. Antar kartu WAJIB ada jeda 1 baris
    kosong (lihat _fstl_next_card_start_row): kalau kartu sebelumnya nempel
    tanpa jeda, disisipkan 1 baris kosong; kalau sudah ada jeda 1 baris,
    dipakai apa adanya; kalau jedanya sudah 2 baris atau lebih, tidak
    ditambah lagi:

      Baris judul (hijau #a9d08e) : SPK/JO : {jo} | Produk : {produk} |
                                    Waste Besar Proses : {p1, p2, ...}
      Baris label (biru  #9bc2e6) : Waste Besar Proses | Keterangan |
                                    Action Plan | Status   (label kolom)
      Baris data..N (hijau)       : satu baris per proses yang dicentang --
                                    {nama proses} | {keterangan} |
                                    {action plan} | {status}

    (kalau user centang 6 proses, berarti ada 6 baris hijau data di bawah
    baris label, bukan 6 kartu terpisah)."""
    body = request.get_json(force=True) or {}
    username = str(body.get("username", "")).strip()
    jo = str(body.get("jo", "")).strip()
    produk = str(body.get("produk", "")).strip()
    processes = body.get("processes") or []
    if not username:
        return jsonify({"error": "username wajib diisi"}), 400
    if not jo:
        return jsonify({"error": "JO wajib diisi"}), 400
    if not processes:
        return jsonify({"error": "Pilih minimal satu proses"}), 400

    safe_username = "".join(ch for ch in username if ch.isalnum() or ch in ("-", "_")).upper() or "USER"
    try:
        sh = _fstl_kitir_spreadsheet()
        ws, sheet_name = _fstl_get_or_create_user_sheet(sh, safe_username)

        process_names = [str(p.get("name", "")).strip() for p in processes if str(p.get("name", "")).strip()]
        card_title_row = [
            "", f"SPK/JO : {jo}", f"Produk : {produk}",
            f"Waste Besar Proses : {', '.join(process_names)}", "",
        ]
        label_row = ["", "Waste Besar Proses", "Keterangan", "Action Plan", "Status"]
        data_rows = [
            ["", p.get("name", ""), p.get("keterangan", ""), p.get("actionPlan", ""), (p.get("status") or "Open")]
            for p in processes
        ]
        rows_to_write = [card_title_row, label_row] + data_rows

        # APPEND ke bawah dengan jeda 1 baris kosong antar kartu (lihat
        # _fstl_next_card_start_row): sheet kosong -> mulai baris 2 seperti
        # biasa; kalau kartu terakhir nempel tanpa jeda -> disisipkan 1
        # baris kosong; kalau jeda sudah 1 baris atau lebih -> lanjut apa
        # adanya, tidak ditambah jeda baru.
        start_row = _fstl_next_card_start_row(ws.get_all_values())
        end_row = start_row + len(rows_to_write) - 1
        ws.update(f"A{start_row}:E{end_row}", rows_to_write, value_input_option="USER_ENTERED")

        # Pewarnaan: baris judul & baris label sama-sama BIRU (judul cuma
        # sampai kolom D, E dibiarkan putih; label sampai kolom E). Baris
        # data: cuma kolom nama proses (B) yang HIJAU, kolom
        # Keterangan/Action Plan/Status (C:E) tetap PUTIH.
        n = len(data_rows)
        title_row_num, label_row_num = start_row, start_row + 1
        ws.format(f"B{title_row_num}:D{title_row_num}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_TITLE_LABEL)})
        ws.format(f"B{label_row_num}:E{label_row_num}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_TITLE_LABEL)})
        if n:
            data_start, data_end = start_row + 2, start_row + 1 + n
            ws.format(f"B{data_start}:B{data_end}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_PROCESS)})
            ws.format(f"C{data_start}:E{data_end}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_WHITE)})

        # Cache isi sheet ini (dipakai endpoint /api/fstl/list) jadi basi
        # begitu ada kartu baru ditulis -- buang dari cache biar list
        # berikutnya baca versi terbaru, bukan versi sebelum kartu ini ada.
        with _fstl_cache_lock:
            _fstl_sheet_values_cache.pop(sheet_name, None)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify({"success": True, "sheet": sheet_name})


FSTL_KITIR_SUFFIX_RE = re.compile(r"^(.+)_kitir$", re.IGNORECASE)


def _fstl_is_done_status(status_text):
    return str(status_text or "").strip().lower() in ("selesai", "done", "closed", "close", "ok")


def _fstl_parse_cards(rows, username, sheet_name):
    """Baca isi sheet '{username}_Kitir' (list baris dari get_all_values(),
    index kolom 0-based -- kolom B=1, C=2, D=3, E=4) balik jadi list kartu
    {username, sheet, jo, produk, processes:[...]}.

    Baris judul kartu dikenali dari kolom B yang diawali 'SPK/JO'. Baris
    label (header per-kartu) dikenali dari kolom B persis 'Waste Besar
    Proses' DAN kolom C persis 'Keterangan', lalu dilewati (bukan data).
    Baris lain yang punya isi di kolom B dianggap satu baris proses untuk
    kartu yang lagi aktif.

    Tiap baris proses ikut menyimpan nomor barisnya sendiri di sheet asli
    ("row", 1-based, sama seperti nomor baris di Google Sheets) -- dipakai
    /api/fstl/revisi buat tahu persis sel Keterangan mana yang mau ditulis
    ulang, tanpa perlu nebak lagi posisi kartunya di sheet."""

    def cell(row, idx):
        return row[idx].strip() if len(row) > idx else ""

    cards = []
    current = None
    for row_idx, row in enumerate(rows):
        row_num = row_idx + 1  # baris 1 di get_all_values() == baris 1 di sheet
        b, c, d, e = cell(row, 1), cell(row, 2), cell(row, 3), cell(row, 4)
        if not (b or c or d or e):
            continue
        if b.upper().startswith("SPK/JO"):
            if current is not None:
                cards.append(current)
            jo = b.split(":", 1)[1].strip() if ":" in b else b
            produk = c.split(":", 1)[1].strip() if ":" in c else c
            current = {"username": username, "sheet": sheet_name, "jo": jo, "produk": produk, "processes": []}
            continue
        if b == "Waste Besar Proses" and c == "Keterangan":
            continue  # baris label, dilewati
        if current is not None and b:
            current["processes"].append({
                "row": row_num, "name": b, "keterangan": c, "actionPlan": d, "status": e or "Open",
            })
    if current is not None:
        cards.append(current)
    return cards


@app.route("/api/fstl/list", methods=["GET"])
def fstl_list():
    """Ambil semua kartu waste dari SEMUA sheet '*_Kitir' (semua user)
    sekaligus, buat ditampilkan gabung di satu index. Filter yang tadinya
    berdasarkan Status (Open/Selesai) diganti jadi filter berdasarkan nama
    orang yang mengerjakan (turunan nama sheet '{USERNAME}_Kitir').

    Sengaja baca LANGSUNG dari Google Sheets (bukan lewat cache TTL yang
    dipakai /api/fstl/keterangan) -- endpoint ini cuma dipanggil sesekali
    (pas buka halaman Lampiran Waste), jadi lebih penting selalu dapat data
    paling baru (termasuk kartu yang baru saja disimpan) daripada hemat
    kuota API lewat cache basi.

    `sh` di sini spreadsheet KITIR (_fstl_kitir_spreadsheet()), BUKAN
    spreadsheet sumber data -- tab '*_Kitir' hidup di spreadsheet kitir."""
    try:
        sh = _fstl_kitir_spreadsheet()
        cards = []
        usernames = set()
        for ws_obj in sh.worksheets():
            sheet_name = ws_obj.title
            m = FSTL_KITIR_SUFFIX_RE.match(sheet_name)
            if not m:
                continue
            username = m.group(1)
            usernames.add(username)
            rows = ws_obj.get_all_values()
            cards.extend(_fstl_parse_cards(rows, username, sheet_name))
        for card in cards:
            statuses = [p["status"] for p in card["processes"]]
            card["status"] = "done" if statuses and all(_fstl_is_done_status(s) for s in statuses) else "open"
        return jsonify({"cards": cards, "usernames": sorted(usernames)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/fstl/revisi", methods=["POST"])
def fstl_revisi():
    """Body: {sheet, processes:[{row, keterangan}, ...]}.

    Revisi kartu waste yang SUDAH TERSIMPAN -- tapi CUMA kolom Keterangan
    (kolom C) di baris-baris proses yang dikirim yang ditulis ulang. JO,
    Produk, nama proses (centang), Action Plan, dan Status sengaja TIDAK
    bisa diubah lewat endpoint ini (sudah dikunci juga di frontend) --
    endpoint ini murni buat kasus "keterangannya salah/kurang lengkap,
    tolong dibetulkan", bukan buat ganti kartu jadi JO/proses lain.

    "row" per proses adalah nomor baris asli di sheet '{username}_Kitir'
    (dikirim balik oleh /api/fstl/list, lihat _fstl_parse_cards), jadi
    revisi ini langsung nulis ke sel yang tepat tanpa perlu cari ulang
    posisi kartunya."""
    body = request.get_json(force=True) or {}
    sheet_name = str(body.get("sheet", "")).strip()
    processes = body.get("processes") or []
    if not sheet_name or not FSTL_KITIR_SUFFIX_RE.match(sheet_name):
        return jsonify({"error": "Sheet tidak valid"}), 400
    if not processes:
        return jsonify({"error": "Tidak ada keterangan yang direvisi"}), 400

    updates = []
    for p in processes:
        row_num = p.get("row")
        if not isinstance(row_num, int) or row_num < 2:
            continue  # baris 1 nggak pernah dipakai buat data kartu, abaikan kalau ada yang aneh
        keterangan = str(p.get("keterangan", ""))
        updates.append({"range": f"C{row_num}", "values": [[keterangan]]})
    if not updates:
        return jsonify({"error": "Tidak ada baris valid untuk direvisi"}), 400

    try:
        sh = _fstl_kitir_spreadsheet()
        try:
            ws = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            return jsonify({"error": f"Sheet '{sheet_name}' tidak ditemukan di spreadsheet kitir"}), 404
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        # Cache isi sheet ini jadi basi begitu keterangan direvisi, biar
        # /api/fstl/list berikutnya nunjukin versi terbaru.
        with _fstl_cache_lock:
            _fstl_sheet_values_cache.pop(sheet_name, None)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify({"success": True})


# --------------------------------------------------------------------------
# 7. CHATBOT "TANYA JO" — AI yang memutuskan sendiri data apa yang perlu
#    dicari & di sheet mana, lewat tool-use ke query_group. Lihat
#    chatbot_engine.py untuk detail skema sheet & prompt yang dipakai.
# --------------------------------------------------------------------------
# Histori percakapan per sesi disimpan in-memory (sederhana, per session_id
# dari frontend) supaya user bisa tanya susulan ("kalau di Dry gimana?")
# tanpa perlu sebut ulang nomor JO. Kalau nanti dipakai multi-worker/lebih
# dari 1 proses, ganti ke penyimpanan bersama (Redis dsb).
_chat_sessions = {}
_chat_sessions_lock = threading.Lock()
CHAT_HISTORY_MAX_TURNS = 6  # jumlah giliran tanya-jawab yang disimpan per sesi


@app.route("/api/chatbot/ask", methods=["POST"])
def chatbot_ask():
    body = request.get_json(force=True) or {}
    message = str(body.get("message", "")).strip()
    session_id = str(body.get("session_id", "")).strip() or "default"
    if not message:
        return jsonify({"answer": "Pertanyaannya kosong, coba ketik dulu ya."}), 400

    with _chat_sessions_lock:
        history = _chat_sessions.get(session_id, [])

    try:
        result = chatbot_engine.run_agent(get_sheet, message, history=history)
    except Exception as exc:
        return jsonify({"answer": f"Gagal memproses pertanyaan lewat AI: {exc}"}), 500

    with _chat_sessions_lock:
        _chat_sessions[session_id] = chatbot_engine.trim_history(
            result["messages"], max_user_turns=CHAT_HISTORY_MAX_TURNS
        )

    return jsonify({"answer": result["answer"], "tool_calls": result["tool_calls"]})


@app.route("/api/chatbot/reset", methods=["POST"])
def chatbot_reset():
    """Mulai percakapan baru (buang histori) -- dipanggil kalau user pindah
    topik/JO dan mau chatbot-nya tidak kebawa konteks lama."""
    body = request.get_json(force=True) or {}
    session_id = str(body.get("session_id", "")).strip() or "default"
    with _chat_sessions_lock:
        _chat_sessions.pop(session_id, None)
    return jsonify({"success": True})


# --------------------------------------------------------------------------
# HEALTH CHECK (untuk memastikan servis & koneksi sheet hidup)
# --------------------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "spreadsheet_configured": bool(SPREADSHEET_ID)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
