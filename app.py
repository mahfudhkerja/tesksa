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
"""

import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

import gspread
from flask import Flask, jsonify, request
from flask_cors import CORS
from google.oauth2.service_account import Credentials

import import_engine

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


def get_client():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


def get_sheet(sheet_name):
    if not SPREADSHEET_ID:
        raise RuntimeError("SPREADSHEET_ID belum diset (lihat file .env)")
    client = get_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(sheet_name)


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
    raw = ws.get_all_records()
    data = [normalize_row(r, VALIDASI_COLUMN_MAP) for r in raw]
    return jsonify(data)


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
    raw = ws.get_all_records()
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
    "ACC": "acc",
}


@app.route("/api/update-stock", methods=["GET"])
def get_update_stock():
    ws = get_sheet("UpdateStock")
    raw = ws.get_all_records()
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
    raw = ws.get_all_records()
    data = [normalize_row(r, STOCKBAHAN_COLUMN_MAP) for r in raw]
    return jsonify(data)


# --------------------------------------------------------------------------
# 6. INPUT DATA PRODUKSI — Load link / Pilih Sheet / Refresh (Run All)
# --------------------------------------------------------------------------

@app.route("/api/produksi/sources", methods=["GET"])
def produksi_sources():
    """Daftar semua source (Printing 2..5, RW, SL, SF, Dry 1..5) beserta
    status koneksi & sheet yang sudah dicentang — dipakai untuk render kartu."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}))


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
        _, src = import_engine.get_source(source_key)
    except KeyError as e:
        return jsonify({"success": False, "message": str(e)}), 404

    try:
        source_id = import_engine.extract_id_from_link(link)
    except ValueError as e:
        return jsonify({"success": False, "message": str(e)}), 400

    try:
        detected_sheets, file_name = import_engine.detect_sheets(source_id, src["type"])
    except Exception as e:
        return jsonify({"success": False, "message": f"Gagal connect ke spreadsheet: {e}"}), 400

    updated = import_engine.update_source(
        source_key,
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
# HEALTH CHECK (untuk memastikan servis & koneksi sheet hidup)
# --------------------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "spreadsheet_configured": bool(SPREADSHEET_ID)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
