"""
import_printing_2.py
Sekarang hanya berisi definisi TARGET_HEADERS/header khusus card ini.
SOURCE_SHEET_ID dan SHEETS_TO_IMPORT tidak lagi ditulis manual di sini —
diambil otomatis dari config.json (diisi lewat halaman Input Data / tombol
Load + Pilih Sheet), supaya tidak ada lagi salah ketik nama sheet.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "printing_2"
TARGET_SHEET_NAME = "PRINTING_2"

TARGET_HEADERS = [
    "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO", "NAMA_PRODUK",
    "UKURAN_PRODUK", "JENIS_BAHAN", "MICRON_BAHAN", "UK_BAHAN", "NO_LOT_BAHAN",
    "METER_BAHAN", "KG_BAHAN", "NO_ROL_JADI", "METER_AKHIR_JADI", "KG_JADI",
    "JAM_TURUN_JADI", "WIP_RAK", "WIP_BARIS", "KETERANGAN_JADI"
]

HEADER_KEYWORDS = ["TANGGAL", "SHIFT/OPERATOR", "JO", "SPK"]

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS)
