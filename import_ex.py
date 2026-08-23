"""
import_ex.py — lihat catatan di import_printing_2.py.
Menggantikan versi lama (hardcoded SOURCE_SHEET_ID/SHEETS_TO_IMPORT) —
sekarang baca konfigurasi dari config.json (diisi lewat halaman Input Data,
card "Extrusi").
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "ex"
TARGET_SHEET_NAME = "EX_1"

TARGET_HEADERS = [
    "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO", "NAMA_PRODUK",
    "BAHAN_BAKU_LOT", "BAHAN_BAKU_NO", "BAHAN_BAKU_METER", "BAHAN_BAKU_KG", "BAHAN_BAKU_JAM_NAIK",
    "BAHAN_LAPISAN_PP_RANDOM", "BAHAN_LAPISAN_MASTER_BATCH",
    "HASIL_PRODUKSI_NO", "HASIL_PRODUKSI_METER", "HASIL_PRODUKSI_KG", "HASIL_PRODUKSI_JAM_TURUN",
    "WASTE", "KETERANGAN_WASTE"
]

HEADER_KEYWORDS = ["TANGGAL", "SHIFT/OPERATOR"]

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS)
