"""
import_rw.py — lihat catatan di import_printing_2.py.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "rw"
TARGET_SHEET_NAME = "RW_1"

TARGET_HEADERS = [
    "TANGGAL", "MESIN", "OPERATOR", "SHIFT", "JAM_KERJA", "FINISH", "SPK", "JO",
    "NAMA_PRODUK", "HPREW_MC", "HPREW_NO", "HPREW_METER", "HPREW_KG",
    "HPREW_METER_AKHIR", "HPREW_KG_AKHIR", "WASTE", "LETAK_WIP", "LETAK_BARIS",
    "WAKTU_NAIK", "WAKTU_TURUN", "KETERANGAN_WASTE"
]

HEADER_KEYWORDS = ["TANGGAL", "OPERATOR", "SHIFT", "JO", "SPK"]

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS)
