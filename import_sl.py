"""
import_sl.py — lihat catatan di import_printing_2.py.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "sl"
TARGET_SHEET_NAME = "SL_1"

TARGET_HEADERS = [
    "TANGGAL", "MESIN", "SHIFT", "SPK/JO", "NAMA_PRODUK", "UK", "UP", "NO_ROLL",
    "METER", "KG", "HASIL_ROLL", "RW", "ROLL_BAIK", "ROLL_KURLEB", "METER/ROLL",
    "TOTAL_METER", "HASIL_RW", "KG_BRUTO", "KG_NETTO", "OPERATOR", "JAM", "PERSEN",
    "STATUS", "WASTE", "KETERANGAN_WASTE"
]

HEADER_KEYWORDS = ["TANGGAL", "MESIN", "SHIFT", "SPK/JO"]

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS)
