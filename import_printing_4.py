"""
import_printing_4.py — lihat catatan di import_printing_2.py.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "printing_4"
TARGET_SHEET_NAME = "PRINTING_4"

TARGET_HEADERS = [
    "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO", "NAMA_PRODUK",
    "UKURAN_PRODUK", "JENIS_BAHAN", "MICRON_BAHAN", "UK_BAHAN", "NO_LOT_BAHAN",
    "METER_BAHAN", "KG_BAHAN", "NO_ROL_JADI", "METER_AKHIR_JADI", "KG_JADI",
    "JAM_TURUN_JADI", "WIP_RAK", "WIP_BARIS", "KETERANGAN_JADI"
]

HEADER_KEYWORDS = ["TANGGAL", "SHIFT/OPERATOR", "JO", "SPK"]

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS)
