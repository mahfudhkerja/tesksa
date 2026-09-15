"""
import_rewind_kecil.py
Import Rewind Kecil -> sheet target "REWIND".

Sumber bisa fleksibel:
- type="gsheet"  : Google Sheets, seperti Printing/RW
- type="excel"   : file Excel di Google Drive, seperti Dry

source_id + sheets dibaca dari config.json.
"""

from import_engine import get_source, run_gsheet_import, run_excel_import


SOURCE_KEY = "rewind_kecil"
TARGET_SHEET_NAME = "REWIND"

TARGET_HEADERS = [
    "TANGGAL",
    "SHIFT",
    "OPERATOR",
    "JAM_KERJA",
    "SPK",
    "JO",
    "ORDER",
    "UK",
    "KITIR_MASUK",
    "JUMLAH_AWAL",
    "JUMLAH_AKHIR",
    "KILO_BRUTO",
    "KILO_NETTO",
    "METER",
    "TOTAL_METER",
    "WASTE",
    "KETERANGAN",
    "SISA_RIWEN_BRUTO",
    "SISA_RIWEN_NETTO",
    "METER_ROLL_KECIL",
]

HEADER_KEYWORDS = ["TANGGAL", "SPK", "JO"]

# Untuk source Rewind Kecil kita lebih ketat daripada source lama:
# minimal 2 keyword harus muncul pada baris header yang sama.
HEADER_MIN_MATCHES = 2

# Jangan terlalu agresif membuang baris produksi Rewind Kecil.
JUNK_KEYWORDS = []


def run():
    cfg, src = get_source(SOURCE_KEY)
    source_type = str(src.get("type", "gsheet")).strip().lower()

    if source_type == "gsheet":
        return run_gsheet_import(
            SOURCE_KEY,
            TARGET_SHEET_NAME,
            TARGET_HEADERS,
            HEADER_KEYWORDS,
            junk_keywords=JUNK_KEYWORDS,
            header_min_matches=HEADER_MIN_MATCHES,
        )

    if source_type == "excel":
        return run_excel_import(
            SOURCE_KEY,
            TARGET_SHEET_NAME,
            TARGET_HEADERS,
            HEADER_KEYWORDS,
        )

    raise ValueError(
        f"Source '{SOURCE_KEY}' memiliki type '{source_type}'. "
        "Yang didukung hanya 'gsheet' atau 'excel'."
    )


if __name__ == "__main__":
    run()
