"""
run_all.py

Menjalankan semua script import secara berurutan:
Printing -> Rw -> Sl -> Dry -> Sf -> Ex

Setiap script sekarang baca konfigurasinya sendiri (link spreadsheet +
sheet yang dicentang) dari config.json — bukan lagi dari variabel yang
ditulis manual di dalam file masing-masing. Isi/ubah konfigurasi itu
lewat halaman "Input Data" di index.html (tombol Load + Pilih Sheet).

Cara pakai:
    Taruh file ini di folder yang sama dengan script-script lainnya,
    lalu jalankan:

        python run_all.py

Setiap script dijalankan sebagai proses terpisah (subprocess), jadi kalau
salah satu error, script lain tidak ikut rusak dan Anda akan melihat
ringkasan sukses/gagal di akhir.
"""

import os
import subprocess
import sys
import time
from pathlib import Path

# Pastikan output run_all.py sendiri juga aman terhadap emoji/simbol unicode,
# supaya tidak ikut crash saat menampilkan ulang output dari script lain.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Folder tempat semua file .py berada — otomatis mengikuti lokasi run_all.py
# sendiri, jadi tidak perlu diedit manual kalau folder/PC berubah.
SCRIPTS_DIR = Path(__file__).resolve().parent

# Jeda (detik) antar script, supaya kuota Google Sheets API
# (read requests per minute) sempat reset sebelum script berikutnya jalan.
DELAY_BETWEEN_SCRIPTS = 30

# Kalau sebuah script gagal karena quota exceeded (429), berapa kali
# dicoba ulang, dan berapa lama jeda tiap percobaan ulang.
MAX_RETRIES_ON_QUOTA_ERROR = 3
RETRY_DELAY_SECONDS = 65  # kuota Google reset tiap 60 detik, kasih buffer

# Urutan file yang akan dijalankan, dikelompokkan sesuai kategori.
SCRIPTS_ORDER = [
    # --- Printing ---
    "import_printing_2.py",
    "import_printing_3.py",
    "import_printing_4.py",
    "import_printing_5.py",
    # --- Rw ---
    "import_rw.py",
    # --- Sl ---
    "import_sl.py",
    # --- Dry ---
    "import_dry_1.py",
    "import_dry_2.py",
    "import_dry_3.py",
    "import_dry_4.py",
    "import_dry_5.py",
    # --- Sf ---
    "import_sf.py",
    # --- Ex ---
    "import_ex.py",
]


def main():
    results = []

    for script_name in SCRIPTS_ORDER:
        script_path = SCRIPTS_DIR / script_name

        if not script_path.exists():
            print(f"[SKIP] {script_name} tidak ditemukan di {SCRIPTS_DIR}")
            results.append((script_name, "NOT FOUND"))
            continue

        print(f"\n{'=' * 60}")
        print(f"Menjalankan: {script_name}")
        print(f"{'=' * 60}")

        attempt = 0
        final_status = None

        while attempt <= MAX_RETRIES_ON_QUOTA_ERROR:
            attempt += 1
            try:
                # Paksa Python (dan output-nya) pakai UTF-8, supaya emoji/simbol
                # (⚠️, ✅, 📝, dll) yang di-print oleh script tidak bikin crash
                # dengan UnicodeEncodeError saat outputnya di-capture lewat pipe.
                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"
                env["PYTHONUTF8"] = "1"

                proc = subprocess.run(
                    [sys.executable, str(script_path)],
                    cwd=str(SCRIPTS_DIR),
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    env=env,
                )

                # Tampilkan output asli script ke layar
                if proc.stdout:
                    print(proc.stdout, end="")
                if proc.stderr:
                    print(proc.stderr, end="")

                if proc.returncode == 0:
                    print(f"[OK] {script_name} selesai.")
                    final_status = "OK"
                    break

                # Cek apakah gagalnya karena quota/429 dari Google Sheets API
                is_quota_error = "429" in proc.stderr or "Quota exceeded" in proc.stderr

                if is_quota_error and attempt <= MAX_RETRIES_ON_QUOTA_ERROR:
                    print(
                        f"[QUOTA] {script_name} kena limit Google Sheets API. "
                        f"Menunggu {RETRY_DELAY_SECONDS} detik lalu coba lagi "
                        f"(percobaan {attempt}/{MAX_RETRIES_ON_QUOTA_ERROR})..."
                    )
                    time.sleep(RETRY_DELAY_SECONDS)
                    continue
                else:
                    print(f"[GAGAL] {script_name} error dengan exit code {proc.returncode}")
                    final_status = f"FAILED (exit {proc.returncode})"
                    break

            except Exception as e:
                print(f"[GAGAL] {script_name} error saat dijalankan: {e}")
                final_status = f"FAILED ({e})"
                break

        results.append((script_name, final_status))

        # Jeda sebelum lanjut ke script berikutnya (kecuali ini script terakhir)
        if script_name != SCRIPTS_ORDER[-1]:
            print(f"... jeda {DELAY_BETWEEN_SCRIPTS} detik sebelum script berikutnya ...")
            time.sleep(DELAY_BETWEEN_SCRIPTS)

    # Ringkasan akhir
    print(f"\n{'=' * 60}")
    print("RINGKASAN")
    print(f"{'=' * 60}")
    for name, status in results:
        print(f"{name:<30} {status}")


if __name__ == "__main__":
    main()
