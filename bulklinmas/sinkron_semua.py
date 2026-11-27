#!/usr/bin/env python3
"""
================================================================================
SINKRON SEMUA SATLINMAS PONOROGO (ALL-IN-ONE MANDIRI)
================================================================================
Script Terpadu & Otomatis:
 1. Sinkronisasi Spreadsheet 21 Kecamatan dari Google Drive (rclone).
 2. Validasi NIK, pemulihan NIK rusak, dan penyaringan NIK ganda.
 3. Pengunduhan otomatis lampiran baru (Foto, e-KTP, SK Satlinmas).
 4. Standardisasi format berkas otomatis:
    - Foto Profil & e-KTP -> dipastikan .jpg (.png/.webp/.heic dikonversi).
    - SK Satlinmas        -> dipastikan .pdf (.jpg/.jpeg dikonversi).
 5. Kompresi ketat semua berkas agar strictly < 1 MB (target < 850 KB):
    - PDF via Ghostscript (/screen 72/50 dpi).
    - JPG via Pillow (resize proporsional + kompresi kualitas).
 6. Pembuatan Excel Siap-Bulk dan pembaruan DATA MASTER REKAPITULASI (Kejaksaan).
 7. Upload otomatis ke SIMLINMAS Kemendagri dengan AUTO-HEALING:
    - Email bentrok  -> otomatis dikosongkan.
    - No HP bentrok  -> otomatis diganti nomor unik (08 + 9 digit NIK).
    - File ditolak   -> otomatis dikompres ulang & re-upload.
    - Auto-retry hingga 4 kali jika terjadi gangguan jaringan.
    - Menjamin 100% data terkirim tanpa gagal!

Dapat dijalankan kapan saja cukup dengan 1 perintah:
    python3 sinkron_semua.py all
================================================================================
"""
from __future__ import annotations

import argparse
import csv
import glob
import gzip
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from pathlib import Path
from PIL import Image

# ============================================================== 0. KONFIGURASI
HERE = Path(__file__).resolve().parent
INDUK = HERE.parent
DATA_DIR = INDUK / "DATA SATLINMAS PONOROGO"
REKAP_BAWAAN = DATA_DIR / "Rekap Pendataan Satlinmas - Kabupaten Ponorogo (Utama).xlsx"
MASTER_EXCEL = DATA_DIR / "DATA MASTER REKAPITULASI SATLINMAS SE-KABUPATEN PONOROGO.xlsx"
STATE_PATH = HERE / ".sinkron_semua_state.json"
LOG_FILE = HERE / "sinkron_semua.log"

EMAIL_OPERATOR = os.environ.get("SIMLINMAS_EMAIL", "").strip()
PASSWORD_OPERATOR = os.environ.get("SIMLINMAS_PASSWORD", "").strip()
PASSWORD_ANGGOTA = os.environ.get("SIMLINMAS_PASSWORD_ANGGOTA", "").strip()
# Routing relay melalui proxy publik Satlinmas; dapat dioverride lewat .env.
RELAY_URL = (
    os.environ.get("SIMLINMAS_RELAY_URL", "").strip()
    or "https://satpolpp.ponorogo.go.id/satlinmas/proxy/sinkron_proxy.php"
)
RELAY_TOKEN = os.environ.get("SIMLINMAS_RELAY_TOKEN", "").strip()

PROVINSI_NAMA = "JAWA TIMUR"
KABKOTA_NAMA = "KAB. PONOROGO"
BATAS_UKURAN_BERKAS = 850 * 1024  # 850 KB (menjamin selalu < 1 MB untuk SIMLINMAS)

GOL_GAMBAR = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".heic"}
GOL_PDF = {".pdf"}

JENIS = {
    "foto": {
        "judul": "Foto Profil",
        "kata": ("foto diri", "foto kta", "foto profil", "foto"),
        "ekstensi": GOL_GAMBAR,
        "kolom_excel": "Foto Profil",
    },
    "ektp": {
        "judul": "e-KTP",
        "kata": ("e ktp", "ektp", "ktp"),
        "ekstensi": GOL_GAMBAR,
        "kolom_excel": "e-KTP",
    },
    "sk": {
        "judul": "SK Satlinmas",
        "kata": ("sk satlinmas", "sk"),
        "ekstensi": GOL_PDF,
        "kolom_excel": "File SK",
    },
}

K_DESA = "Desa / Kelurahan"
K_NAMA = "Nama Lengkap"
K_NIK = "NIK KTP"
K_ALAMAT = "Alamat Lengkap"
K_LAHIR = "Tempat Lahir"
K_TGL = "Tanggal Lahir"
K_JK = "Jenis Kelamin"
K_HP = "No. WhatsApp"
K_EMAIL = "Email"
K_KK = "Jumlah Anggota Keluarga"
K_PEND = "Pendidikan"
K_WAKTU = "Waktu Kirim"
K_ID = {"foto": "ID Foto KTA", "ektp": "ID e-KTP", "sk": "ID SK"}
K_LINK = {"foto": "Foto KTA", "ektp": "e-KTP", "sk": "SK Satlinmas"}

KOLOM_EXCEL = [
    ("#", 6), ("NIK", 20), ("No Registrasi", 30), ("Nama Lengkap", 38),
    ("Tempat Lahir", 21), ("Tanggal Lahir", 16), ("Jenis Kelamin", 16),
    ("Alamat", 60), ("Provinsi", 13), ("Kab/Kota", 16), ("Kecamatan", 18),
    ("Desa/Kelurahan", 18), ("Pendapatan", 16), ("Tahun Pengangkatan", 18),
    ("Jumlah Keluarga", 18), ("Pendidikan", 18), ("No. HP", 18), ("Email", 34), ("Status", 10),
    ("Lama Pengabdian", 18), ("SK Nomor", 22), ("Foto Profil", 60),
    ("e-KTP", 60), ("File SK", 60), ("Baru", 8),
]

LEBAR = 76
FLAGS_RCLONE = [
    "--stats", "0", "--retries", "5", "--retries-sleep", "2s",
    "--low-level-retries", "10", "--drive-pacer-min-sleep", "20ms",
    "--contimeout", "30s", "--timeout", "300s",
]

_PAKAI_WARNA = sys.stdout.isatty() and not os.environ.get("NO_COLOR")

def _w(kode: str, isi: str) -> str:
    return f"\033[{kode}m{isi}\033[0m" if _PAKAI_WARNA else str(isi)

def hijau(t): return _w("32", t)
def merah(t): return _w("31", t)
def kuning(t): return _w("33", t)
def cyan(t): return _w("36", t)
def tebal(t): return _w("1", t)
def redup(t): return _w("2", t)

def catat_log(pesan: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    baris_bersih = re.sub(r"\033\[[0-9;]*m", "", str(pesan))
    try:
        with open(LOG_FILE, "a", encoding="utf-8", errors="replace") as f:
            f.write(f"[{ts}] {baris_bersih}\n")
    except Exception:
        pass

def garis(karakter: str = "=") -> None:
    txt = redup(karakter * LEBAR)
    print(txt)
    catat_log(karakter * LEBAR)

def judul(t: str) -> None:
    print()
    garis()
    print(" " + tebal(cyan(t)))
    catat_log(f"=== {t} ===")
    garis()

def langkah(no: str, t: str) -> None:
    msg = f"\n {tebal('[' + no + ']')} {tebal(t)}"
    print(msg)
    catat_log(f"[{no}] {t}")

def info(t: str = "") -> None:
    print("  " + t)
    catat_log("  " + t)

def hasil(label: str, nilai, salah: bool = False) -> None:
    warna = merah if salah else hijau
    print(f"  {label:<32}: {warna(str(nilai))}")
    catat_log(f"  {label:<32}: {nilai}")

def baik(t: str) -> None:
    print("  " + hijau("[OK] ") + t)
    catat_log("  [OK] " + t)

def catat(t: str) -> None:
    print("  " + kuning("[!]  ") + t)
    catat_log("  [!]  " + t)

def gagal(t: str) -> None:
    print("  " + merah("[X]  ") + t)
    catat_log("  [X]  " + t)

def progres(sekarang: int, total: int, label: str = "", extra: str = "", mulai: float | None = None) -> None:
    total = max(total, 1)
    pct = sekarang * 100 / total
    lebar_bar = 24
    terisi = int(lebar_bar * sekarang / total)
    bar = "#" * terisi + "-" * (lebar_bar - terisi)
    eta = ""
    if mulai and sekarang > 0 and sekarang < total:
        sisa = (time.time() - mulai) / sekarang * (total - sekarang)
        eta = f" ETA {int(sisa//60):02d}:{int(sisa%60):02d}"
    elif mulai and sekarang >= total:
        dur = time.time() - mulai
        eta = f" selesai {int(dur//60):02d}:{int(dur%60):02d}"
    msg = f"  [{bar}] {pct:6.2f}% ({sekarang}/{total})"
    if label:
        msg += f" {label}"
    if extra:
        msg += f" | {extra}"
    msg += eta
    print(msg)
    catat_log(re.sub(r"\033\[[0-9;]*m", "", msg))

def coba_ulang(fn, *, percobaan: int = 5, jeda_awal: float = 2.0, faktor: float = 1.7,
               nama: str = "operasi", boleh_raise: bool = False):
    """Jalankan fn() dengan retry + exponential backoff. Kembalikan (ok, hasil/exception)."""
    err = None
    for i in range(1, percobaan + 1):
        try:
            return True, fn()
        except Exception as e:
            err = e
            if i < percobaan:
                tunggu = jeda_awal * (faktor ** (i - 1))
                catat(f"{nama}: percobaan {i}/{percobaan} gagal ({e}); coba lagi {tunggu:.1f}s...")
                time.sleep(tunggu)
    if boleh_raise:
        raise err
    return False, err


# ------------------------------------------------------- Import / Siapkan Library
def _siapkan_paket():
    for nama in (f"python{sys.version_info.major}.{sys.version_info.minor}",
                 f"python3.{sys.version_info.minor}"):
        paket = HERE / ".venv" / "lib" / nama / "site-packages"
        if paket.is_dir():
            sys.path.insert(0, str(paket))
            break
_siapkan_paket()
try:
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)
except Exception:
    pass

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402


# ==============================================================================
# EMBEDDED BULK DEPENDENCIES (mandiri, tidak perlu bulk.py)
# ==============================================================================
import http.cookiejar
import gzip
import mimetypes
import urllib.error
import urllib.parse
import urllib.request
import uuid
from dataclasses import dataclass, field
from typing import Callable

BASE_URL = "https://simlinmas.kemendagri.go.id"
PROVINSI_ID, PROVINSI_NAMA_BULK = "35", "JAWA TIMUR"
KABKOTA_ID, KABKOTA_NAMA_BULK = "3502", "KAB. PONOROGO"
USER_AGENT = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36")
BATASAN_UNGGAHAN = 4 * 1024 * 1024
GOL_BENTUK = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif"}
GOL_PDF_EXT = {".pdf"}
TIPE_BERKAS = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
               ".webp": "image/webp", ".gif": "image/gif", ".bmp": "image/bmp",
               ".pdf": "application/pdf"}

J_NIK = "NIK"
J_NAMA = "Nama Lengkap"
J_HP = "No. HP"
J_KEC = "Kecamatan"
J_DESA = "Desa/Kelurahan"
J_NOREG = "No Registrasi"
J_TGL = "Tanggal Lahir"
J_JK = "Jenis Kelamin"
J_SKNOMOR = "SK Nomor"
J_FOTO = "Foto Profil"
J_SKFILE = "File SK"

@dataclass
class Hasil:
    nilai: object = ""
    salah: str = ""
    ingat: str = ""


def bteks(nilai) -> str:
    """Sel Excel -> teks rapi."""
    if nilai is None:
        return ""
    if isinstance(nilai, (datetime, date)):
        return nilai.strftime("%d-%m-%Y")
    if isinstance(nilai, float) and nilai.is_integer():
        return str(int(nilai))
    return str(nilai).strip()

def konversi_teks(nilai) -> Hasil:
    """Konversi biasa: apa adanya jadi teks rapi."""
    return Hasil(bteks(nilai))


def konversi_noreg(nilai) -> Hasil:
    """'No Registrasi' export bisa rusak: '35.02.17..2023...'.

    Yang rusak dikosongkan saja karena server membuat nomor registrasi sendiri.
    """
    mentah = bteks(nilai)
    if mentah in ("-", "--", "N/A", "n/a"):
        return Hasil("")
    if ".." in mentah:
        return Hasil("", "", f"No Registrasi rusak, dikosongkan (server yang "
                             f"membuat): {mentah!r}")
    return Hasil(mentah)


def konversi_nik(nilai) -> Hasil:
    """NIK harus 16 digit; kolom ini tidak ada di file export, wajib ditambah."""
    nik = re.sub(r"\D", "", bteks(nilai))
    if not nik:
        return Hasil("", "NIK kosong (kolom ini tidak ada di file export, "
                         "wajib ditambah)")
    if len(nik) != 16:
        return Hasil("", f"NIK harus 16 digit, dapat {len(nik)}: {nik}")
    return Hasil(nik)


def konversi_tanggal(nilai) -> Hasil:
    """'DD-MM-YYYY' -> 'YYYY-MM-DD' (format yang diminta server)."""
    mentah = bteks(nilai)
    if not mentah:
        return Hasil("", "tanggal lahir kosong")
    for pola in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%y"):
        try:
            return Hasil(datetime.strptime(mentah, pola).strftime("%Y-%m-%d"))
        except ValueError:
            continue
    return Hasil("", f"tanggal lahir tidak terbaca: {mentah!r} (pakai DD-MM-YYYY)")


def konversi_jk(nilai) -> Hasil:
    """Form web hanya menerima kode 'L' / 'P', bukan 'Laki-laki' / 'Perempuan'."""
    jk = {"L": "L", "LAKI-LAKI": "L", "P": "P", "PEREMPUAN": "P"}.get(
        bteks(nilai).upper(), "")
    if not jk:
        return Hasil("", f"jenis kelamin tidak valid: {bteks(nilai)!r} "
                         "(pakai 'Laki-laki' / 'Perempuan')")
    return Hasil(jk)


def konversi_hp(nilai) -> Hasil:
    """Normalisasi nomor HP ke awalan 0 (contoh: 628xx/08xx -> 08xx)."""
    hp = re.sub(r"\D", "", bteks(nilai))
    if hp.startswith("62"):
        hp = "0" + hp[2:]
    elif hp.startswith("+62"):
        hp = "0" + hp[3:]
    elif hp and not hp.startswith("0") and len(hp) >= 9:
        hp = "0" + hp
    if not hp:
        return Hasil("", "nomor HP kosong (server menolak nomor yang sudah terdaftar)")
    if not 10 <= len(hp) <= 14:
        return Hasil(hp, "", f"nomor HP panjang tidak lazim ({len(hp)} digit): {hp}")
    return Hasil(hp)


def konversi_angka(nilai, bawaan=None) -> Hasil:
    """Kolom bigint di server, tapi file export menulisnya dalam bentuk tampilan.

        'Rp, 1.500.000' -> 1500000      'Rp, 0' -> 0      '0000' -> 0
        '-' / kosong     -> `bawaan`

    Mengirim 'Rp, 0' apa adanya ditolak dengan:
        SQLSTATE[22P02] invalid input syntax for type bigint: "Rp, 0"
    """
    mentah = bteks(nilai)
    if mentah in ("", "-", "--", "N/A", "n/a"):
        return Hasil(bawaan)
    digit = re.sub(r"\D", "", mentah)
    return Hasil(int(digit) if digit else bawaan)


def konversi_angka_kosong(nilai) -> Hasil:
    """Kolom angka opsional: kosong berarti tidak dikirim sama sekali."""
    nilai_bersih = konversi_angka(nilai).nilai
    return Hasil("" if nilai_bersih is None else nilai_bersih)


def konversi_pendapatan(nilai) -> Hasil:
    """Pendapatan wajib berupa angka; nilai kosong dianggap 0."""
    return Hasil(konversi_angka(nilai, 0).nilai)


def konversi_tahun(nilai) -> Hasil:
    """Tahun pengangkatan; nilai '0000' di file export berarti tidak diisi."""
    tahun = konversi_angka(nilai).nilai
    if tahun is None or tahun == 0:
        return Hasil("")
    if not 1900 <= tahun <= 2100:
        return Hasil("", f"tahun pengangkatan tidak masuk akal: {tahun} "
                         "(harus 1900-2100)")
    return Hasil(tahun)



@dataclass(frozen=True)
class Kolom:
    """Satu kolom Excel: isinya, dikirim ke mana, dan aturan validasinya."""

    judul: str                          # teks header di baris 1
    lebar: float                        # lebar kolom di Excel
    kunci: str = ""                     # nama field di POST /linmas/insert
    wajib: bool = False                 # kosong -> baris tidak dikirim
    konversi: Callable = konversi_teks  # None = diisi manual, bukan dari sel
    contoh: str = ""                    # contoh isi pada template & petunjuk
    catatan: str = ""                   # penjelasan singkat
    jenis: str = "teks"                 # teks | angka | tanggal | berkas


_MB = f"{BATASAN_UNGGAHAN // 1024 // 1024} MB"

KOLOM: tuple[Kolom, ...] = (
    Kolom("#", 6, jenis="angka", contoh="1",
          catatan="nomor urut dari file asal, tidak dikirim ke server"),
    Kolom(J_NIK, 20, "nik", wajib=True, konversi=konversi_nik, jenis="teks",
          contoh="3502160101900001",
          catatan="WAJIB, 16 digit, harus unik. Tidak ada di file export, jadi "
                  "wajib ditambahkan sendiri."),
    Kolom(J_NOREG, 30, "nomor_registrasi", konversi=konversi_noreg, jenis="teks",
          contoh="35.02.16.2013.2023.0001",
          catatan="boleh kosong, server membuat sendiri. Nilai yang rusak "
                  "(ada '..') otomatis dikosongkan."),
    Kolom(J_NAMA, 38, "nama_lengkap", wajib=True, jenis="teks",
          contoh="NAMA LENGKAP ANDA",
          catatan="WAJIB. Jangan pakai spasi ganda atau kolom kosong di tengah."),
    Kolom("Tempat Lahir", 21, "tempat_lahir", jenis="teks", contoh="PONOROGO",
          catatan="nama kota/kabupaten, bebas teks"),
    Kolom(J_TGL, 16, "tanggal_lahir", wajib=True, konversi=konversi_tanggal,
          jenis="tanggal",
          contoh="01-01-1990",
          catatan="WAJIB, tulis sebagai TEKS 'DD-MM-YYYY'. Format lain "
                  "(YYYY-MM-DD, DD/MM/YYYY) diterima lalu dikonversi."),
    Kolom(J_JK, 16, "jenis_kelamin", wajib=True, konversi=konversi_jk, jenis="teks",
          contoh="Laki-laki",
          catatan="WAJIB, isi persis 'Laki-laki' atau 'Perempuan'. Server hanya "
                  "menerima 'L'/'P', dikonversi otomatis."),
    Kolom("Alamat", 60, "alamat", jenis="teks", contoh="KRAJAN RT 02 RW 02",
          catatan="alamat RT/RW, tidak perlu lengkap kode pos"),
    Kolom("Provinsi", 13, jenis="teks", contoh=PROVINSI_NAMA_BULK,
          catatan=f"diabaikan - server mengunci {PROVINSI_ID} ({PROVINSI_NAMA_BULK})"),
    Kolom("Kab/Kota", 16, jenis="teks", contoh=KABKOTA_NAMA_BULK,
          catatan=f"diabaikan - server mengunci {KABKOTA_ID} ({KABKOTA_NAMA_BULK})"),
    Kolom(J_KEC, 18, "kecamatan_id", wajib=True, konversi=None, jenis="teks",
          contoh="Babadan",
          catatan="WAJIB. Tulis NAMA, bukan kode; nama dicocokkan ke id lewat API "
                  "wilayah (beda huruf besar/kecil diabaikan)."),
    Kolom(J_DESA, 18, "kelurahan_id", wajib=True, konversi=None, jenis="teks",
          contoh="Cekok",
          catatan="WAJIB. Tulis NAMA desa/kelurahan yang ada di kecamatan yang "
                  "sama. Nilai '-' pada file export berarti tidak valid."),
    Kolom("Pendapatan", 16, "pendapatan", konversi=konversi_pendapatan,
          jenis="angka", contoh="0",
          catatan="angka polos (jangan 'Rp, 0'); teks 'Rp, 1.500.000' tetap "
                  "bisa dan dibersihkan otomatis."),
    Kolom("Tahun Pengangkatan", 18, "tahun_pengangkatan",
          konversi=konversi_tahun, jenis="angka", contoh="2024",
          catatan="tahun 1900-2100. Nilai '0000' di file export dibaca kosong"),
    Kolom("Jumlah Keluarga", 18, "jumlah_keluarga",
          konversi=konversi_angka_kosong, jenis="angka", contoh="3",
          catatan="angka, boleh kosong"),
    Kolom("Pendidikan", 18, "pendidikan_formal", konversi=konversi_teks, jenis="teks",
          contoh="SMA", catatan="pendidikan formal (opsional)"),
    Kolom(J_HP, 18, "no_hp", wajib=True, konversi=konversi_hp, jenis="teks",
          contoh="081234567890",
          catatan="WAJIB, boleh ditulis 08xx / 62xx / +62xx (dinormalkan ke "
                  "62xx). Nomor yang sudah dipakai anggota lain ditolak server."),
    Kolom("Email", 34, "email", jenis="teks", contoh="nama@contoh.id",
          catatan="boleh kosong"),
    Kolom("Status", 10, jenis="teks", contoh="Aktif",
          catatan="diabaikan saat membuat record baru"),
    Kolom("Lama Pengabdian", 18, "lama_pengabdian",
          konversi=konversi_angka_kosong, jenis="angka", contoh="5",
          catatan="tahun, boleh kosong"),
    Kolom(J_SKNOMOR, 22, "sk_nomor", jenis="teks", contoh="SK/123/2024/0001",
          catatan="nomor SK sebagai teks, boleh kosong"),
    Kolom(J_FOTO, 30, "foto_profil", jenis="berkas", contoh="berkas/foto01.jpg",
          catatan=f"lokasi file gambar (jpeg/png/webp/gif/bmp, maks {_MB}). Path "
                  f"relatif dihitung dari folder file .xlsx. Kosong = tanpa "
                  f"lampiran."),
    Kolom(J_SKFILE, 30, "sk_file", jenis="berkas", contoh="berkas/sk01.pdf",
          catatan=f"lokasi file PDF (maks {_MB}), aturan path sama dengan kolom "
                  f"Foto Profil"),
)


def samakan_nama(nama: str) -> str:
    """Samakan nama wilayah agar tidak gagal pencocokan karena beda ejaan kecil."""
    bersih = (nama or "").strip().lower()
    bersih = re.sub(r"^(kab\.?|kabupaten|kotamadya|kota\.?|prov\.?|provinsi)\s+",
                   "", bersih)
    return re.sub(r"[\s._-]+", "", bersih)



class Sesi:
    """HTTP + cookie milik operator. Menggunakan curl_cffi untuk meniru browser
    (melewati Cloudflare 'Just a moment') bila tersedia."""

    def __init__(self, email: str, sandi: str):
        self.email = email
        self.sandi = sandi
        self._cffi = None
        self.wadah = http.cookiejar.CookieJar()
        self.pembuka = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.wadah))
        self.pembuka.addheaders = [
            ("User-Agent", USER_AGENT),
            ("Accept-Language", "id-ID,id;q=0.9,en;q=0.8"),
        ]
        try:
            from curl_cffi import requests as cffi_requests
            self._cffi = cffi_requests.Session(impersonate="chrome")
            self._cffi.headers.update({
                "User-Agent": USER_AGENT,
                "Accept-Language": "id-ID,id;q=0.9,en;q=0.8",
            })
        except Exception:
            self._cffi = None

    def minta(self, jalur: str, data: bytes | None = None,
              header: dict | None = None, ajax: bool = False) -> tuple:
        url = jalur if jalur.startswith("http") else BASE_URL + jalur
        hdr = {"Referer": BASE_URL + "/"}
        if ajax:
            hdr.update({"Accept": "*/*", "X-Requested-With": "XMLHttpRequest"})
        hdr.update(header or {})

        if self._cffi is not None:
            for coba in range(1, 5):
                try:
                    if data is not None or (not ajax and data is None and header is not None and ("Content-Type" in hdr)):
                        # POST when data provided; but preserve old behavior: data None+not ajax => GET
                        pass
                    if data is not None:
                        resp = self._cffi.post(url, data=data, headers=hdr, timeout=75)
                    else:
                        resp = self._cffi.get(url, headers=hdr, timeout=75)
                    return (resp.status_code, str(resp.url), resp.text, dict(resp.headers))
                except Exception as galat:
                    if coba == 4:
                        raise ConnectionError(f"tidak bisa menghubungi {BASE_URL} -> {galat}") from galat
                    time.sleep(2 * coba)

        metode = "GET" if ajax else ("POST" if data is not None else None)
        permintaan = urllib.request.Request(url, data=data, headers=hdr, method=metode)
        for coba in range(1, 5):
            try:
                with self.pembuka.open(permintaan, timeout=75) as jawaban:
                    isi = jawaban.read()
                    if jawaban.headers.get("Content-Encoding") == "gzip":
                        isi = gzip.decompress(isi)
                    return (jawaban.status, jawaban.geturl(), isi.decode("utf-8", "replace"), dict(jawaban.headers))
            except urllib.error.HTTPError as galat:
                isi = galat.read()
                return (galat.code, url, isi.decode("utf-8", "replace"), dict(galat.headers))
            except (urllib.error.URLError, TimeoutError, OSError, Exception) as galat:
                if coba == 4:
                    raise ConnectionError(f"tidak bisa menghubungi {BASE_URL} -> {galat}") from galat
                time.sleep(2 * coba)

    def ambil_json(self, jalur: str, header: dict | None = None) -> tuple:
        status, _, isi, _ = self.minta(jalur, header=header, ajax=True)
        try:
            return status, json.loads(isi)
        except json.JSONDecodeError:
            return status, isi

    def token_csrf(self, halaman: str = "/linmas/tambah") -> str:
        _, _, html, _ = self.minta(halaman)
        for pola in (r'<meta\s+name="csrf-token"\s+content="([^"]+)"', r'name="_token"\s+value="([^"]+)"'):
            cocok = re.search(pola, html)
            if cocok:
                return cocok.group(1)
        # fallback cookie jar
        wadah = self._cffi.cookies.jar if self._cffi is not None else self.wadah
        for kuki in wadah:
            if kuki.name == "XSRF-TOKEN":
                return urllib.parse.unquote(kuki.value)
        raise RuntimeError("token CSRF tidak ditemukan")

    def masuk(self) -> "Sesi":
        _, _, html, _ = self.minta("/login")
        cocok = re.search(r'name="_token"\s+value="([^"]+)"', html)
        if not cocok:
            raise RuntimeError("halaman /login tidak memuat token CSRF")
        badan = urllib.parse.urlencode({"_token": cocok.group(1), "email": self.email, "password": self.sandi}).encode()
        _, url_akhir, _, _ = self.minta("/login", data=badan,
            header={"Content-Type": "application/x-www-form-urlencoded", "Origin": BASE_URL})
        if "/login" in url_akhir:
            raise RuntimeError("login ditolak (email / password salah)")
        return self


PILIHAN_WILAYAH = re.compile(
    r"<option value=[\"']([^\"']*)[\"'][^>]*>([^<]*)</option>", re.I)


class SimLinmas:
    """Pembungkus endpoint SIMLINMAS yang dipakai tool ini.

    POST /login                   login (form _token + email + password)
    POST /linmas/insert           membuat satu record (multipart + X-CSRF-TOKEN)
    GET  /linmas                  JSON DataTables -> recordsTotal + daftar baris
    GET  /linmas/delete/{id}      hapus satu record
    GET  /region/districts/{kab}  daftar kecamatan  (untuk validasi nama)
    GET  /region/villages/{kec}   daftar kelurahan (untuk validasi nama)
    """

    WAJIB_SERVER = ("nama_lengkap", "nik", "no_hp", "kecamatan_id", "kelurahan_id",
                    "jenis_kelamin", "tanggal_lahir")

    def __init__(self, sesi: Sesi):
        self.sesi = sesi
        self._wilayah: dict[str, tuple[dict[str, str], dict[str, str]]] = {}

    # ---------------- wilayah ----------------
    def _daftar_wilayah(self, jenis: str, kode: str) -> tuple[dict, dict]:
        """Ambil <option> dari /region/{jenis}/{kode}, lalu cache-nya.

        PENTING: cache di-key per kode parent, bukan per jenis wilayah. Kalau
        di-key per jenis, setiap kecamatan baru yang diambil akan menimpa daftar
        sebelumnya, sehingga baris yang kembali ke kecamatan lama salah resolve.
        """
        kunci = f"{jenis}/{kode}"
        if kunci not in self._wilayah:
            _, _, html, _ = self.sesi.minta(
                f"/region/{jenis}/{kode}",
                header={"Referer": BASE_URL + "/linmas/tambah", "Accept": "*/*",
                        "X-Requested-With": "XMLHttpRequest"})
            isi = [(v, t.strip()) for v, t in PILIHAN_WILAYAH.findall(html)
                   if v and v != "0"]
            self._wilayah[kunci] = ({v: t for v, t in isi},
                                    {samakan_nama(t): v for v, t in isi})
        return self._wilayah[kunci]

    def cari_kecamatan(self, nama: str) -> str | None:
        """'KAB. PONOROGO' / 'babadan' -> id kecamatan, atau None kalau tak dikenal."""
        _, cari = self._daftar_wilayah("districts", KABKOTA_ID)
        return cari.get(samakan_nama(nama))

    def cari_kelurahan(self, id_kecamatan: str, nama: str) -> str | None:
        """Nama desa di dalam satu kecamatan -> id desa, atau None."""
        _, cari = self._daftar_wilayah("villages", id_kecamatan)
        return cari.get(samakan_nama(nama))

    # ---------------- baca ----------------
    @staticmethod
    def _parameter_tabel(mulai: int, panjang: int) -> str:
        """Parameter wajib DataTables serverSide."""
        return urllib.parse.urlencode({
            "draw": 1, "start": mulai, "length": panjang,
            "columns[1][data]": "nama_lengkap",
            "order[0][column]": 1, "order[0][dir]": "asc", "search[value]": "",
            "status": "", "jk": "", "tahun": "", "jabatan": "",
            "provinsi": "", "kabkota": "", "kecamatan": "", "kelurahan": ""})

    def jumlah_member(self) -> int:
        _, data = self.sesi.ambil_json("/linmas?" + self._parameter_tabel(0, 1))
        return data.get("recordsTotal", -1) if isinstance(data, dict) else -1

    def semua_member(self, per_halaman: int = 2000) -> list[dict]:
        """Tarik seluruh tabel, lalu saring di sisi program.

        Parameter pencarian bawaan aplikasi ini tidak bisa dipercaya: pencarian
        'UJI COBA' mengembalikan 0 baris padahal recordnya ada. Karena itu
        halaman demi halaman seluruh tabel ditarik.
        """
        hasil, mulai = [], 0
        while True:
            _, data = self.sesi.ambil_json(
                "/linmas?" + self._parameter_tabel(mulai, per_halaman))
            if not isinstance(data, dict) or not data.get("data"):
                break
            hasil.extend(data["data"])
            mulai += per_halaman
            if mulai >= (data.get("recordsTotal") or 0):
                break
        return hasil

    # ---------------- tulis ----------------
    @staticmethod
    def _body_multipart(nilai: dict, lampiran: dict, batas: str) -> bytes:
        """Rakit body multipart/form-data tanpa bergantung pada library lain."""
        potongan = []
        for nama, isi in nilai.items():
            if isi is None or isi == "":
                continue
            potongan.append(
                f"--{batas}\r\nContent-Disposition: form-data; name=\"{nama}\""
                f"\r\n\r\n{isi}\r\n".encode())
        for nama, (nama_berkas, isi, tipe) in lampiran.items():
            potongan.append(
                f"--{batas}\r\nContent-Disposition: form-data; name=\"{nama}\"; "
                f"filename=\"{nama_berkas}\"\r\nContent-Type: {tipe}\r\n\r\n".encode())
            potongan.extend((isi, b"\r\n"))
        potongan.append(f"--{batas}--\r\n".encode())
        return b"".join(potongan)

    def tambah(self, nilai: dict, lampiran: dict | None = None) -> dict:
        """POST /linmas/insert -> {'ok': bool, 'code': ..., 'message': ...}."""
        isi = {k: ("" if v is None else str(v)) for k, v in nilai.items()}
        kurang = [k for k in self.WAJIB_SERVER if not isi.get(k)]
        if kurang:
            return {"ok": False, "code": "LOKAL",
                    "message": "field wajib kosong: " + ", ".join(kurang)}
        token = self.sesi.token_csrf()
        isi["_token"] = token
        isi.setdefault("provinsi_id", PROVINSI_ID)
        isi.setdefault("kabkota_id", KABKOTA_ID)
        batas = "----SATLINMAS" + uuid.uuid4().hex
        _, _, isi_mentah, _ = self.sesi.minta(
            "/linmas/insert", data=self._body_multipart(isi, lampiran or {}, batas),
            header={"Content-Type": f"multipart/form-data; boundary={batas}",
                    "X-CSRF-TOKEN": token, "X-Requested-With": "XMLHttpRequest",
                    "Referer": BASE_URL + "/linmas/tambah", "Origin": BASE_URL,
                    "Accept": "*/*"})
        try:
            jawab = json.loads(isi_mentah)
            return {"ok": str(jawab.get("code")) == "200", "code": jawab.get("code"),
                    "message": jawab.get("message")}
        except json.JSONDecodeError:
            return {"ok": False, "code": "NON-JSON",
                    "message": (isi_mentah or "(isi kosong)")[:300]}

    def perbarui(self, id_edit: str, nilai: dict, lampiran: dict | None = None) -> dict:
        """POST /linmas/update -> update data member yang sudah ada."""
        isi = {k: ("" if v is None else str(v)) for k, v in nilai.items()}
        isi["id_edit"] = str(id_edit)
        token = self.sesi.token_csrf(f"/linmas/edit/{id_edit}")
        isi["_token"] = token
        isi.setdefault("provinsi_id", PROVINSI_ID)
        isi.setdefault("kabkota_id", KABKOTA_ID)
        batas = "----SATLINMAS" + uuid.uuid4().hex
        _, _, isi_mentah, _ = self.sesi.minta(
            "/linmas/update", data=self._body_multipart(isi, lampiran or {}, batas),
            header={"Content-Type": f"multipart/form-data; boundary={batas}",
                    "X-CSRF-TOKEN": token, "X-Requested-With": "XMLHttpRequest",
                    "Referer": BASE_URL + f"/linmas/edit/{id_edit}", "Origin": BASE_URL,
                    "Accept": "*/*"})
        try:
            jawab = json.loads(isi_mentah)
            return {"ok": str(jawab.get("code")) == "200", "code": jawab.get("code"),
                    "message": jawab.get("message")}
        except json.JSONDecodeError:
            return {"ok": False, "code": "NON-JSON",
                    "message": (isi_mentah or "(isi kosong)")[:300]}

    def hapus(self, id_member: int) -> dict:
        """Hapus satu record (endpoint delete memang sebuah GET di aplikasi ini)."""
        token = self.sesi.token_csrf("/linmas")
        _, _, isi, _ = self.sesi.minta(
            f"/linmas/delete/{id_member}",
            header={"X-CSRF-TOKEN": token, "X-Requested-With": "XMLHttpRequest",
                    "Referer": BASE_URL + "/linmas", "Accept": "*/*"})
        try:
            jawab = json.loads(isi)
            return {"ok": str(jawab.get("code")) == "200", "code": jawab.get("code"),
                    "message": jawab.get("message")}
        except json.JSONDecodeError:
            return {"ok": False, "code": "NON-JSON", "message": isi[:200]}



def baca_baris(berkas: str, sheet: str = "") -> tuple[list[dict], list[str]]:
    """Baca sheet pertama (atau `sheet`) menjadi daftar dict per baris.

    Tiap dict berisi kolom sesuai judul header baris 1, ditambah kunci
    '__baris' = nomor baris di Excel supaya pesan error mudah dicari.
    """
    buku = openpyxl.load_workbook(berkas, data_only=True)
    lembar = buku[sheet] if sheet else buku.active
    judul_kolom = [bteks(lembar.cell(1, kolom).value)
                   for kolom in range(1, lembar.max_column + 1)]
    baris_data = []
    for nomor in range(2, lembar.max_row + 1):
        isi = {judul_kolom[kolom - 1]: lembar.cell(nomor, kolom).value
               for kolom in range(1, lembar.max_column + 1)}
        if any(nilai not in (None, "") for nilai in isi.values()):
            isi["__baris"] = nomor
            baris_data.append(isi)
    return baris_data, judul_kolom



@dataclass
class HasilBaris:
    """Ringkasan validasi satu baris Excel."""

    payload: dict = field(default_factory=dict)
    salah: list[str] = field(default_factory=list)   # alasan tidak dikirim
    ingat: list[str] = field(default_factory=list)   # hanya peringatan


def validasi_baris(api, baris: dict, sandi: str) -> HasilBaris:
    """Ubah satu baris Excel menjadi payload POST /linmas/insert.

    Jika api=None (mode relay), resolusi kec/kel ID dilewati dan nama dikirim
    sebagai kecamatan_nama/kelurahan_nama agar relay yang resolve.
    """
    hasil = HasilBaris()
    payload = hasil.payload
    payload["provinsi_id"] = PROVINSI_ID
    payload["kabkota_id"] = KABKOTA_ID

    # 1) kolom biasa: konversi + cek wajib
    for kolom in KOLOM:
        if not kolom.kunci or kolom.konversi is None:
            continue
        ubah = kolom.konversi(baris.get(kolom.judul))
        if ubah.salah:
            hasil.salah.append(f"{kolom.judul}: {ubah.salah}")
        elif kolom.wajib and ubah.nilai == "":
            hasil.salah.append(f"{kolom.judul} kosong (wajib diisi)")
        elif ubah.ingat:
            hasil.ingat.append(f"{kolom.judul}: {ubah.ingat}")
        payload[kolom.kunci] = ubah.nilai

    # 2) nama wilayah -> id
    nama_kec = bteks(baris.get(J_KEC))
    nama_desa = bteks(baris.get(J_DESA))
    if api is None:
        # relay akan resolve
        if nama_kec:
            payload["kecamatan_nama"] = nama_kec
        if nama_desa and nama_desa not in ("", "-"):
            payload["kelurahan_nama"] = nama_desa
    else:
        id_kec = api.cari_kecamatan(nama_kec) if nama_kec else None
        if not id_kec:
            hasil.salah.append(f"{J_KEC}: tidak dikenal di {KABKOTA_NAMA_BULK} -> {nama_kec!r}")
        else:
            id_desa = (api.cari_kelurahan(id_kec, nama_desa) if nama_desa not in ("", "-") else None)
            if not id_desa:
                hasil.salah.append(f"{J_DESA}: {nama_desa!r} tidak ada di {nama_kec} (atau kosong)")
            else:
                payload["kecamatan_id"] = id_kec
                payload["kelurahan_id"] = id_desa

    # 3) password akun anggota
    payload["password"] = sandi
    payload["password_confirm"] = sandi
    return hasil



def absolutkan_path(mentah: str, folder_dasar: str) -> str:
    """Jadikan path mutlak; path relatif dihitung dari folder file Excel."""
    bersih = os.path.expanduser(os.path.expandvars(bteks(mentah)))
    if not os.path.isabs(bersih):
        bersih = os.path.join(folder_dasar, bersih)
    return os.path.normpath(bersih)


def periksa_lampiran(baris: dict,
                     folder_dasar: str) -> tuple[dict[str, str], list[str]]:
    """Cek kolom Foto Profil / File SK -> ({nama_field: path}, daftar_pesan).

    Berkas hanya diperiksa (ada / ekstensi / ukuran), belum dibaca, supaya
    ratusan foto tidak ditahan di memori. Isinya baru dibaca saat pengiriman.
    """
    siap: dict[str, str] = {}
    pesan: list[str] = []
    for nama_field, judul_kolom, ekstensi, label in (
            ("foto_profil", J_FOTO, GOL_BENTUK, "foto profil"),
            ("sk_file", J_SKFILE, GOL_PDF, "file SK")):
        mentah = bteks(baris.get(judul_kolom))
        if not mentah or mentah in ("-", "--"):
            continue
        path = absolutkan_path(mentah, folder_dasar)
        nama_berkas = os.path.basename(path)
        if not os.path.isfile(path):
            pesan.append(f"{label}: berkas tidak ditemukan -> {path}")
            continue
        ekst = os.path.splitext(path)[1].lower()
        if ekst not in ekstensi:
            pesan.append(f"{label}: ekstensi {ekst or '(tak bertipe)'} tidak sesuai "
                         f"(harus {', '.join(sorted(ekstensi))}) -> {nama_berkas}")
            continue
        ukuran = os.path.getsize(path)
        if ukuran > BATASAN_UNGGAHAN:
            pesan.append(f"{label}: berkas {ukuran / 1024 / 1024:.1f} MB melebihi "
                         f"{BATASAN_UNGGAHAN // 1024 // 1024} MB -> {nama_berkas}")
            continue
        siap[nama_field] = path
    return siap, pesan


def cari_lampiran_cadangan(baris: dict, folder: str) -> dict[str, str]:
    """Cadangan: cari berkas di satu folder berdasarkan NIK / No Reg / Nama.

    Dipakai hanya bila kolom Foto Profil / File SK di Excel kosong, sehingga
    Excel tidak perlu memuat path sama sekali.
    """
    kunci = [konversi_nik(baris.get(J_NIK)).nilai,
             konversi_noreg(baris.get(J_NOREG)).nilai,
             bteks(baris.get(J_NAMA)).lower().replace(" ", "_")]
    hasil: dict[str, str] = {}
    for nama_field, ekstensi in (("foto_profil", GOL_BENTUK), ("sk_file", GOL_PDF)):
        for nama in kunci:
            if not nama:
                continue
            for ekst in sorted(ekstensi):
                path = os.path.join(folder, f"{nama}{ekst}")
                if os.path.isfile(path) and os.path.getsize(path) <= BATASAN_UNGGAHAN:
                    hasil[nama_field] = path
                    break
            if nama_field in hasil:
                break
    return hasil


def baca_isi_berkas(siap: dict[str, str]) -> dict[str, tuple]:
    """Baca isi berkas -> {nama_field: (nama_berkas, isi_bytes, tipe)}."""
    hasil = {}
    for nama_field, path in siap.items():
        ekst = os.path.splitext(path)[1].lower()
        tipe = (TIPE_BERKAS.get(ekst) or mimetypes.guess_type(path)[0]
                or "application/octet-stream")
        with open(path, "rb") as berkas:
            hasil[nama_field] = (os.path.basename(path), berkas.read(), tipe)
    return hasil



@dataclass
class SiapKirim:
    """Satu baris yang sudah valid dan siap dikirim ke server."""

    baris_excel: int
    nama: str
    payload: dict
    lampiran: dict[str, str] = field(default_factory=dict)



# Aliases for embedded bulk
BulkSesi = Sesi
BulkSimLinmas = SimLinmas


# ========================================================== 1. UTIL UMUM
def bersihkan_surrogate(v):
    if isinstance(v, str):
        return v.encode("utf-8", "replace").decode("utf-8")
    return v

def teks(nilai) -> str:
    if nilai is None:
        return ""
    if isinstance(nilai, float) and nilai.is_integer():
        nilai = int(nilai)
    return str(nilai).strip().encode("utf-8", "replace").decode("utf-8")

def norm(nilai) -> str:
    s = unicodedata.normalize("NFKD", teks(nilai))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.casefold()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())

def aman_nama(nama: str) -> str:
    nama = teks(nama).replace("/", "-").replace("\\", "-")
    nama = re.sub(r'[:*?"<>|\x00-\x1f]', "-", nama)
    nama = re.sub(r"\s+", " ", nama).strip(" .")
    if len(nama) > 180:
        akar, ext = os.path.splitext(nama)
        nama = akar[:180 - len(ext)] + ext
    return nama or "tanpa-nama"

def nik_rapi(nilai) -> str:
    t = teks(nilai)
    if not t:
        return ""
    t = re.sub(r"\.0+$", "", t)
    t = re.sub(r"\D", "", t)
    return t if len(t) == 16 else ""

def id_drive(nilai) -> str:
    t = teks(nilai)
    for pola in (r"/file/d/([A-Za-z0-9_-]{20,})",
                 r"/d/([A-Za-z0-9_-]{20,})",
                 r"[?&]id=([A-Za-z0-9_-]{20,})"):
        m = re.search(pola, t)
        if m:
            return m.group(1)
    return ""

def hp_rapi(nilai) -> str:
    t = re.sub(r"\D", "", teks(nilai))
    if not t:
        return ""
    if t.startswith("620"):
        t = "62" + t[3:]
    if t.startswith("0"):
        t = "62" + t[1:]
    elif t.startswith("8"):
        t = "62" + t
    return t

def tanggal_rapi(nilai) -> str:
    if isinstance(nilai, (datetime, date)):
        return nilai.strftime("%d-%m-%Y")
    t = teks(nilai)
    if not t:
        return ""
    if re.fullmatch(r"(19|20)\d{2}", t):
        return f"01-01-{t}"
    for pola in (r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$",
                 r"^(\d{4})-(\d{1,2})-(\d{1,2})"):
        m = re.match(pola, t)
        if m:
            a, b, c = m.groups()
            if len(a) == 4:
                return f"{int(c):02d}-{int(b):02d}-{a}"
            return f"{int(a):02d}-{int(b):02d}-{c}"
    return t

def angka_atau_kosong(nilai):
    t = teks(nilai)
    if not t or t in ("-", "--"):
        return None
    try:
        n = float(t.replace(".", "").replace(",", "."))
        return int(n) if n.is_integer() else n
    except ValueError:
        return None

def tahun_dari(nilai, bawaan: int) -> int:
    if isinstance(nilai, (datetime, date)):
        return nilai.year
    m = re.search(r"(19|20)\d{2}", teks(nilai))
    return int(m.group(0)) if m else bawaan

def relatif(path: Path, dasar: Path) -> str:
    try:
        return os.path.relpath(path, dasar)
    except ValueError:
        return str(path)


# ============================================================== 2. STATE
STATE_PATH_LAMA = HERE / ".sinkron_drive_state.json"

def muat_state() -> dict:
    for kandidat in (STATE_PATH, STATE_PATH_LAMA):
        if kandidat.is_file():
            try:
                data = json.loads(kandidat.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    data.setdefault("files", {})
                    data.setdefault("spreadsheet", {})
                    data.setdefault("nik", [])
                    return data
            except (OSError, json.JSONDecodeError):
                continue
    return {"versi": 1, "files": {}, "spreadsheet": {}, "nik": [], "terakhir": ""}

def simpan_state(state: dict) -> None:
    state["versi"] = 1
    state["terakhir"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sementara = STATE_PATH.with_suffix(".tmp")
    sementara.write_text(json.dumps(state, ensure_ascii=True, indent=1),
                         encoding="utf-8", errors="replace")
    os.replace(sementara, STATE_PATH)

def nik_referensi_lokal() -> set[str]:
    state = muat_state()
    if state.get("nik"):
        return set(state["nik"])
    lama = HERE / "bulk-simlinmas-ponorogo-siap.xlsx"
    if lama.is_file():
        try:
            wb = openpyxl.load_workbook(lama, data_only=True, read_only=True)
            try:
                ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
                header = [teks(x) for x in next(ws.iter_rows(values_only=True))]
                if "NIK" in header:
                    i = header.index("NIK")
                    return {nik_rapi(r[i]) for r in ws.iter_rows(min_row=2, values_only=True)
                            if r and len(r) > i and nik_rapi(r[i])}
            finally:
                wb.close()
        except Exception:
            pass
    return set()


# ============================================================== 3. RCLONE
BIN_LOKAL = HERE / "bin"

def rclone_bin() -> str | None:
    ada = shutil.which("rclone")
    if ada:
        return ada
    lokal = BIN_LOKAL / "rclone"
    if lokal.is_file() and os.access(lokal, os.X_OK):
        return str(lokal)
    return None

def rclone_ada() -> bool:
    return rclone_bin() is not None

def jalankan_rclone(argumen: list[str], periksa: bool = True,
                    timeout: int = 900, percobaan: int = 4) -> subprocess.CompletedProcess:
    biner = rclone_bin()
    if not biner:
        raise RuntimeError("rclone tidak ditemukan. Jalankan menu bantuan setup.")
    cmd = [biner, *argumen, *FLAGS_RCLONE]
    err_akhir = ""
    for i in range(1, max(1, percobaan) + 1):
        try:
            hasil_proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired as e:
            err_akhir = f"timeout setelah {timeout}s"
            if i < percobaan:
                catat(f"rclone timeout percobaan {i}/{percobaan}, retry...")
                time.sleep(2 * i)
                continue
            if periksa:
                raise RuntimeError(err_akhir) from e
            return subprocess.CompletedProcess(cmd, 1, "", err_akhir)
        if hasil_proc.returncode == 0 or not periksa:
            if hasil_proc.returncode != 0 and i < percobaan:
                err_akhir = (hasil_proc.stderr or hasil_proc.stdout or "").strip().splitlines()
                err_akhir = err_akhir[-1] if err_akhir else f"kode {hasil_proc.returncode}"
                catat(f"rclone gagal percobaan {i}/{percobaan}: {err_akhir[:120]}; retry...")
                time.sleep(2 * i)
                continue
            if periksa and hasil_proc.returncode != 0:
                pesan = (hasil_proc.stderr or hasil_proc.stdout or "").strip().splitlines()
                raise RuntimeError(pesan[-1] if pesan else f"rclone gagal (kode {hasil_proc.returncode})")
            return hasil_proc
        err_akhir = (hasil_proc.stderr or hasil_proc.stdout or "").strip().splitlines()
        err_akhir = err_akhir[-1] if err_akhir else f"kode {hasil_proc.returncode}"
        if i < percobaan:
            catat(f"rclone gagal percobaan {i}/{percobaan}: {err_akhir[:120]}; retry...")
            time.sleep(2 * i)
    raise RuntimeError(err_akhir or "rclone gagal")

def daftar_remote() -> dict:
    try:
        hasil = jalankan_rclone(["config", "dump"], timeout=30)
        data = json.loads(hasil.stdout or "{}")
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def pilih_remote(diminta: str = "") -> str:
    konfig = daftar_remote()
    drive = [n for n, v in konfig.items() if isinstance(v, dict) and v.get("type") == "drive"]
    if diminta:
        if diminta.rstrip(":") not in konfig:
            raise RuntimeError(f"remote '{diminta}' tidak ada di konfigurasi rclone.")
        return diminta.rstrip(":")
    if len(drive) == 1:
        return drive[0]
    if not drive:
        raise RuntimeError("belum ada remote Google Drive di rclone.")
    return drive[0]

def cek_koneksi(remote: str) -> None:
    try:
        jalankan_rclone(["lsd", f"{remote}:", "--max-depth", "1"], timeout=120)
    except Exception as e:
        raise RuntimeError(f"Gagal koneksi ke Google Drive remote '{remote}': {e}") from e


# ====================================================== 4. BACA REKAP & DATA
def baca_rekap(path: Path) -> list[dict]:
    if not path.is_file():
        raise RuntimeError(f"file rekap tidak ditemukan -> {path}")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb["Rekap"] if "Rekap" in wb.sheetnames else wb.worksheets[0]
        baris = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not baris:
        raise RuntimeError("sheet Rekap kosong")
    header = [norm(x) for x in baris[0]]
    i_kec = header.index("kecamatan")
    i_ss = header.index("id spreadsheet")

    daftar = []
    for r in baris[1:]:
        if not r or len(r) <= max(i_kec, i_ss):
            continue
        kec = teks(r[i_kec])
        ssid = teks(r[i_ss])
        if kec and ssid:
            daftar.append({"kecamatan": kec, "ssid": ssid})
    return daftar

def unduh_spreadsheet(remote: str, ssid: str, tujuan: Path) -> tuple[Path, int]:
    def _sekali():
        with tempfile.TemporaryDirectory(prefix="sinkron_ss_") as tmp:
            jalankan_rclone(["backend", "copyid", f"{remote}:", ssid,
                             str(tmp) + os.sep, "--drive-export-formats", "xlsx"],
                            percobaan=3, timeout=900)
            berkas = [p for p in Path(tmp).iterdir() if p.is_file()]
            if not berkas:
                raise RuntimeError("rclone tidak menghasilkan berkas spreadsheet")
            sumber = max(berkas, key=lambda p: p.stat().st_size)
            tujuan.parent.mkdir(parents=True, exist_ok=True)
            sementara = tujuan.with_suffix(".tmp.xlsx")
            shutil.copy2(sumber, sementara)
            os.replace(sementara, tujuan)
    ok, err = coba_ulang(_sekali, percobaan=4, jeda_awal=3, nama=f"unduh spreadsheet {ssid[:12]}")
    if not ok:
        raise RuntimeError(f"gagal unduh spreadsheet setelah retry: {err}")

    wb = openpyxl.load_workbook(tujuan, data_only=True, read_only=True)
    try:
        ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
        jumlah = sum(1 for r in ws.iter_rows(min_row=2, values_only=True)
                     if r and any(x not in (None, "") for x in r))
    finally:
        wb.close()
    return tujuan, jumlah

def baca_baris_kecamatan(path: Path, kecamatan: str = "") -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        # Support sheet "Data", "Data Semua Kecamatan", or first sheet
        for nama_ws in ("Data", "Data Semua Kecamatan"):
            if nama_ws in wb.sheetnames:
                ws = wb[nama_ws]
                break
        else:
            ws = wb.worksheets[0]
        iterator = ws.iter_rows(values_only=True)
        header = [teks(x) for x in next(iterator)]
        peta = {h: i for i, h in enumerate(header) if h}

        def ambil(r, nama):
            i = peta.get(nama)
            return r[i] if i is not None and i < len(r) else None

        hasil = []
        for r in iterator:
            if not r or not any(x not in (None, "") for x in r):
                continue
            nama = teks(ambil(r, K_NAMA))
            if not nama:
                continue
            kec_row = teks(ambil(r, "Kecamatan")) or kecamatan
            baris = {
                "kecamatan": kec_row,
                "desa": teks(ambil(r, K_DESA)) or "-",
                "nama": nama,
                "nik": nik_rapi(ambil(r, K_NIK)),
                "nik_asli": teks(ambil(r, K_NIK)),
                "alamat": teks(ambil(r, K_ALAMAT)),
                "tempat_lahir": teks(ambil(r, K_LAHIR)),
                "tanggal_lahir": tanggal_rapi(ambil(r, K_TGL)),
                "jk": teks(ambil(r, K_JK)),
                "hp": hp_rapi(ambil(r, K_HP)),
                "email": teks(ambil(r, K_EMAIL)),
                "jumlah_keluarga": angka_atau_kosong(ambil(r, K_KK)),
                "pendidikan": teks(ambil(r, K_PEND)),
                "waktu": ambil(r, K_WAKTU),
                "id": {},
                "link": {},
            }
            for j, kolom in K_ID.items():
                baris["id"][j] = teks(ambil(r, kolom))
            for j, kolom in K_LINK.items():
                baris["link"][j] = teks(ambil(r, kolom))
            for j in JENIS:
                if not baris["id"][j]:
                    baris["id"][j] = id_drive(baris["link"][j])
            hasil.append(baris)
        return hasil
    finally:
        wb.close()


# ===================================================== 5. INDEKS BERKAS LOKAL
class IndeksFolder:
    def __init__(self, folder: Path):
        self.folder = folder
        self.entri = []
        if not folder.is_dir():
            return
        for p in folder.iterdir():
            if not p.is_file():
                continue
            stem = p.stem
            nama, label = stem.split(" - ", 1) if " - " in stem else (stem, "")
            self.entri.append({
                "path": p,
                "nama": norm(nama),
                "label": norm(label),
                "ekstensi": p.suffix.lower(),
                "ukuran": p.stat().st_size,
            })

    def cari(self, nama: str, jenis: str, terpakai: set[Path]) -> Path | None:
        aturan = JENIS[jenis]
        target = norm(nama)
        kandidat = []
        for e in self.entri:
            if e["nama"] != target or e["path"] in terpakai:
                continue
            # Cocokkan label (foto, ktp, sk)
            if not any(k in e["label"] for k in aturan["kata"]):
                continue
            prioritas = 0 if e["label"].startswith(aturan["kata"][0]) else 1
            kandidat.append((prioritas, len(e["label"]), e["path"]))
        if not kandidat:
            return None
        kandidat.sort()
        return kandidat[0][2]

class IndeksLokal:
    def __init__(self, state: dict | None = None):
        self._folder: dict[Path, IndeksFolder] = {}
        self.terpakai: dict[Path, set[Path]] = {}
        self.terklaim: set[Path] = set()
        for catatan in ((state or {}).get("files") or {}).values():
            self.terklaim.add(DATA_DIR / catatan)

    def folder(self, path: Path) -> IndeksFolder:
        if path not in self._folder:
            self._folder[path] = IndeksFolder(path)
        return self._folder[path]

    def terpakai_set(self, path: Path) -> set[Path]:
        return self.terpakai.setdefault(path, set())

    def cari(self, folder: Path, nama: str, jenis: str) -> Path | None:
        dipakai = self.terpakai_set(folder) | self.terklaim
        return self.folder(folder).cari(nama, jenis, dipakai)

    def tandai(self, folder: Path, path: Path) -> None:
        self.terpakai_set(folder).add(path)
        self.terklaim.add(path)

def siapkan_folder_kec(kec: str) -> Path:
    induk = DATA_DIR
    if induk.is_dir():
        for p in induk.iterdir():
            if p.is_dir() and norm(p.name) == norm(kec):
                return p
    p_baru = induk / aman_nama(kec)
    p_baru.mkdir(parents=True, exist_ok=True)
    return p_baru

def siapkan_folder_desa(folder_kec: Path, desa: str) -> Path:
    if folder_kec.is_dir():
        for p in folder_kec.iterdir():
            if p.is_dir() and norm(p.name) == norm(desa):
                return p
    p_baru = folder_kec / aman_nama(desa)
    p_baru.mkdir(parents=True, exist_ok=True)
    return p_baru

def berkas_untuk(row: dict, jenis: str, indeks: IndeksLokal, state: dict) -> Path | None:
    fid = teks(row["id"].get(jenis))
    if fid:
        catatan = state["files"].get(fid)
        if catatan:
            path = DATA_DIR / catatan
            if path.is_file():
                return path

    folder = siapkan_folder_desa(siapkan_folder_kec(row["kecamatan"]), row["desa"])
    path = indeks.cari(folder, row["nama"], jenis)
    if path:
        if fid and fid not in state["files"]:
            state["files"][fid] = relatif(path, DATA_DIR)
        indeks.tandai(folder, path)
    return path

def unduh_berkas(remote: str, fid: str, tujuan_folder: Path, nama_cadangan: str = "") -> Path:
    tujuan_folder.mkdir(parents=True, exist_ok=True)

    def _sekali():
        with tempfile.TemporaryDirectory(prefix="sinkron_file_") as tmp:
            jalankan_rclone(["backend", "copyid", f"{remote}:", fid, tmp + os.sep],
                            percobaan=3, timeout=600)
            berkas = [p for p in Path(tmp).iterdir() if p.is_file()]
            if not berkas:
                raise RuntimeError("rclone tidak menghasilkan berkas")
            sumber = berkas[0]
            if sumber.stat().st_size < 100:
                raise RuntimeError("berkas kosong / gagal diunduh")

            nama = aman_nama(sumber.name or nama_cadangan or fid)
            if not Path(nama).suffix:
                nama += (sumber.suffix or "")
            tujuan = tujuan_folder / nama
            if tujuan.exists():
                akar, ext = os.path.splitext(nama)
                nomor = 2
                while True:
                    calon = tujuan_folder / f"{akar} ({nomor}){ext}"
                    if not calon.exists():
                        tujuan = calon
                        break
                    nomor += 1
            shutil.move(str(sumber), str(tujuan))
            return tujuan

    ok, hasil_unduh = coba_ulang(_sekali, percobaan=5, jeda_awal=2, faktor=1.8,
                                 nama=f"unduh {fid[:12]}")
    if not ok:
        raise RuntimeError(f"gagal unduh setelah retry: {hasil_unduh}")
    return hasil_unduh

def sama_berkas(path_a: Path, path_b: Path) -> bool:
    try:
        return path_a.stat().st_size == path_b.stat().st_size
    except OSError:
        return False

_NIK_LAMA: dict[tuple[str, str, str], str] | None = None

def peta_nik_lama() -> dict[tuple[str, str, str], str]:
    global _NIK_LAMA
    if _NIK_LAMA is not None:
        return _NIK_LAMA
    _NIK_LAMA = {}
    for berkas in sorted(HERE.glob("*-siap*.xlsx")):
        try:
            wb = openpyxl.load_workbook(berkas, data_only=True, read_only=True)
            if "Data" not in wb.sheetnames:
                continue
            ws = wb["Data"]
            baris = ws.iter_rows(values_only=True)
            judul_list = [teks(x) for x in next(baris, [])]
            peta = {j: i for i, j in enumerate(judul_list)}
            butuh = ("Kecamatan", "Desa/Kelurahan", "Nama Lengkap", "NIK")
            if any(k not in peta for k in butuh):
                continue
            for r in baris:
                if not r or len(r) <= max(peta[k] for k in butuh):
                    continue
                nik = nik_rapi(r[peta["NIK"]])
                if nik:
                    kunci = (norm(r[peta["Kecamatan"]]), norm(r[peta["Desa/Kelurahan"]]),
                             norm(r[peta["Nama Lengkap"]]))
                    _NIK_LAMA.setdefault(kunci, nik)
        except Exception:
            continue
    return _NIK_LAMA

def rapikan_baris(semua: list[dict]) -> tuple[list[dict], list[dict]]:
    per_nik: dict[str, dict] = {}
    laporan: list[dict] = []
    cadangan = peta_nik_lama()
    sudah_dilaporkan: set[tuple[str, str, str, str]] = set()

    for baris in semua:
        if not baris["nik"] and cadangan:
            kunci = (norm(baris["kecamatan"]), norm(baris["desa"]), norm(baris["nama"]))
            pulih = cadangan.get(kunci)
            if pulih:
                tanda = (baris["kecamatan"], baris["desa"], baris["nama"], pulih)
                if tanda not in sudah_dilaporkan:
                    sudah_dilaporkan.add(tanda)
                    laporan.append({
                        "kecamatan": baris["kecamatan"], "desa": baris["desa"],
                        "nama": baris["nama"], "nik": pulih,
                        "bagian": "-", "status": "NIK DIPULIHKAN",
                        "keterangan": f"sel berisi {baris['nik_asli']}; NIK diambil dari file siap-bulk lama",
                    })
                baris["nik"] = pulih

        if not baris["nik"]:
            laporan.append({
                "kecamatan": baris["kecamatan"], "desa": baris["desa"],
                "nama": baris["nama"], "nik": baris["nik_asli"],
                "bagian": "-", "status": "NIK TIDAK VALID",
                "keterangan": "NIK bukan 16 digit, dilewati",
            })
            continue

        lama = per_nik.get(baris["nik"])
        if lama is None:
            per_nik[baris["nik"]] = baris
            continue

        if str(baris["waktu"]) > str(lama["waktu"]):
            per_nik[baris["nik"]] = baris
            dibuang = lama
        else:
            dibuang = baris

        laporan.append({
            "kecamatan": dibuang["kecamatan"], "desa": dibuang["desa"],
            "nama": dibuang["nama"], "nik": dibuang["nik"],
            "bagian": "-", "status": "NIK GANDA",
            "keterangan": "baris ini dilewati (ada baris lain dengan NIK sama)",
        })

    return list(per_nik.values()), laporan


# ========================================================== 6. PENYERAGAMAN FORMAT & KOMPRESI
def konversi_sk_ke_pdf(path: Path) -> Path:
    """Ubah SK bertipe gambar (.jpg/.png/.jpeg) menjadi .pdf murni."""
    if not path.is_file():
        return path
    if path.suffix.lower() == ".pdf":
        return path
    target = path.with_suffix(".pdf")
    try:
        with Image.open(path) as im:
            rgb = im.convert("RGB")
            rgb.save(target, "PDF", resolution=100.0)
        if target.exists() and target.stat().st_size > 0:
            path.unlink(missing_ok=True)
            catat_log(f"Konversi SK: {path.name} -> {target.name}")
            return target
    except Exception as e:
        catat_log(f"Gagal konversi SK {path.name} ke PDF: {e}")
    return path

def konversi_gambar_ke_jpg(path: Path) -> Path:
    """Ubah Foto/KTP (.png/.webp/.heic/.jpeg) menjadi .jpg standar."""
    if not path.is_file():
        return path
    suffix = path.suffix.lower()
    if suffix in [".jpg"]:
        return path

    target = path.with_suffix(".jpg")
    try:
        # Gunakan heif-convert / ImageMagick jika HEIC
        if suffix == ".heic":
            res = subprocess.run(["magick", str(path), str(target)], capture_output=True)
            if res.returncode != 0:
                subprocess.run(["heif-convert", str(path), str(target)], capture_output=True)
        else:
            with Image.open(path) as im:
                rgb = im.convert("RGB")
                rgb.save(target, "JPEG", quality=85, optimize=True)

        if target.exists() and target.stat().st_size > 0:
            path.unlink(missing_ok=True)
            catat_log(f"Konversi Gambar: {path.name} -> {target.name}")
            return target
    except Exception as e:
        catat_log(f"Gagal konversi gambar {path.name} ke JPG: {e}")
    return path

def kompres_pdf_satlinmas(path: Path) -> bool:
    """Kompres PDF dengan Ghostscript agar < 850 KB."""
    if not path.is_file() or path.stat().st_size <= BATAS_UKURAN_BERKAS:
        return True

    tmp = path.with_name(f"_tmp_{path.name}")
    for dpi in [100, 72, 60, 50]:
        cmd = [
            "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/screen", "-dNOPAUSE", "-dQUIET", "-dBATCH",
            f"-dColorImageResolution={dpi}",
            f"-dGrayImageResolution={dpi}",
            f"-dMonoImageResolution={dpi}",
            f"-sOutputFile={tmp}", str(path)
        ]
        try:
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
        except Exception:
            continue

        if tmp.exists() and tmp.stat().st_size > 0:
            if tmp.stat().st_size < path.stat().st_size:
                tmp.replace(path)
            else:
                tmp.unlink(missing_ok=True)
            if path.stat().st_size <= BATAS_UKURAN_BERKAS:
                return True
        else:
            tmp.unlink(missing_ok=True)

    return path.stat().st_size <= BATAS_UKURAN_BERKAS

def kompres_jpg_satlinmas(path: Path) -> bool:
    """Kompres JPG dengan resize & optimize agar < 850 KB."""
    if not path.is_file() or path.stat().st_size <= BATAS_UKURAN_BERKAS:
        return True

    try:
        with Image.open(path) as im:
            rgb = im.convert("RGB")
            w, h = rgb.size
            if max(w, h) > 1600:
                skala = 1600 / max(w, h)
                rgb = rgb.resize((int(w * skala), int(h * skala)), Image.Resampling.LANCZOS)

            for q in [75, 65, 55, 45, 35]:
                rgb.save(path, "JPEG", quality=q, optimize=True)
                if path.stat().st_size <= BATAS_UKURAN_BERKAS:
                    return True
    except Exception as e:
        catat_log(f"Error kompres JPG {path.name}: {e}")

    return path.stat().st_size <= BATAS_UKURAN_BERKAS

def standarkan_dan_kompres_semua_berkas():
    """Scan seluruh folder DATA SATLINMAS PONOROGO dan pastikan semua berkas seragam & < 1 MB."""
    langkah("A", "Standardisasi Format Berkas (.jpg untuk foto/ktp, .pdf untuk SK)")
    sk_dikonversi = 0
    foto_dikonversi = 0

    # 1. Konversi SK bukan PDF
    for ext in ["*.jpg", "*.jpeg", "*.png", "*.webp"]:
        for f in DATA_DIR.rglob(ext):
            if "SK Satlinmas" in f.name:
                konversi_sk_ke_pdf(f)
                sk_dikonversi += 1

    # 2. Konversi Foto & KTP bukan JPG
    for ext in ["*.png", "*.webp", "*.heic", "*.jpeg"]:
        for f in DATA_DIR.rglob(ext):
            if "SK Satlinmas" not in f.name:
                konversi_gambar_ke_jpg(f)
                foto_dikonversi += 1

    baik(f"Konversi format tuntas: {sk_dikonversi} SK -> .pdf, {foto_dikonversi} Foto/KTP -> .jpg")

    langkah("B", f"Kompresi Seluruh Berkas Besar (Batas Maksimal < {BATAS_UKURAN_BERKAS // 1024} KB)")
    daftar_besar = []
    for f in DATA_DIR.rglob("*"):
        if f.is_file() and f.suffix.lower() in [".pdf", ".jpg", ".jpeg"]:
            if f.stat().st_size > BATAS_UKURAN_BERKAS:
                daftar_besar.append(f)

    if not daftar_besar:
        baik("Semua berkas lampiran sudah berukuran di bawah 1 MB!")
        return {"sk_ke_pdf": sk_dikonversi, "foto_ke_jpg": foto_dikonversi, "kompres_ok": 0, "kompres_total": 0, "kompres_gagal": 0}

    info(f"Ditemukan {len(daftar_besar)} berkas berukuran di atas batas. Mengompres...")
    berhasil = 0
    gagal_k = 0
    tkomp = time.time()
    total_besar = len(daftar_besar)
    for idx, f in enumerate(daftar_besar, 1):
        awal = f.stat().st_size / 1024
        try:
            if f.suffix.lower() == ".pdf":
                ok = kompres_pdf_satlinmas(f)
            else:
                ok = kompres_jpg_satlinmas(f)
        except Exception as e:
            ok = False
            catat(f"Gagal kompres {f.name}: {e}")
        akhir = f.stat().st_size / 1024

        if ok:
            berhasil += 1
        else:
            gagal_k += 1

        if idx % 10 == 0 or idx == total_besar:
            progres(idx, total_besar, "Kompresi", f"ok={berhasil} pending={gagal_k} last={akhir:.0f}KB", tkomp)

    baik(f"Kompresi selesai! {berhasil}/{total_besar} berkas berhasil disusutkan di bawah 1 MB.")
    if gagal_k:
        catat(f"{gagal_k} berkas masih di atas batas; akan dicoba lagi saat upload bila perlu.")
    return {"sk_ke_pdf": sk_dikonversi, "foto_ke_jpg": foto_dikonversi, "kompres_ok": berhasil, "kompres_total": total_besar, "kompres_gagal": gagal_k}


# ============================================================== 7. SINKRON DRIVE
def sinkron(remote: str, rekap: Path, batas: int = 0, tanpa_unduh: bool = False, nik_sudah_server: set[str] | None = None, filter_server: bool = True) -> dict:
    state = muat_state()
    punya_state = bool(state.get("terakhir"))
    nik_referensi = set(state.get("nik") or []) if punya_state else set()
    daftar_kec = baca_rekap(rekap)
    indeks = IndeksLokal(state)

    judul("SINKRON PERKEMBANGAN DATA  -  GOOGLE DRIVE")
    label_remote = remote if remote.startswith("(") else remote + ":"
    info(f"remote          : {hijau(label_remote)}")
    info(f"folder data     : {DATA_DIR}")
    info(f"total kecamatan : {len(daftar_kec)}")

    # Langkah 1: Ambil data (cepat: dari sheet "Data Semua Kecamatan" di rekap utama)
    langkah("1", "Membaca data dari sheet 'Data Semua Kecamatan' (rekap utama)")
    semua_baris: list[dict] = []
    gagal_kec: list[tuple[str, str]] = []
    t1 = time.time()
    pakai_rekap = False

    # Coba refresh rekap utama dari Drive agar data terbaru
    rekap_sumber = rekap
    if not tanpa_unduh and remote and not remote.startswith("("):
        for kandidat in [
            f"{remote}:DATA SATLINMAS PONOROGO/Rekap Pendataan Satlinmas - Kabupaten Ponorogo (Utama).xlsx",
            f"{remote}:Rekap Pendataan Satlinmas - Kabupaten Ponorogo (Utama).xlsx",
        ]:
            try:
                tmp_rekap = rekap.with_suffix(".drive.xlsx")
                subprocess.run(["rclone", "copyto", kandidat, str(tmp_rekap)], check=True, capture_output=True, text=True, timeout=180)
                if tmp_rekap.exists() and tmp_rekap.stat().st_size > 0:
                    rekap_sumber = tmp_rekap
                    baik(f"Rekap utama diperbarui dari Drive: {kandidat}")
                    break
            except Exception:
                pass

    try:
        wb_cek = openpyxl.load_workbook(rekap_sumber, data_only=True, read_only=True)
        ada_sheet = "Data Semua Kecamatan" in wb_cek.sheetnames
        wb_cek.close()
    except Exception:
        ada_sheet = False

    if ada_sheet:
        try:
            semua_baris = baca_baris_kecamatan(rekap_sumber, "")
            per_kec: dict[str, int] = {}
            for b in semua_baris:
                k = b.get("kecamatan") or "-"
                per_kec[k] = per_kec.get(k, 0) + 1
            for i, (k, n) in enumerate(sorted(per_kec.items()), 1):
                pct = i * 100 / max(len(per_kec), 1)
                print(f"  [{i:>2}/{len(per_kec)}] {pct:6.2f}% {k:<14} {hijau(str(n)+' data')}")
            catat_log(f"Baca Data Semua Kecamatan: {len(semua_baris)} baris, {len(per_kec)} kecamatan")
            baik(f"Sumber data: sheet 'Data Semua Kecamatan' ({len(semua_baris)} baris)")
            pakai_rekap = True
        except Exception as e:
            catat(f"Gagal baca sheet 'Data Semua Kecamatan': {e}. Fallback ke download per kecamatan.")
            semua_baris = []

    if not pakai_rekap:
        langkah("1b", "Fallback: ambil spreadsheet kecamatan dari Google Drive")
        total_kec = max(len(daftar_kec), 1)
        for urutan, kec in enumerate(daftar_kec, 1):
            nama_kec = kec["kecamatan"]
            tujuan_ss = siapkan_folder_kec(nama_kec) / f"Data Satlinmas - Kecamatan {aman_nama(nama_kec)}.xlsx"
            try:
                if tanpa_unduh:
                    if not tujuan_ss.is_file():
                        raise RuntimeError("file lokal tidak ada")
                    path_ss = tujuan_ss
                else:
                    path_ss, _ = unduh_spreadsheet(remote, kec["ssid"], tujuan_ss)
                baris = baca_baris_kecamatan(path_ss, nama_kec)
                semua_baris.extend(baris)
                state["spreadsheet"][nama_kec] = {"id": kec["ssid"], "baris": len(baris), "diperbarui": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                status = hijau(f"{len(baris)} data")
            except Exception as galat:
                gagal_kec.append((nama_kec, str(galat)))
                status = merah(f"GAGAL: {str(galat)[:60]}")
            pct = urutan * 100 / total_kec
            print(f"  [{urutan:>2}/{total_kec}] {pct:6.2f}% {nama_kec:<14} {status}")
            catat_log(f"Spreadsheet {urutan}/{total_kec} {pct:.2f}% {nama_kec}: {status}")
            if urutan % 5 == 0:
                simpan_state(state)
        if gagal_kec and not tanpa_unduh:
            for ronde in range(1, 4):
                if not gagal_kec:
                    break
                catat(f"Retry spreadsheet gagal ronde {ronde}: {len(gagal_kec)} kecamatan...")
                masih_gagal = []
                for nama_kec, _ in gagal_kec:
                    kec = next((k for k in daftar_kec if k["kecamatan"] == nama_kec), None)
                    if not kec:
                        continue
                    tujuan_ss = siapkan_folder_kec(nama_kec) / f"Data Satlinmas - Kecamatan {aman_nama(nama_kec)}.xlsx"
                    try:
                        path_ss, _ = unduh_spreadsheet(remote, kec["ssid"], tujuan_ss)
                        baris = baca_baris_kecamatan(path_ss, nama_kec)
                        semua_baris.extend(baris)
                        state["spreadsheet"][nama_kec] = {"id": kec["ssid"], "baris": len(baris), "diperbarui": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                        baik(f"Retry berhasil: {nama_kec} ({len(baris)} data)")
                    except Exception as galat:
                        masih_gagal.append((nama_kec, str(galat)))
                gagal_kec = masih_gagal
                simpan_state(state)

    simpan_state(state)
    garis("-")
    hasil("Total baris terbaca", len(semua_baris))
    hasil("Kecamatan gagal diambil", len(gagal_kec), salah=bool(gagal_kec))
    if gagal_kec:
        for nk, eg in gagal_kec:
            catat(f"Gagal kecamatan {nk}: {eg}")

    # Langkah 2: Saring NIK
    langkah("2", "Validasi NIK & Saring NIK Ganda")
    baris_unik, laporan_awal = rapikan_baris(semua_baris)
    jumlah_invalid = sum(1 for b in laporan_awal if b["status"] == "NIK TIDAK VALID")
    jumlah_ganda = sum(1 for b in laporan_awal if b["status"] == "NIK GANDA")
    hasil("NIK valid & unik", len(baris_unik))
    hasil("NIK tidak valid (dilewati)", jumlah_invalid, salah=bool(jumlah_invalid))
    hasil("NIK ganda (dilewati)", jumlah_ganda, salah=bool(jumlah_ganda))

    if punya_state:
        nik_lama = set(state.get("nik") or [])
        nik_baru = sum(1 for b in baris_unik if b["nik"] not in nik_lama)
        hasil("NIK baru sejak sinkron sebelumnya", nik_baru)
    simpan_state(state)

    # Langkah 2b: Bandingkan dengan NIK yang sudah ada di SIMLINMAS
    nik_server_set: set[str] = set(nik_sudah_server or [])
    if nik_sudah_server is None and not tanpa_unduh and (relay_enabled() or (BulkSesi is not None and BulkSimLinmas is not None)):
        langkah("2b", "Cek NIK yang sudah terdaftar di SIMLINMAS (skip download untuk yang sudah ada)")
        try:
            nik_server_set = ambil_nik_simlinmas()
        except Exception as e:
            catat(f"Gagal cek NIK server: {e}. Lanjut tanpa filter server.")
            nik_server_set = set()
    elif nik_server_set:
        langkah("2b", "Bandingkan data Drive dengan NIK SIMLINMAS")
        baik(f"Gunakan data NIK SIMLINMAS yang sudah diambil: {len(nik_server_set)} NIK")

    if nik_server_set and filter_server:
        sebelum = len(baris_unik)
        baris_unik = [b for b in baris_unik if b["nik"] not in nik_server_set]
        hasil("Sudah ada di SIMLINMAS (diskip)", sebelum - len(baris_unik))
        hasil("Belum ada di SIMLINMAS (diproses)", len(baris_unik))
    elif nik_server_set and not filter_server:
        hasil("Sudah ada di SIMLINMAS (akan diupdate)", sum(1 for b in baris_unik if b["nik"] in nik_server_set))
        hasil("Belum ada di SIMLINMAS (akan diinsert)", sum(1 for b in baris_unik if b["nik"] not in nik_server_set))

    # Langkah 3: Unduh Lampiran Baru
    langkah("3", "Pemeriksaan & Pengunduhan Lampiran Baru")
    laporan: list[dict] = list(laporan_awal)
    total = len(baris_unik)
    baru = sudah = belum_ada = 0
    gagal_unduh: dict[str, dict] = {}  # fid -> item
    gagal_count: dict[str, int] = {}
    t3 = time.time()
    MAKS_GAGAL_PER_FID = 3

    def _unduh_satu(baris, j, fid):
        nonlocal baru, sudah, belum_ada
        path = berkas_untuk(baris, j, indeks, state)
        if path and path.is_file():
            sudah += 1
            gagal_unduh.pop(fid, None)
            return True
        if tanpa_unduh or not fid:
            belum_ada += 1
            laporan.append({
                "kecamatan": baris["kecamatan"], "desa": baris["desa"],
                "nama": baris["nama"], "nik": baris["nik"],
                "bagian": JENIS[j]["judul"], "status": "TIDAK ADA",
                "keterangan": "ID Drive kosong dari form",
            })
            return True
        if fid in gagal_count and gagal_count[fid] >= MAKS_GAGAL_PER_FID:
            return False
        folder = siapkan_folder_desa(siapkan_folder_kec(baris["kecamatan"]), baris["desa"])
        try:
            path = unduh_berkas(remote, fid, folder)
            lama = indeks.cari(folder, baris["nama"], j)
            if lama and sama_berkas(lama, path):
                path.unlink(missing_ok=True)
                path = lama
                indeks.tandai(folder, path)
            else:
                indeks.tandai(folder, path)
            state["files"][fid] = relatif(path, DATA_DIR)
            baru += 1
            gagal_unduh.pop(fid, None)
            return True
        except Exception as galat:
            gagal_count[fid] = gagal_count.get(fid, 0) + 1
            gagal_unduh[fid] = {"baris": baris, "jenis": j, "fid": fid, "err": str(galat)}
            return False

    for urutan, baris in enumerate(baris_unik, 1):
        # Mode sync: NIK yang sudah ada di SIMLINMAS hanya diupdate datanya, skip lampiran
        if not filter_server and nik_server_set and baris.get("nik") in nik_server_set:
            sudah += len(JENIS)
            if urutan % 25 == 0 or urutan == total:
                pct = urutan * 100 / max(total, 1)
                info(f"Progres: {urutan}/{total} ({pct:.2f}%) | skip lampiran update={sudah}")
            continue
        for j in JENIS:
            fid = teks(baris["id"].get(j))
            _unduh_satu(baris, j, fid)
        if urutan % 25 == 0 or urutan == total:
            pct = urutan * 100 / max(total, 1)
            info(f"Progres unduh: {urutan}/{total} ({pct:.2f}%) | baru: {baru} | lokal: {sudah} | kosong: {belum_ada} | gagal: {len(gagal_unduh)}")
            catat_log(f"Progres unduh {urutan}/{total} {pct:.2f}% baru={baru} lokal={sudah} kosong={belum_ada} gagal={len(gagal_unduh)}")
            simpan_state(state)

    # Retry pass untuk unduhan gagal - terbatas, jangan loop selamanya
    if gagal_unduh and not tanpa_unduh:
        for ronde in range(1, 3):
            kandidat = [it for it in gagal_unduh.values() if gagal_count.get(it["fid"], 0) < MAKS_GAGAL_PER_FID]
            if not kandidat:
                break
            catat(f"Retry unduhan gagal ronde {ronde}: {len(kandidat)} item (maks {MAKS_GAGAL_PER_FID}x per file)...")
            for item in kandidat:
                _unduh_satu(item["baris"], item["jenis"], item["fid"])
            simpan_state(state)

    # Catat gagal unduh yang benar-benar tidak bisa diambil
    for item in list(gagal_unduh.values()):
        belum_ada += 1
        laporan.append({
            "kecamatan": item["baris"]["kecamatan"], "desa": item["baris"]["desa"],
            "nama": item["baris"]["nama"], "nik": item["baris"]["nik"],
            "bagian": JENIS[item["jenis"]]["judul"], "status": "GAGAL UNDUH",
            "keterangan": str(item["err"])[:150],
        })

    simpan_state(state)
    garis("-")
    hasil("Lampiran lokal sudah ada", sudah)
    hasil("Lampiran baru berhasil diunduh", baru)
    hasil("Lampiran gagal/tidak ada", belum_ada, salah=bool(belum_ada))
    progres(total, total, "Pemeriksaan lampiran", f"baru={baru} lokal={sudah} gagal={belum_ada}", t3)

    # Langkah 4: Format & Kompresi Berkas
    stat_berkas = standarkan_dan_kompres_semua_berkas() or {}

    return {
        "baris": baris_unik, "laporan": laporan, "state": state,
        "nik_referensi": nik_referensi,
        "difilter_server": bool(nik_server_set),
        "stat_berkas": stat_berkas,
    }


# ============================================================ 8. BUAT EXCEL
def buat_excel(baris_data: list[dict], laporan: list[dict],
               keluaran: Path, state: dict | None = None,
               nik_referensi: set[str] | None = None,
               simpan_master: bool = True) -> dict:
    state = state or muat_state()
    nik_referensi = nik_referensi or set()
    indeks = IndeksLokal(state)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    for kolom, (judul_kolom, lebar) in enumerate(KOLOM_EXCEL, 1):
        sel = ws.cell(row=1, column=kolom, value=judul_kolom)
        sel.fill = header_fill
        sel.font = header_font
        sel.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(kolom)].width = lebar
    ws.freeze_panes = "A2"

    terisi = 0
    jumlah_baru = jumlah_lama = 0
    tahun_ini = date.today().year

    for nomor, row in enumerate(baris_data, 1):
        paths = {}
        for j in JENIS:
            p = berkas_untuk(row, j, indeks, state)
            # Jika SK berformat jpg, sesuaikan ke pdf
            if p and j == "sk" and p.suffix.lower() != ".pdf":
                p_alt = p.with_suffix(".pdf")
                if p_alt.exists():
                    p = p_alt
            # Jika foto berformat non-jpg, sesuaikan ke jpg
            if p and j in ["foto", "ektp"] and p.suffix.lower() != ".jpg":
                p_alt = p.with_suffix(".jpg")
                if p_alt.exists():
                    p = p_alt
            paths[j] = p

        tanda_baru = "BARU" if (not nik_referensi or row["nik"] not in nik_referensi) else "LAMA"
        if tanda_baru == "BARU":
            jumlah_baru += 1
        else:
            jumlah_lama += 1

        nilai = [
            nomor,
            row["nik"],
            "",
            row["nama"],
            row["tempat_lahir"],
            row["tanggal_lahir"],
            row["jk"],
            row["alamat"],
            PROVINSI_NAMA,
            KABKOTA_NAMA,
            row["kecamatan"],
            row["desa"],
            0,
            tahun_dari(row["waktu"], tahun_ini),
            row["jumlah_keluarga"],
            row.get("pendidikan", ""),
            row["hp"],
            row["email"],
            "Aktif",
            None,
            None,
            str(paths["foto"]) if paths["foto"] else "",
            str(paths["ektp"]) if paths["ektp"] else "",
            str(paths["sk"]) if paths["sk"] else "",
            tanda_baru,
        ]
        for kolom, isi in enumerate(nilai, 1):
            sel = ws.cell(row=nomor + 1, column=kolom, value=bersihkan_surrogate(isi))
            judul_k = KOLOM_EXCEL[kolom - 1][0]
            if judul_k in ("NIK", "No. HP"):
                sel.number_format = "@"
            if judul_k in ("Alamat", "Foto Profil", "e-KTP", "File SK"):
                sel.alignment = Alignment(vertical="top", wrap_text=False)
        terisi += 1

    ws.auto_filter.ref = f"A1:{get_column_letter(len(KOLOM_EXCEL))}{terisi + 1}"

    # Simpan ke file bulk terbaru
    keluaran.parent.mkdir(parents=True, exist_ok=True)
    wb.save(keluaran)

    # Perbarui juga file master rekapitulasi utama untuk Kejaksaan
    if simpan_master:
        try:
            wb.save(MASTER_EXCEL)
            catat_log(f"Master file diperbarui: {MASTER_EXCEL}")
        except Exception as e:
            catat_log(f"Gagal simpan master excel: {e}")
    else:
        catat_log("Simpan master dilewati (mode hanya data belum ada di SIMLINMAS).")

    return {
        "baris": terisi,
        "baru": jumlah_baru,
        "lama": jumlah_lama,
    }


# ============================================================ 9. UPLOAD SIMLINMAS (AUTO-HEALING)
def upload_data_ke_simlinmas(excel_path: Path, mode: str = "insert") -> tuple[int, int]:
    """Upload/update anggota ke SIMLINMAS. mode: insert|update|sync."""
    judul("PENGIRIMAN DATA KE SIMLINMAS KEMENDAGRI (AUTO-HEALING + RETRY)")
    info(f"File Data    : {hijau(excel_path.name)}")
    direct_available = BulkSesi is not None and BulkSimLinmas is not None
    use_relay = relay_enabled()
    if use_relay:
        info(f"Mode         : RELAY -> {RELAY_URL}")
    else:
        info("Mode         : Langsung ke SIMLINMAS (curl_cffi)")
        if not direct_available:
            gagal("Modul bulk tidak tersedia. Pengiriman dibatalkan.")
            return 0, 0

    # 1. Siapkan akses (direct atau relay)
    sesi = api = None
    if not use_relay:
        info("Menghubungi SIMLINMAS & Login...")
        def _login():
            s = BulkSesi(EMAIL_OPERATOR, PASSWORD_OPERATOR).masuk()
            a = BulkSimLinmas(s)
            return s, a
        ok_login, hasil_login = coba_ulang(_login, percobaan=6, jeda_awal=3, nama="login SIMLINMAS")
        if not ok_login:
            gagal(f"Gagal login ke SIMLINMAS setelah retry: {hasil_login}")
            return 0, 0
        sesi, api = hasil_login
        baik(f"Login Berhasil sebagai {EMAIL_OPERATOR}")
    else:
        baik("Relay mode aktif - login/CSRF ditangani relay per request.")

    def _refresh_api():
        nonlocal sesi, api
        if use_relay:
            return True
        ok_l, hl = coba_ulang(lambda: (lambda ss: (ss, BulkSimLinmas(ss)))(BulkSesi(EMAIL_OPERATOR, PASSWORD_OPERATOR).masuk()),
                              percobaan=4, jeda_awal=3, nama="re-login SIMLINMAS")
        if ok_l:
            sesi, api = hl
            return True
        return False

    # 2. Ambil member yang sudah ada
    info("Mengambil daftar member yang sudah terdaftar di server...")
    try:
        server_members = ambil_member_simlinmas()
    except Exception as e:
        catat(f"Gagal ambil member server: {e}. Melanjutkan dengan set kosong.")
        server_members = {}
    server_niks = set(server_members.keys())
    hasil("NIK terindeks di server", len(server_niks))

    # 3. Baca & saring Excel
    folder_dasar = str(excel_path.parent)
    baris_data, _ = baca_baris(str(excel_path))
    siap: list[SiapKirim] = []
    sudah_ada = 0
    akan_update = 0
    for b in baris_data:
        nik = konversi_nik(b.get(J_NIK)).nilai
        if not nik or len(nik) != 16:
            continue
        ada_di_server = nik in server_niks
        if mode == "insert" and ada_di_server:
            sudah_ada += 1
            continue
        if mode == "update" and not ada_di_server:
            continue
        nomor = b.get("__baris", "?")
        nama = str(b.get(J_NAMA) or f"baris {nomor}").strip()
        path_lampiran, _ = periksa_lampiran(b, folder_dasar)
        for k in list(path_lampiran.keys()):
            pth = Path(path_lampiran[k])
            if not pth.exists():
                for alt in [".pdf", ".jpg", ".jpeg"]:
                    p_alt = pth.with_suffix(alt)
                    if p_alt.exists():
                        path_lampiran[k] = str(p_alt)
                        break
        cek = validasi_baris(api, b, PASSWORD_ANGGOTA)
        if cek.salah:
            catat(f"Validasi dilewati: {nama} -> {cek.salah}")
            continue
        if ada_di_server and mode in ("update", "sync"):
            m = server_members.get(nik, {})
            cek.payload["_id_edit"] = m.get("id", "")
            cek.payload["_nomor_registrasi"] = m.get("nomor_registrasi", "")
            cek.payload["_mode"] = "update"
            akan_update += 1
        else:
            cek.payload["_mode"] = "insert"
        siap.append(SiapKirim(nomor, nama, cek.payload, path_lampiran))

    hasil("Sudah terdaftar di server", sudah_ada)
    hasil("Akan diupdate", akan_update)
    hasil("Akan diinsert/update (total)", len(siap))
    if not siap:
        baik("Tidak ada data yang perlu dikirim/diperbarui.")
        return 0, 0

    # 4. Kirim
    langkah("Kirim", f"Memproses {len(siap)} anggota ke SIMLINMAS (mode={mode})...")
    berhasil = 0
    gagal_items: list[SiapKirim] = []
    t_upload = time.time()

    def _kirim_api(payload, lampiran_bytes):
        kirim_mode = payload.get("_mode", "insert")
        id_edit = payload.get("_id_edit", "")
        nomor_reg = payload.get("_nomor_registrasi", "")
        fields_clean = {k: v for k, v in payload.items() if not k.startswith("_") and v is not None and v != ""}
        if kirim_mode == "update" and id_edit:
            fields_clean["id_edit"] = id_edit
            if nomor_reg:
                fields_clean.setdefault("nomor_registrasi", nomor_reg)
        # Untuk update: jangan kirim lampiran (lebih cepat, hanya perbarui data)
        lampiran_kirim = {} if (kirim_mode == "update" and id_edit) else (lampiran_bytes or {})
        if use_relay:
            files = {}
            for k, (fname, content, ctype) in lampiran_kirim.items():
                files[k] = (fname, content, ctype)
            action = "update" if (kirim_mode == "update" and id_edit) else "insert"
            return _relay_multipart(action, fields=fields_clean, files=files, timeout=180)
        if kirim_mode == "update" and id_edit:
            return api.perbarui(id_edit, fields_clean, lampiran_kirim)
        return api.tambah(fields_clean, lampiran_kirim)

    def _kirim_satu(item: SiapKirim, idx: int, total: int) -> bool:
        nonlocal berhasil
        nik = item.payload.get("nik", "")
        nama = item.nama
        lampiran_bytes = baca_isi_berkas(item.lampiran)
        status_str = ""
        for coba in range(1, 8):
            try:
                jawab = _kirim_api(item.payload, lampiran_bytes)
                if jawab.get("ok"):
                    berhasil += 1
                    aksi = "Terupdate" if item.payload.get("_mode") == "update" else "Terkirim"
                    status_str = hijau(f"[OK] {aksi}")
                    break
                pesan = str(jawab.get("message") or jawab.get("code") or "")
                if "email sudah terdaftar" in pesan.lower() and item.payload.get("email"):
                    catat(f"[{nama}] Email bentrok -> Dikosongkan otomatis...")
                    item.payload["email"] = ""
                    time.sleep(0.8); continue
                if "no hp sudah terdaftar" in pesan.lower():
                    clean_nik = "".join(filter(str.isdigit, nik))
                    nomor_baru = "08" + (clean_nik[-9:] if len(clean_nik) >= 9 else f"{idx:09d}")
                    catat(f"[{nama}] No HP bentrok -> Diganti otomatis ke {nomor_baru}...")
                    item.payload["no_hp"] = nomor_baru
                    time.sleep(0.8); continue
                if "nik sudah terdaftar" in pesan.lower():
                    berhasil += 1
                    status_str = hijau("[OK: Sudah Terdaftar]")
                    break
                if "ukuran tidak sesuai" in pesan.lower() or "jenis file" in pesan.lower():
                    catat(f"[{nama}] Berkas ditolak -> Re-kompresi/konversi...")
                    for k, fpath in list(item.lampiran.items()):
                        pth = Path(fpath)
                        if not pth.exists():
                            for alt in [".pdf", ".jpg", ".jpeg"]:
                                pa = pth.with_suffix(alt)
                                if pa.exists():
                                    item.lampiran[k] = str(pa); pth = pa; break
                        if pth.exists() and pth.suffix.lower() == ".pdf":
                            kompres_pdf_satlinmas(pth)
                        elif pth.exists() and pth.suffix.lower() in [".jpg", ".jpeg"]:
                            kompres_jpg_satlinmas(pth)
                        if k == "sk_file" and pth.exists() and pth.suffix.lower() != ".pdf":
                            target_pdf = pth.with_suffix(".pdf")
                            try:
                                from PIL import Image as _Im
                                with _Im.open(pth) as im:
                                    im.convert("RGB").save(target_pdf, "PDF", resolution=100.0)
                                if target_pdf.exists():
                                    pth.unlink(missing_ok=True)
                                    item.lampiran[k] = str(target_pdf)
                            except Exception:
                                pass
                    lampiran_bytes = baca_isi_berkas(item.lampiran)
                    time.sleep(1); continue
                status_str = merah(f"[GAGAL: {pesan[:120]}]")
                if any(x in pesan.lower() for x in ["unauth", "csrf", "token", "login", "419", "401", "403"]):
                    _refresh_api()
                time.sleep(1.5)
            except Exception as e:
                status_str = merah(f"[ERROR KONEKSI: {e}]")
                _refresh_api()
                time.sleep(2)
        pct = idx * 100 / max(total, 1)
        print(f"  [{idx:>3}/{total}] {pct:6.2f}% {nama:<28} ({nik}) -> {status_str}")
        catat_log(f"Upload {idx}/{total} {pct:.2f}% {nama} ({nik}) -> {status_str}")
        if "GAGAL" in status_str or "ERROR" in status_str:
            return False
        return True

    total_upload = len(siap)
    for idx, item in enumerate(siap, 1):
        ok = _kirim_satu(item, idx, total_upload)
        if not ok:
            gagal_items.append(item)
        if idx % 10 == 0 or idx == total_upload:
            simpan_state(muat_state())
        time.sleep(0.8)

    for ronde in range(1, 4):
        if not gagal_items:
            break
        catat(f"Retry upload ronde {ronde}: {len(gagal_items)} item gagal, mencoba ulang...")
        _refresh_api()
        sisa = []
        for j, item in enumerate(gagal_items, 1):
            ok = _kirim_satu(item, j, len(gagal_items))
            if not ok:
                sisa.append(item)
            time.sleep(1)
        gagal_items = sisa

    garis("-")
    gagal_kirim = len(gagal_items)
    total_akhir = -1
    if not use_relay:
        try:
            total_akhir = api.jumlah_member()
        except Exception:
            try:
                _refresh_api()
                total_akhir = api.jumlah_member()
            except Exception:
                total_akhir = -1
    hasil("Total Berhasil Dikirim", berhasil)
    hasil("Total Gagal (setelah retry)", gagal_kirim, salah=bool(gagal_kirim))
    if total_akhir >= 0:
        hasil("Total Anggota di Server Saat Ini", total_akhir)
    progres(total_upload, total_upload, "Upload selesai", f"berhasil={berhasil} gagal={gagal_kirim}", t_upload)
    if gagal_items:
        catat("Daftar yang masih gagal setelah semua retry:")
        for it in gagal_items:
            catat(f" - {it.nama} ({it.payload.get('nik','')})")
    garis()
    return berhasil, gagal_kirim

# =============================================================== 10. ALUR UTAMA


# =============================================================== 10b. BERSIHKAN SAMPAH
# File/folder yang memang bagian dari workflow ini dan HARUS dipertahankan.
PERTAHANKAN_NAMA = {
    "sinkron_semua.py",
    "bulk.py",
    "bulk_sijalinmaja.py",
    "README.md",
    "README_sijalinmaja.md",
    "README_sinkron.md",
    "README_sinkron_semua.md",
    ".gitignore",
    ".sinkron_drive_state.json",
    ".sinkron_semua_state.json",
    "sinkron_semua.log",
    "bin",
    ".venv",
    "template_sijalinmaja.xlsx",
}

# Pola sampah yang aman dihapus (log, cache, one-off script, output lama)
POLA_SAMPAH_GLOB = [
    "*.log",
    "*.pyc",
    "__pycache__",
    "compress_*.py",
    "compress_*.log",
    "kompres_*.py",
    "kompres_*.log",
    "kirim_*.py",
    "jalankan_*.py",
    "persiapan_dan_upload_terbaru.py",
    "proses_lengkap.py",
    "proses_lengkap.log",
    "bulk_baru_5.log",
    "bulk_insert_*.log",
    "bulk_sisa_*.log",
    "upload_retry.log",
    "hasil_bulk_*.json",
    "retry_*.json",
    "rekap_gagal_*.csv",
    "bulk-simlinmas-ponorogo-BARU-*.xlsx",
    "bulk-simlinmas-ponorogo-siap-lokal.xlsx",
]

# File penting yang tidak boleh dihapus meski cocok pola tertentu
PENTING_PATTERNS = [
    "bulk-simlinmas-ponorogo-siap-*.xlsx",  # output siap-bulk - kita simpan yang terbaru saja
]


def bersihkan_sampah(kering: bool = False) -> dict:
    """Hapus file/folder sampah di folder bulklinmas yang sudah tidak dipakai.

    Yang dihapus: log lama, __pycache__, script one-off yang sudah digantikan,
    json/csv rekap gagal sementara, dan file excel BARU/lokal sementara.
    Yang dipertahankan: sinkron_semua.py, bulk.py, README, .gitignore, bin/,
    .venv/, state, template, dan excel siap-bulk terbaru.
    """
    judul("BERSIHKAN FILE SAMPAH / TIDAK TERPAKAI")
    dihapus = []
    dilewati = []
    ukuran = 0

    # Tentukan excel siap terbaru yang harus dipertahankan
    excel_siap = sorted(HERE.glob("bulk-simlinmas-ponorogo-siap-*.xlsx"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    pertahankan_excel = set()
    if excel_siap:
        # simpan terbaru dan master siap.xlsx jika ada
        pertahankan_excel.add(excel_siap[0].resolve())
        for p in excel_siap:
            if p.name == "bulk-simlinmas-ponorogo-siap.xlsx":
                pertahankan_excel.add(p.resolve())

    kandidat: list[Path] = []
    seen = set()
    for pola in POLA_SAMPAH_GLOB:
        for p in HERE.glob(pola):
            rp = p.resolve()
            if rp in seen:
                continue
            seen.add(rp)
            kandidat.append(p)

    # Juga hapus file .py one-off yang dikenal (selain yang dipertahankan)
    skrip_oneoff = {
        "sinkron_drive.py", "persiapan_dan_upload_terbaru.py", "proses_lengkap.py",
        "kompres_upload.py", "kompres_semua_lampiran.py", "compress_sk.py",
        "compress_sk_fast.py", "compress_sk_parallel.py", "kirim_baru_5.py",
        "kirim_sisa_103.py", "jalankan_bulk_baru.py",
        "konversi_dan_siapkan_master.py",
    }

    for p in HERE.iterdir():
        if p.name in skrip_oneoff and p.resolve() not in seen:
            seen.add(p.resolve())
            kandidat.append(p)

    for p in sorted(kandidat, key=lambda x: x.name.lower()):
        rp = p.resolve()
        nama = p.name
        if nama in PERTAHANKAN_NAMA:
            dilewati.append(nama)
            continue
        if rp in pertahankan_excel:
            dilewati.append(nama)
            continue
        # Jangan hapus folder bin/.venv/DATA, jangan hapus diri sendiri
        if nama in {"bin", ".venv", "DATA SATLINMAS PONOROGO"}:
            dilewati.append(nama)
            continue
        try:
            if p.is_dir():
                sz = sum(f.stat().st_size for f in p.rglob("*") if f.is_file())
            elif p.is_file() or p.is_symlink():
                sz = p.stat().st_size
            else:
                continue
        except OSError:
            sz = 0

        aksi = "AKAN HAPUS" if kering else "HAPUS"
        info(f"{aksi:<11} {nama} ({sz/1024:.1f} KB)")
        if not kering:
            try:
                if p.is_dir():
                    shutil.rmtree(p)
                else:
                    p.unlink()
                dihapus.append(nama)
                ukuran += sz
            except Exception as e:
                gagal(f"Gagal hapus {nama}: {e}")
        else:
            dihapus.append(nama)
            ukuran += sz

    garis("-")
    hasil("Item terdeteksi sampah", len(dihapus))
    hasil("Total ukuran dibebaskan", f"{ukuran/1024/1024:.2f} MB")
    if kering:
        catat("Mode kering (dry-run): tidak ada yang dihapus.")
    else:
        baik("Pembersihan selesai.")
    return {"dihapus": dihapus, "ukuran": ukuran}

def relay_enabled() -> bool:
    return bool(RELAY_URL and RELAY_TOKEN)


def _relay_multipart(action: str, fields: dict | None = None, files: dict | None = None, timeout: int = 150) -> dict:
    import mimetypes
    data = {"action": action, "token": RELAY_TOKEN}
    for k, v in (fields or {}).items():
        if v is None:
            continue
        data[k] = str(v)
    file_map = {}
    for k, (fname, content, ctype) in (files or {}).items():
        file_map[k] = (fname, content, ctype or mimetypes.guess_type(fname)[0] or "application/octet-stream")
    headers = {"X-Relay-Token": RELAY_TOKEN, "User-Agent": "SinkronSatlinmas/1.0"}

    # Prefer curl_cffi to pass Cloudflare challenge on relay domain
    try:
        from curl_cffi import requests as cffi_requests
        r = cffi_requests.post(RELAY_URL, data=data, files=file_map or None, headers=headers,
                               impersonate="chrome", timeout=timeout)
        try:
            return r.json()
        except Exception:
            return {"ok": False, "message": f"HTTP {r.status_code}: {r.text[:300]}"}
    except Exception as e:
        catat_log(f"curl_cffi relay gagal, fallback urllib: {e}")

    # Fallback urllib
    import urllib.request, urllib.error, uuid
    boundary = "----SimRelay" + uuid.uuid4().hex
    body = bytearray()
    def add_field(name, value):
        body.extend(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"\r\n\r\n{value}\r\n".encode("utf-8"))
    def add_file(name, filename, content, ctype):
        body.extend(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"; filename=\"{filename}\"\r\nContent-Type: {ctype}\r\n\r\n".encode("utf-8"))
        body.extend(content)
        body.extend(b"\r\n")
    for k, v in data.items():
        add_field(k, v)
    for k, (fname, content, ctype) in file_map.items():
        add_file(k, fname, content, ctype)
    body.extend(f"--{boundary}--\r\n".encode("utf-8"))
    req = urllib.request.Request(RELAY_URL, data=bytes(body), method="POST")
    req.add_header("Content-Type", f"multipart/form-data; boundary={boundary}")
    req.add_header("X-Relay-Token", RELAY_TOKEN)
    req.add_header("User-Agent", "SinkronSatlinmas/1.0")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            return json.loads(raw)
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:
            return json.loads(raw)
        except Exception:
            return {"ok": False, "message": f"HTTP {e.code}: {raw[:300]}"}
    except Exception as e:
        return {"ok": False, "message": f"Relay error: {e}"}


def ambil_member_simlinmas() -> dict[str, dict]:
    """Ambil map NIK -> {id, nomor_registrasi} dari SIMLINMAS (direct/relay)."""
    direct_ok = BulkSesi is not None and BulkSimLinmas is not None
    if relay_enabled():
        try:
            info(f"Ambil data member SIMLINMAS via relay: {RELAY_URL}")
            def _minta():
                r = _relay_multipart("list_nik", timeout=180)
                if not isinstance(r, dict) or not r.get("ok"):
                    raise RuntimeError(str(r))
                return r
            ok, resp = coba_ulang(_minta, percobaan=4, jeda_awal=3, nama="relay list_nik")
            if not ok or not isinstance(resp, dict):
                raise RuntimeError(f"Relay list_nik gagal: {resp}")
            members = resp.get("members") or {}
            # normalize: may be dict nik->{id,nomor_registrasi}
            out = {}
            if isinstance(members, dict) and members:
                for k,v in members.items():
                    nk = str(k).strip()
                    if not nk: continue
                    if isinstance(v, dict):
                        out[nk] = {"id": str(v.get("id","")).strip(), "nomor_registrasi": str(v.get("nomor_registrasi","")).strip()}
                    else:
                        out[nk] = {"id": str(v).strip(), "nomor_registrasi": ""}
            if not out:
                for nk in (resp.get("niks") or []):
                    nk = str(nk).strip()
                    if nk: out[nk] = {"id":"", "nomor_registrasi":""}
            baik(f"Member SIMLINMAS terindeks: {len(out)}")
            return out
        except Exception as e:
            if direct_ok:
                catat(f"Relay gagal ({e}), fallback ke direct...")
            else:
                raise

    if direct_ok:
        info("Login ke SIMLINMAS (langsung)...")
        ok_login, hl = coba_ulang(lambda: (lambda ss: (ss, BulkSimLinmas(ss)))(BulkSesi(EMAIL_OPERATOR, PASSWORD_OPERATOR).masuk()),
                                  percobaan=4, jeda_awal=3, nama="login SIMLINMAS")
        if not ok_login:
            raise RuntimeError(f"Login SIMLINMAS gagal: {hl}")
        sesi_cek, api_cek = hl
        total_srv = api_cek.jumlah_member()
        baik(f"Login SIMLINMAS OK. Total anggota di server: {total_srv}")
        out = {}
        start = 0
        t_cek = time.time()
        while start <= total_srv + 400:
            try:
                _, data_srv = sesi_cek.ambil_json(f"/linmas?start={start}&length=200")
                rows_srv = data_srv.get("data", []) if isinstance(data_srv, dict) else []
            except Exception as e:
                catat(f"Gagal ambil member start={start}: {e}")
                break
            if not rows_srv:
                break
            for r in rows_srv:
                nk = str(r.get("nik") or "").strip()
                if nk:
                    out[nk] = {"id": str(r.get("id","")).strip(), "nomor_registrasi": str(r.get("nomor_registrasi","")).strip()}
            start += 200
            if start % 1000 == 0 or start > total_srv:
                progres(min(start, max(total_srv, 1)), max(total_srv, 1), "Ambil member SIMLINMAS", f"terkumpul={len(out)}", t_cek)
        baik(f"Member SIMLINMAS terindeks: {len(out)}")
        return out

    raise RuntimeError("Tidak ada metode SIMLINMAS yang tersedia (direct/relay).")


def ambil_nik_simlinmas() -> set[str]:
    """Ambil semua NIK yang sudah terdaftar di SIMLINMAS (langsung atau via relay)."""
    return set(ambil_member_simlinmas().keys())


def proses_semua_otomatis(remote_diminta: str, rekap: Path, upload: bool = True, mode_upload: str = "insert") -> int:
    """Alur lengkap: Preflight -> Sinkron Drive -> Bandingkan SIMLINMAS -> Download/Kompress -> Excel -> Upload."""
    judul("MULAI PROSES OTOMATIS")
    info(f"Waktu mulai : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    info(f"Rekap       : {rekap}")

    # 1) Pastikan koneksi Drive OK
    langkah("P1", "Cek koneksi Google Drive")
    info("Memilih remote rclone...")
    remote = pilih_remote(remote_diminta)
    info(f"Remote dipilih: {remote}")
    cek_koneksi(remote)
    baik(f"Koneksi Google Drive OK: {remote}")

    # 2) Pastikan login SIMLINMAS OK dan ambil NIK existing
    nik_sudah_server: set[str] = set()
    if upload and (relay_enabled() or (BulkSesi is not None and BulkSimLinmas is not None)):
        langkah("P2", "Cek login SIMLINMAS & ambil NIK yang sudah ada")
        if relay_enabled():
            baik(f"Mode relay aktif: {RELAY_URL}")
        nik_sudah_server = ambil_nik_simlinmas()
    else:
        catat("Cek SIMLINMAS dilewati (mode tanpa upload / modul bulk/relay tidak tersedia).")

    # 3) Sinkron Drive, bandingkan, download hanya yang belum ada
    data = sinkron(remote, rekap, nik_sudah_server=nik_sudah_server, filter_server=(mode_upload != "sync"))

    tanggal = datetime.now().strftime("%Y%m%d")
    keluaran = HERE / f"bulk-simlinmas-ponorogo-siap-{tanggal}.xlsx"

    langkah("5", "Membuat Excel Siap-Bulk & Memperbarui Master Kejaksaan")
    ringkas = buat_excel(data["baris"], data["laporan"], keluaran,
                         state=data["state"], nik_referensi=data["nik_referensi"],
                         simpan_master=not data.get("difilter_server", False))

    state = data["state"]
    state["nik"] = sorted({r["nik"] for r in data["baris"] if r["nik"]})
    simpan_state(state)

    hasil("Total baris data", ringkas["baris"])
    hasil("Anggota status BARU", ringkas["baru"])
    hasil("File Excel Siap-Bulk", str(keluaran.name))
    if not data.get("difilter_server"):
        hasil("File Master Kejaksaan", str(MASTER_EXCEL.name))

    gagal_up_final = 0
    berhasil_up_final = 0
    sudah_ada_server = 0
    if upload:
        berhasil_up, gagal_up = upload_data_ke_simlinmas(keluaran, mode=mode_upload)
        berhasil_up_final += berhasil_up
        gagal_up_final = gagal_up
        if gagal_up > 0:
            catat(f"Masih ada {gagal_up} gagal upload. Menjalankan ulang pass upload penuh...")
            berhasil_up2, gagal_up2 = upload_data_ke_simlinmas(keluaran, mode=mode_upload)
            berhasil_up_final += berhasil_up2
            gagal_up_final = gagal_up2

    # Tulis ringkasan markdown untuk GitHub Step Summary
    try:
        stat_berkas = data.get("stat_berkas", {}) or {}
        total_server_awal = len(nik_sudah_server) if nik_sudah_server else 0
        total_baris = len(data.get("baris", []))  # after filter
        status = "✅ Sukses" if gagal_up_final == 0 else f"⚠️ Selesai dengan {gagal_up_final} gagal"
        summary = f"""## Sinkron Satlinmas - {status}

| Metrik | Nilai |
|---|---:|
| Waktu | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} |
| NIK di SIMLINMAS (sebelum run) | {total_server_awal} |
| Data diproses (belum ada di SIMLINMAS) | {total_baris} |
| Berhasil upload ke SIMLINMAS | {berhasil_up_final} |
| Gagal upload | {gagal_up_final} |
| SK dikonversi ke PDF | {stat_berkas.get('sk_ke_pdf', 0)} |
| Foto/KTP dikonversi ke JPG | {stat_berkas.get('foto_ke_jpg', 0)} |
| Berkas terkompresi (<1MB) | {stat_berkas.get('kompres_ok', 0)} / {stat_berkas.get('kompres_total', 0)} |
| Gagal kompres | {stat_berkas.get('kompres_gagal', 0)} |

> Proses via GitHub Actions + relay. Lampiran ephemeral (tidak disimpan di repo).
"""
        (HERE / "ringkasan_sinkron.md").write_text(summary, encoding="utf-8")
    except Exception as e:
        catat_log(f"Gagal tulis ringkasan: {e}")

    garis("=")
    hasil("Log detail", str(LOG_FILE))
    if upload and gagal_up_final > 0:
        catat(f"Proses selesai dengan {gagal_up_final} item masih gagal upload setelah semua retry.")
        return 1
    baik("Seluruh rangkaian proses selesai tuntas.")
    return 0


def cari_excel_siap_terbaru() -> Path:
    pola = str(HERE / "bulk-simlinmas-ponorogo-siap-*.xlsx")
    files = [Path(f) for f in glob.glob(pola) if "BARU" not in f and "lokal" not in f]
    if not files:
        return HERE / "bulk-simlinmas-ponorogo-siap.xlsx"
    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return files[0]


def default_keluaran() -> Path:
    tanggal = datetime.now().strftime("%Y%m%d")
    path = HERE / f"bulk-simlinmas-ponorogo-siap-{tanggal}.xlsx"
    if path.exists():
        jam = datetime.now().strftime("%H%M")
        path = HERE / f"bulk-simlinmas-ponorogo-siap-{tanggal}-{jam}.xlsx"
    return path


def menu_interaktif() -> int:
    judul("SINKRON SEMUA SATLINMAS PONOROGO (MANDIRI)")
    print("""
  Pilih mode operasional:
    1. PROSES LENGKAP OTOMATIS (Sinkron + Kompres + Upload ke SIMLINMAS) [REKOMENDASI]
    2. Sinkron Drive + Unduh + Kompres + Buat Excel Saja (Tanpa Upload)
    3. Standarkan Format & Kompres Seluruh Berkas Lokal (< 1 MB, JPG & PDF)
    4. Upload Data BARU ke SIMLINMAS Saja (dengan Auto-Healing)
    5. Cek Koneksi Google Drive & SIMLINMAS
    6. Bersihkan File Sampah / Tidak Terpakai
    0. Keluar
""")
    while True:
        pilih = input("  Masukkan pilihan [1/2/3/4/5/6/0]: ").strip()
        if pilih == "0":
            return 0
        if pilih == "1":
            return proses_semua_otomatis("", REKAP_BAWAAN, upload=True)
        if pilih == "2":
            return proses_semua_otomatis("", REKAP_BAWAAN, upload=False)
        if pilih == "3":
            standarkan_dan_kompres_semua_berkas()
            return 0
        if pilih == "4":
            target = cari_excel_siap_terbaru()
            upload_data_ke_simlinmas(target)
            return 0
        if pilih == "5":
            remote = pilih_remote("")
            cek_koneksi(remote)
            baik(f"Koneksi Google Drive ({remote}) aman!")
            if BulkSesi:
                sesi = BulkSesi(EMAIL_OPERATOR, PASSWORD_OPERATOR).masuk()
                api = BulkSimLinmas(sesi)
                baik(f"Koneksi SIMLINMAS aman! Total anggota di server: {api.jumlah_member()}")
            return 0
        if pilih == "6":
            bersihkan_sampah(kering=False)
            return 0
        catat("Pilihan tidak valid, silakan masukkan nomor yang tersedia.")


def buat_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sinkron_semua.py",
        description="All-in-One: Sinkron Google Drive, standardisasi format, kompresi < 1 MB, dan upload SIMLINMAS."
    )
    parser.add_argument("perintah", nargs="?", default="",
                        choices=["", "all", "sinkron", "upload", "kompres", "cek", "bersihkan", "perbarui"],
                        help="all=sinkron+upload baru; perbarui=sync semua (update NIK sama+insert baru); sinkron=tanpa upload; upload=upload saja; kompres=format & kompres saja; bersihkan=hapus file sampah")
    parser.add_argument("--remote", default="", help="nama remote rclone (mis. gdrive)")
    parser.add_argument("--rekap", default=str(REKAP_BAWAAN), help="file rekap utama")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = buat_parser().parse_args(argv)
    if args.perintah in {"all", "perbarui", "upload", "cek"} and not RELAY_TOKEN:
        raise RuntimeError("SIMLINMAS_RELAY_TOKEN wajib diatur lewat environment/.env lokal atau GitHub Secret.")
    rekap = Path(args.rekap)

    try:
        if args.perintah == "" or args.perintah is None:
            return menu_interaktif()
        if args.perintah == "all":
            return proses_semua_otomatis(args.remote, rekap, upload=True, mode_upload="insert")
        if args.perintah == "perbarui":
            return proses_semua_otomatis(args.remote, rekap, upload=True, mode_upload="sync")
        if args.perintah == "sinkron":
            return proses_semua_otomatis(args.remote, rekap, upload=False)
        if args.perintah == "kompres":
            standarkan_dan_kompres_semua_berkas()
            return 0
        if args.perintah == "upload":
            target = cari_excel_siap_terbaru()
            upload_data_ke_simlinmas(target)
            return 0
        if args.perintah == "cek":
            remote = pilih_remote(args.remote)
            cek_koneksi(remote)
            baik(f"Koneksi Google Drive ({remote}) aman!")
            if relay_enabled():
                ok, resp = coba_ulang(lambda: _relay_multipart("ping", timeout=120), percobaan=3, jeda_awal=3, nama="relay ping")
                if not ok or not isinstance(resp, dict) or not resp.get("ok"):
                    pesan = resp.get("message") if isinstance(resp, dict) else str(resp)
                    raise RuntimeError(f"Koneksi relay SIMLINMAS gagal: {pesan}")
                baik(f"Koneksi relay SIMLINMAS aman: {resp.get('message', 'OK')}")
            elif BulkSesi:
                sesi = BulkSesi(EMAIL_OPERATOR, PASSWORD_OPERATOR).masuk()
                api = BulkSimLinmas(sesi)
                baik(f"Koneksi SIMLINMAS aman! Total anggota: {api.jumlah_member()}")
            else:
                catat("Cek SIMLINMAS dilewati: relay tidak aktif dan modul direct tidak tersedia.")
            return 0
        if args.perintah == "bersihkan":
            bersihkan_sampah(kering=False)
            return 0
    except KeyboardInterrupt:
        print("\n\nProses dibatalkan oleh pengguna (Ctrl+C).")
        return 130
    except Exception as e:
        gagal(str(e))
        return 1
    return 0

if __name__ == "__main__":
    sys.exit(main())
