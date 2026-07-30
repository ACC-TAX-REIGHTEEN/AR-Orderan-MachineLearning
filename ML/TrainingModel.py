import os
import re
import pandas as pd
from rapidfuzz import fuzz, process

KODE_BARANG_KOTOR = [
    r"\bIRC\b",
    r"\bZN\b",
    r"\bTT\b",
    r"\bBD\b",
    r"\bTL\b",
    r"\bMC\b",
    r"\bTC\b",
    r"\bFL\b",
    r"\bSL\b",
    r"\bHD\b",
]


def bersihkan_nama_rekomendasi(teks):
    if pd.isna(teks):
        return ""
    teks_clean = str(teks).strip()
    for kode in KODE_BARANG_KOTOR:
        teks_clean = re.sub(kode, "", teks_clean, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", teks_clean).strip()


def bersihkan_teks_input(teks):
    if pd.isna(teks):
        return ""
    teks_clean = str(teks).lower()
    teks_clean = re.sub(r"[^\w\s]", " ", teks_clean)
    return " ".join(teks_clean.split())


def proses_pencocokan_historis_cerdas(
    file_input="TheTrainningData.xlsx",
    file_output="Hasil_Latihan.xlsx",
):
    print(f"--> Membaca file input: '{file_input}' ...")

    if not os.path.exists(file_input):
        print(f"--> [ERROR] File '{file_input}' tidak ditemukan.")
        return

    df = pd.read_excel(file_input)
    df.columns = df.columns.str.strip()

    col_input = "Nama Customer dan Kota"
    col_f = "Ekstrak_Customer_Detail (F)"
    col_f2 = "Ekstrak_Customer_Detail (F2)"
    col_f3 = "Ekstrak_Customer_Detail (F3)"

    memori_input_ke_master = {}

    for _, row in df.iterrows():
        input_clean = bersihkan_teks_input(row[col_input])

        target_resmi = None
        for col_target in [col_f2, col_f, col_f3]:
            val = row[col_target]
            if pd.notna(val) and str(val).strip() not in ["", "-", "nan"]:
                target_resmi = str(val).strip()
                break

        if target_resmi and input_clean:
            master_clean = bersihkan_nama_rekomendasi(target_resmi)
            memori_input_ke_master[input_clean] = master_clean

    daftar_input_historis = list(memori_input_ke_master.keys())

    print(
        f"--> Berhasil mempelajari {len(daftar_input_historis)} variasi pola input historis yang valid!"
    )

    hasil_rekomendasi = []
    hasil_skor = []
    hasil_status = []
    hasil_sumber = []

    BATAS_MINIMAL_SKOR = 65.0

    for idx, row in df.iterrows():
        input_raw = row[col_input]
        input_clean = bersihkan_teks_input(input_raw)

        if not input_clean:
            hasil_rekomendasi.append("TIDAK DITEMUKAN")
            hasil_skor.append(0.0)
            hasil_status.append("FAILED (Input Kosong)")
            hasil_sumber.append("-")
            continue

        rekomendasi = "TIDAK DITEMUKAN"
        skor = 0.0
        status = "FAILED (Data Kosong / Belum Ada Master)"
        sumber = "-"

        kandidat_lokal = []
        for col_name in [col_f2, col_f, col_f3]:
            if col_name in df.columns and pd.notna(row[col_name]):
                v = str(row[col_name]).strip()
                if v and v.lower() not in ["nan", "none", "-"]:
                    kandidat_lokal.append((col_name, v))

        if kandidat_lokal:
            sumber_col, val_lokal = kandidat_lokal[0]
            rekomendasi = bersihkan_nama_rekomendasi(val_lokal)
            skor = 100.0
            status = "SUCCESS (Match Lokal)"
            sumber = sumber_col
        else:
            if input_clean in memori_input_ke_master:
                rekomendasi = memori_input_ke_master[input_clean]
                skor = 100.0
                status = "SUCCESS (Match Exact Memori Historis)"
                sumber = "Kamus Ingatan Historis"
            elif daftar_input_historis:
                match_paling_mirip = process.extract(
                    query=input_clean,
                    choices=daftar_input_historis,
                    scorer=fuzz.token_set_ratio,
                    limit=1,
                )

                if match_paling_mirip:
                    input_lama_mirip, skor_kemiripan, _ = match_paling_mirip[0]

                    if float(skor_kemiripan) >= BATAS_MINIMAL_SKOR:
                        skor = float(skor_kemiripan)
                        rekomendasi = memori_input_ke_master[input_lama_mirip]
                        status = f"SUCCESS (Match Pola Input: '{input_lama_mirip}')"
                        sumber = "Asosiasi Ingatan Input Historis"

        hasil_rekomendasi.append(rekomendasi)
        hasil_skor.append(round(skor, 2))
        hasil_status.append(status)
        hasil_sumber.append(sumber)

    df["Hasil_Nama_Rekomendasi"] = hasil_rekomendasi
    df["Skor_Kemiripan_%"] = hasil_skor
    df["Status_Pencocokan"] = hasil_status
    df["Sumber_Pencocokan"] = hasil_sumber

    df.to_excel(file_output, index=False)
    print("--> PEMROSESAN SELESAI!")
    print(f"--> File hasil disimpan ke: '{file_output}'")


if __name__ == "__main__":
    proses_pencocokan_historis_cerdas()