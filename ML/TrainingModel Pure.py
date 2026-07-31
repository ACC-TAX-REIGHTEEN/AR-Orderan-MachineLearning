import os
import re
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.neighbors import NearestNeighbors

KODE_BARANG_KOTOR = [
    r"\bIRC\b", r"\bZN\b", r"\bTT\b", r"\bBD\b", r"\bTL\b",
    r"\bMC\b", r"\bTC\b", r"\bFL\b", r"\bSL\b", r"\bHD\b",
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


def proses_pencocokan_ml(
    file_input="TheTrainningData.xlsx",
    file_output="Hasil_Latihan_ML.xlsx",
    batas_minimal_skor=0.65
):
    print(f"--> Membaca file input: '{file_input}' ...")

    if not os.path.exists(file_input):
        print(f"--> Error file '{file_input}' tidak ditemukan.")
        return

    df = pd.read_excel(file_input)
    df.columns = df.columns.str.strip()

    col_input = "Nama Customer dan Kota"
    col_f = "Ekstrak_Customer_Detail (F)"
    col_f2 = "Ekstrak_Customer_Detail (F2)"
    col_f3 = "Ekstrak_Customer_Detail (F3)"

    data_input_clean = []
    data_target_master = []

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
            data_input_clean.append(input_clean)
            data_target_master.append(master_clean)

    if not data_input_clean:
        print("--> Error tidak ada data historis valid untuk melatih model ML.")
        return

    print(f"--> Melatih Model Machine Learning dengan {len(data_input_clean)} sampel data...")

    vectorizer = TfidfVectorizer(analyzer='char_wb', ngram_range=(2, 4))
    X_train = vectorizer.fit_transform(data_input_clean)

    model_nn = NearestNeighbors(n_neighbors=1, metric='cosine')
    model_nn.fit(X_train)

    print("--> Pelatihan model selesai! Memproses prediksi rekomendasi...")

    hasil_rekomendasi = []
    hasil_skor = []
    hasil_status = []
    hasil_sumber = []

    for idx, row in df.iterrows():
        input_raw = row[col_input]
        input_clean = bersihkan_teks_input(input_raw)

        if not input_clean:
            hasil_rekomendasi.append("TIDAK DITEMUKAN")
            hasil_skor.append(0.0)
            hasil_status.append("GAGAL (Input Kosong)")
            hasil_sumber.append("-")
            continue

        kandidat_lokal = []
        for col_name in [col_f2, col_f, col_f3]:
            if col_name in df.columns and pd.notna(row[col_name]):
                v = str(row[col_name]).strip()
                if v and v.lower() not in ["nan", "none", "-"]:
                    kandidat_lokal.append((col_name, v))

        if kandidat_lokal:
            sumber_col, val_lokal = kandidat_lokal[0]
            hasil_rekomendasi.append(bersihkan_nama_rekomendasi(val_lokal))
            hasil_skor.append(100.0)
            hasil_status.append("SUCCESS (Match Lokal)")
            hasil_sumber.append(sumber_col)
        else:
            vec_input = vectorizer.transform([input_clean])
            distances, indices = model_nn.kneighbors(vec_input)

            jarak_cosine = distances[0][0]
            idx_terdekat = indices[0][0]

            kemiripan = 1.0 - jarak_cosine
            skor_persen = round(kemiripan * 100, 2)

            if kemiripan >= batas_minimal_skor:
                rekomendasi_ml = data_target_master[idx_terdekat]
                input_mirip = data_input_clean[idx_terdekat]

                hasil_rekomendasi.append(rekomendasi_ml)
                hasil_skor.append(skor_persen)
                hasil_status.append(f"SUCCESS (ML Prediction - Match: '{input_mirip}')")
                hasil_sumber.append("Model ML (TF-IDF + k-NN)")
            else:
                hasil_rekomendasi.append("TIDAK DITEMUKAN")
                hasil_skor.append(skor_persen)
                hasil_status.append("GAGAL (Skor ML di bawah batas minimal)")
                hasil_sumber.append("Model ML (TF-IDF + k-NN)")

    df["Hasil_Nama_Rekomendasi"] = hasil_rekomendasi
    df["Skor_Kemiripan_%"] = hasil_skor
    df["Status_Pencocokan"] = hasil_status
    df["Sumber_Pencocokan"] = hasil_sumber

    df.to_excel(file_output, index=False)
    print("--> Pemrosesan selesai!")
    print(f"--> File hasil disimpan ke: '{file_output}'")


if __name__ == "__main__":
    proses_pencocokan_ml()