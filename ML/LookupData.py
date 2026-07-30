import os
import re
import pandas as pd

FILE_UTAMA = "TheTrainningData.xlsx"
FOLDER_BASE = r"E:\ADM IRC AND ZN\2026"


def bersihkan_angka(val):
    if pd.isnull(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_clean = re.sub(r"[^\d]", "", str(val))
    return float(val_clean) if val_clean else 0.0


def pisahkan_nama_file(nama_file):
    nama_tanpa_ext = os.path.splitext(nama_file)[0]
    bagian = [b.strip() for b in nama_tanpa_ext.split(",") if b.strip()]
    sales, customer_detail, sr_code, tgl_nota = "", "", "", ""

    if len(bagian) >= 4:
        sales = bagian[0]
        customer_detail = bagian[1]
        sr_code = bagian[2]
        tgl_nota = bagian[3]
    elif len(bagian) == 3:
        sales = bagian[0]
        customer_detail = bagian[1]
        sr_code = bagian[2]
    elif len(bagian) == 2:
        customer_detail = bagian[0]
        sr_code = bagian[1]
    elif len(bagian) == 1:
        customer_detail = bagian[0]

    return sales, customer_detail, sr_code, tgl_nota


if __name__ == "__main__":
    if not os.path.exists(FILE_UTAMA):
        print(f"--> Error: File utama '{FILE_UTAMA}' tidak ditemukan!")
        exit()

    print(
        "--> Menganalisis dan membuat kamus dari seluruh file di folder referensi..."
    )

    kamus_nominal = {}
    jumlah_file = 0

    for root, dirs, files in os.walk(FOLDER_BASE):
        for file in files:
            if (file.endswith(".xlsx") or file.endswith(".xls")) and not file.startswith("~$"):
                f_path = os.path.join(root, file)
                jumlah_file += 1

                if jumlah_file % 20 == 0:
                    print(f"--> Sedang membaca file ke-{jumlah_file}...")

                try:
                    sales, customer, sr, tgl = pisahkan_nama_file(file)
                    df_ref = pd.read_excel(f_path, header=None)

                    for _, row_ref in df_ref.iterrows():
                        for cell in row_ref:
                            if pd.notnull(cell):
                                nom = bersihkan_angka(cell)
                                if nom > 0 and nom not in kamus_nominal:
                                    kamus_nominal[nom] = (sales, customer, sr, tgl)
                except Exception:
                    continue

    print(
        f"--> Selesai membaca {jumlah_file} file referensi. Total nominal unik tercatat: {len(kamus_nominal)}"
    )

    print(
        f"--> Membaca file utama '{FILE_UTAMA}' dan mencocokkan data..."
    )
    df_utama = pd.read_excel(FILE_UTAMA)

    list_sales = []
    list_customer = []
    list_sr = []
    list_tgl = []

    for index, row in df_utama.iterrows():
        val_nominal_d = row.iloc[3]
        nominal_target = bersihkan_angka(val_nominal_d)

        if nominal_target > 0 and nominal_target in kamus_nominal:
            sales, customer, sr, tgl = kamus_nominal[nominal_target]
        else:
            sales, customer, sr, tgl = "", "", "", ""

        list_sales.append(sales)
        list_customer.append(customer)
        list_sr.append(sr)
        list_tgl.append(tgl)

    print("--> Menyimpan hasil ekstraksi ke file Excel...")
    df_utama["Ekstrak_Sales (E)"] = list_sales
    df_utama["Ekstrak_Customer_Detail (F)"] = list_customer
    df_utama["Ekstrak_SR (G)"] = list_sr
    df_utama["Ekstrak_Tgl_Nota (H)"] = list_tgl

    df_utama.to_excel(FILE_UTAMA, index=False)
    print(
        f"--> SELESAI! Data berhasil diekstrak dan disimpan di '{FILE_UTAMA}'."
    )