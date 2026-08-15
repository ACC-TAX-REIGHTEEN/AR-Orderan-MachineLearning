import os
import subprocess
import sys

def verifikasi_folder_dan_file(nama_folder, daftar_file):
    if not os.path.exists(nama_folder) or not os.path.isdir(nama_folder):
        print(f"--> Folder '{nama_folder}' tidak ditemukan!")
        return False

    file_lengkap = True
    for file_name in daftar_file:
        path_file = os.path.join(nama_folder, file_name)
        if not os.path.isfile(path_file):
            print(
                f"--> File '{file_name}' TIDAK ditemukan di dalam folder"
                f" '{nama_folder}'."
            )
            file_lengkap = False

    return file_lengkap

def jalankan_otomatisasi():
    print("--> Memulai proses automasi")

    folder_dapur = "Dapur"
    file_dapur_wajib = [
        "1_CopyData.py",
        "2_AdjDateFormat.py",
        "3_InjectDataToSS.py",
        "config.conf",
        "credentials.json",
    ]

    folder_ml = "ML"
    file_ml_wajib = [
        "Hasil_Latihan_temp.xlsx",
        "LookupData.py",
        "TrainingModel.py",
        "TheTrainningData.xlsx",
    ]

    print("--> Memeriksa struktur folder Dapur...")
    dapur_ok = verifikasi_folder_dan_file(folder_dapur, file_dapur_wajib)

    print("--> Memeriksa struktur folder ML...")
    ml_ok = verifikasi_folder_dan_file(folder_ml, file_ml_wajib)

    if not dapur_ok or not ml_ok:
        print("--> Persyaratan file/folder belum terpenuhi.")
        print("--> Harap lengkapi file yang kurang sebelum melanjutkan.")
        input("--> Tekan Enter untuk keluar...")
        return

    print("--> Semua folder dan file terverifikasi lengkap!")

    try:
        script_1 = "1_CopyData.py"
        print(f"--> Memulai eksekusi '{script_1}' di folder '{folder_dapur}'...")
        subprocess.run([sys.executable, script_1], cwd=folder_dapur, check=True)
        print(f"--> Eksekusi '{script_1}' berhasil dijalankan.")
        
        script_2 = "2_AdjDateFormat.py"
        print(f"--> Memulai eksekusi '{script_2}' di folder '{folder_dapur}'...")
        subprocess.run([sys.executable, script_2], cwd=folder_dapur, check=True)
        print(f"--> Eksekusi '{script_2}' berhasil dijalankan.")

        script_3 = "3_InjectDataToSS.py"
        print(
            f"--> Memulai eksekusi '{script_3}' di folder '{folder_dapur}'..."
        )
        print(
            "--> Catatan: Jika script ini berjalan terus dalam loop, tekan Ctrl+C"
            " untuk menghentikan."
        )

        subprocess.run([sys.executable, script_3], cwd=folder_dapur)

    except subprocess.CalledProcessError as err:
        print(f"--> Terjadi kesalahan saat menjalankan script: {err}")
    except KeyboardInterrupt:
        print("--> Eksekusi dihentikan oleh pengguna (Ctrl+C).")
    except Exception as e:
        print(f"--> Terjadi kesalahan sistem: {e}")

    print("--> Seluruh proses automasi selesai dijalankan.")
    input("--> Tekan Enter untuk keluar...")

if __name__ == "__main__":
    jalankan_otomatisasi()