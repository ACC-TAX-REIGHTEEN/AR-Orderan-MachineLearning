import configparser
import os
import requests


def download_sheet(url, output_name):
    if not url or str(url).strip() == "":
        print(
            f"--> Peringatan: URL untuk {output_name} kosong di config.conf!"
            " Proses unduh dilewati."
        )
        return

    if "/d/" in url:
        base_id = url.split("/d/")[1].split("/")[0]
        download_url = (
            f"https://docs.google.com/spreadsheets/d/{base_id}/export?format=xlsx"
        )
    else:
        print(f"--> URL tidak valid untuk {output_name}: {url}")
        return

    try:
        print(f"--> Sedang mengunduh {output_name}...")
        response = requests.get(download_url)
        response.raise_for_status()

        with open(output_name, "wb") as f:
            f.write(response.content)
        print(f"--> Berhasil! File disimpan sebagai {output_name}")

    except Exception as e:
        print(f"--> Gagal mengunduh {output_name}: {e}")


def main():
    config = configparser.ConfigParser(allow_no_value=True)
    config_file = "config.conf"

    if not os.path.exists(config_file):
        print(f"--> File '{config_file}' tidak ditemukan!")
        return

    config.read(config_file)

    if config.has_section("MLDT"):
        ml_url = config.get("MLDT", "ml_url", fallback="").strip()
        output_file = "Hasil_Latihan.xlsx"

        if ml_url:
            download_sheet(ml_url, output_file)
        else:
            print(
                "--> Peringatan: 'ml_url' tidak ditemukan atau nilainya kosong"
                " di bawah [MLDT]!"
            )
    else:
        print("--> Section [MLDT] tidak ditemukan di config.conf")

if __name__ == "__main__":
    main()
