import configparser
import requests
import os

def download_sheet(url, output_name):
    if not url or str(url).strip() == "":
        print(f"--> Peringatan: URL untuk {output_name} kosong di config.conf! Proses unduh dilewati.")
        return

    if '/d/' in url:
        base_id = url.split('/d/')[1].split('/')[0]
        download_url = f"https://docs.google.com/spreadsheets/d/{base_id}/export?format=xlsx"
    else:
        print(f"--> URL tidak valid untuk {output_name}: {url}")
        return

    try:
        print(f"--> Sedang mengunduh {output_name}...")
        response = requests.get(download_url)
        response.raise_for_status()
        with open(output_name, 'wb') as f:
            f.write(response.content)
        print(f"--> Berhasil! File disimpan sebagai {output_name}")
    except Exception as e:
        print(f"--> Gagal mengunduh {output_name}: {e}")

def main():
    config = configparser.ConfigParser(allow_no_value=True)
    config_file = 'config.conf'
    
    if not os.path.exists(config_file):
        print("--> File config.conf tidak ditemukan!")
        return

    config.read(config_file)

    mapping = {
        'FBACK': 'Minifs_temp.xlsx'
    }

    for section, output_name in mapping.items():
        if config.has_section(section):
            options = config.options(section)
            if options:
                if options[0] == 'url':
                    url = config.get(section, 'url')
                else:
                    url = options[0]
                
                download_sheet(url, output_name)
        else:
            print(f"--> Section [{section}] tidak ditemukan di config.conf")

if __name__ == "__main__":
    main()