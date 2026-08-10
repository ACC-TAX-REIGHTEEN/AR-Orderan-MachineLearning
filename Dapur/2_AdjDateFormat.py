import pandas as pd

df = pd.read_excel('ARClean_temp.xlsx')

mapping_bulan = {
    1: 'Jan',  2: 'Feb',  3: 'Mar',  4: 'Apr',  5: 'Mei',  6: 'Jun',
    7: 'Jul',  8: 'Agu',  9: 'Sep',  10: 'Okt', 11: 'Nov', 12: 'Des'
}

def ubah_format(nilai):
    if pd.isna(nilai):
        return nilai
    dt = pd.to_datetime(nilai)
    return f"{dt.day} {mapping_bulan[dt.month]} {dt.year}"

df['Tgl Faktur'] = df['Tgl Faktur'].apply(ubah_format)
df['Jatuh Tempo'] = df['Jatuh Tempo'].apply(ubah_format)

df.to_excel('ARClean_temp.xlsx', index=False)

print("--> Format tanggal pada kolom Tgl Faktur dan Jatuh Tempo berhasil diubah.")
