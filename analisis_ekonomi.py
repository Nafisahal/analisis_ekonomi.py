
"""PRINT 2026 — analisis perubahan harga pangan dari waktu ke waktu.
Hanya menggunakan file Excel resmi panitia. Tidak menggunakan ML.
Instal: pip install pandas matplotlib openpyxl numpy
"""
from pathlib import Path
import re
import hashlib
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# ==================================================
# PERSIAPAN
# ==================================================
folder = Path(__file__).resolve().parent
output = folder / 'hasil_PRINT2026_fokus_biaya_hidup'
output.mkdir(exist_ok=True)

# Bisa menggunakan salah satu nama file Excel resmi yang sudah dipakai.
pilihan_file = [
    folder / 'print_2026_ekonomi.csv.xlsx',
    folder / 'print_2026_ekonomi.csv(1).xlsx',
]
file_data = next((p for p in pilihan_file if p.exists()), None)
if file_data is None:
    raise FileNotFoundError('Letakkan file Excel resmi PRINT 2026 di folder kode.')
print('DATASET YANG DIBACA:', file_data.resolve())
print('UKURAN FILE:', file_data.stat().st_size, 'byte')
print('SHA256 DATASET:', hashlib.sha256(file_data.read_bytes()).hexdigest())

def nama_baku(nilai, alias):
    if pd.isna(nilai):
        return pd.NA
    teks = re.sub(r'\s+', ' ', str(nilai).strip()).casefold()
    return alias.get(teks, teks.title())

def bulan_baku(nilai):
    if pd.isna(nilai):
        return pd.NaT
    try:
        return pd.to_datetime(nilai, errors='coerce').to_period('M').to_timestamp()
    except (ValueError, TypeError):
        return pd.NaT

def angka_baku(nilai):
    if pd.isna(nilai):
        return np.nan
    if isinstance(nilai, (int, float, np.number)):
        return float(nilai)
    teks = str(nilai).strip().replace(' ', '')
    if ',' in teks and '.' in teks:
        teks = teks.replace('.', '').replace(',', '.')
    elif ',' in teks:
        teks = teks.replace(',', '.')
    return pd.to_numeric(teks, errors='coerce')

alias_provinsi = {
    'ntb': 'Nusa Tenggara Barat', 'ntt': 'Nusa Tenggara Timur',
    'jabar': 'Jawa Barat', 'jateng': 'Jawa Tengah', 'jatim': 'Jawa Timur',
    'sumut': 'Sumatera Utara', 'sumbar': 'Sumatera Barat',
    'sumsel': 'Sumatera Selatan', 'dki jakarta': 'DKI Jakarta',
    'dki jkt': 'DKI Jakarta', 'di yogyakarta': 'DI Yogyakarta',
    'd.i. yogyakarta': 'DI Yogyakarta', 'kep. riau': 'Kepulauan Riau',
    'kep. bangka belitung': 'Kepulauan Bangka Belitung',
    'kalbar': 'Kalimantan Barat', 'kalteng': 'Kalimantan Tengah',
    'kalsel': 'Kalimantan Selatan', 'kaltim': 'Kalimantan Timur',
    'kaltara': 'Kalimantan Utara', 'sulut': 'Sulawesi Utara',
    'sulteng': 'Sulawesi Tengah', 'sulsel': 'Sulawesi Selatan',
    'sultra': 'Sulawesi Tenggara', 'sulbar': 'Sulawesi Barat',
}
alias_komoditas = {
    'bwg merah': 'Bawang Merah', 'bawang merah': 'Bawang Merah',
    'beras med.': 'Beras Medium', 'beras medium': 'Beras Medium',
    'daging ayam': 'Daging Ayam Ras', 'daging ayam ras': 'Daging Ayam Ras',
    'cabai rawit': 'Cabai Rawit Merah', 'cabai rawit merah': 'Cabai Rawit Merah',
    'minyak curah': 'Minyak Goreng Curah',
    'minyak goreng curah': 'Minyak Goreng Curah',
    'gula pasir': 'Gula Konsumsi', 'gula konsumsi': 'Gula Konsumsi',
}
komoditas_resmi = set(alias_komoditas.values())

# ==================================================
# 1. DATA CLEANING
# ==================================================
print('\n=== 1. DATA CLEANING ===')
mentah = pd.read_excel(file_data)
df = mentah.copy()
df.columns = df.columns.str.strip().str.lower()
wajib = {'bulan', 'provinsi', 'komoditas', 'harga_median',
         'hari_berharga', 'cakupan_persen'}
if not wajib.issubset(df.columns):
    raise ValueError(f'Kolom wajib tidak ditemukan: {sorted(wajib - set(df.columns))}')

df['provinsi'] = df['provinsi'].map(lambda x: nama_baku(x, alias_provinsi))
df['komoditas'] = df['komoditas'].map(lambda x: nama_baku(x, alias_komoditas))
df['bulan'] = df['bulan'].map(bulan_baku)
# Simpan identitas baris asal untuk audit
df['baris_asal_excel'] = np.arange(2, len(df) + 2)
for kolom in [c for c in df.columns if c not in ['bulan', 'provinsi', 'komoditas', 'qris_status_data', 'baris_asal_excel']]:
    df[kolom] = df[kolom].map(angka_baku)

duplikat = df.drop(columns='baris_asal_excel').duplicated(keep='first')
df.loc[duplikat].to_excel(output / '01_duplikat_identik.xlsx', index=False)
df = df.loc[~duplikat].copy()
df.to_excel(output / '02_data_setelah_cleaning.xlsx', index=False)
print('Baris mentah:', len(mentah), '| Duplikat identik:', int(duplikat.sum()))

# ==================================================
# 2. DATA VALIDATION
# ==================================================
print('\n=== 2. DATA VALIDATION ===')
flag = pd.DataFrame(index=df.index)
flag['bulan_tidak_valid'] = df['bulan'].isna()
flag['provinsi_kosong'] = df['provinsi'].isna()
flag['komoditas_tidak_dikenal'] = ~df['komoditas'].isin(komoditas_resmi)
flag['harga_kosong_atau_tidak_positif'] = df['harga_median'].isna() | df['harga_median'].le(0)
flag['hari_tidak_wajar'] = df['hari_berharga'].notna() & (
    df['hari_berharga'].lt(0) | df['hari_berharga'].gt(df['bulan'].dt.days_in_month))
flag['cakupan_tidak_wajar'] = df['cakupan_persen'].notna() & ~df['cakupan_persen'].between(0, 100)
# Pemeriksaan indikator tambahan dari kode lama (tidak menjadi fokus analisis).
if 'usd_idr_median' in df:
    flag['kurs_tidak_valid'] = df['usd_idr_median'].notna() & df['usd_idr_median'].le(0)
for kolom in ['internet_laki_laki_2023', 'internet_perempuan_2023']:
    if kolom in df:
        flag[kolom + '_tidak_valid'] = df[kolom].notna() & ~df[kolom].between(0, 100)
for kolom in ['qris_pengguna_juta', 'qris_merchant_juta',
              'qris_nominal_transaksi_triliun_idr', 'qris_volume_transaksi_miliar']:
    if kolom in df:
        flag[kolom + '_negatif'] = df[kolom].notna() & df[kolom].lt(0)
flag['cakupan_perlu_cek_skala'] = (df['cakupan_persen'].notna()
    & df['cakupan_persen'].between(0, 1) & df['hari_berharga'].gt(0))
flag['kunci_berulang'] = df.duplicated(['bulan', 'provinsi', 'komoditas'], keep=False)
flag['hari_kosong'] = df['hari_berharga'].isna()
flag['cakupan_kosong'] = df['cakupan_persen'].isna()
flag.to_excel(output / '03_flag_validasi.xlsx', index=False)
df.loc[flag.any(axis=1)].join(flag).to_excel(
    output / '04_data_perlu_diperiksa.xlsx', index=False)

# Hanya baris dengan waktu, wilayah, komoditas dan harga valid dipakai.
# Flag hari/cakupan tidak otomatis menghapus baris: perlu audit manual.
valid = (~flag['bulan_tidak_valid'] & ~flag['provinsi_kosong']
         & ~flag['komoditas_tidak_dikenal']
         & ~flag['harga_kosong_atau_tidak_positif'])
analisis = df.loc[valid].copy()
analisis.to_excel(output / '05_data_harga_valid.xlsx', index=False)
# Audit kombinasi kunci berulang, termasuk harga yang berbeda
(df.loc[flag['kunci_berulang']]
 .sort_values(['bulan', 'provinsi', 'komoditas'])
 .to_excel(output / '05c_audit_kunci_berulang.xlsx', index=False))
print('Baris harga valid:', len(analisis))
print('Ringkasan flag validasi:')
print(flag.sum().to_string())
# Simpan data yang tidak lolos syarat harga, tanpa mengubah data mentah.
df.loc[~valid].to_excel(output / '05b_data_tidak_dianalisis.xlsx', index=False)

# ==================================================
# 3. DATA EXPLORATION (EDA)
# ==================================================
print('\n=== 3. DATA EXPLORATION ===')
# Satu nilai harga untuk tiap provinsi, bulan, dan komoditas.
# Jika ada beberapa catatan dengan kunci sama, gunakan rata-ratanya.
per_provinsi = (analisis.groupby(['bulan', 'provinsi', 'komoditas'], as_index=False)
                .agg(harga=('harga_median', 'mean'),
                     jumlah_catatan=('harga_median', 'size')))
per_provinsi.to_excel(output / '06_harga_provinsi_bulanan.xlsx', index=False)
# Perubahan per provinsi dihitung hanya pada bulan kalender berurutan
per_provinsi = per_provinsi.sort_values(['provinsi', 'komoditas', 'bulan']).copy()
per_provinsi['bulan_lalu'] = per_provinsi.groupby(['provinsi', 'komoditas'])['bulan'].shift()
per_provinsi['harga_lalu'] = per_provinsi.groupby(['provinsi', 'komoditas'])['harga'].shift()
mask_berurutan = per_provinsi['bulan_lalu'].eq(per_provinsi['bulan'] - pd.DateOffset(months=1))
per_provinsi['perubahan_bulanan_persen'] = np.where(
    mask_berurutan & per_provinsi['harga_lalu'].gt(0),
    (per_provinsi['harga'] / per_provinsi['harga_lalu'] - 1) * 100, np.nan)
per_provinsi.to_excel(output / '06b_perubahan_bulanan_provinsi.xlsx', index=False)

# Tiap provinsi memiliki bobot yang sama dalam rata-rata bulanan.
tren = (per_provinsi.groupby(['bulan', 'komoditas'], as_index=False)
        .agg(harga_rata=('harga', 'mean'),
             harga_min=('harga', 'min'),
             harga_max=('harga', 'max'),
             jumlah_provinsi=('provinsi', 'nunique'))
        .sort_values(['komoditas', 'bulan']))
tren['bulan_sebelumnya'] = tren.groupby('komoditas')['bulan'].shift()
tren['harga_sebelumnya'] = tren.groupby('komoditas')['harga_rata'].shift()
berurutan = tren['bulan_sebelumnya'].eq(tren['bulan'] - pd.DateOffset(months=1))
tren['perubahan_bulanan_persen'] = np.where(
    berurutan & tren['harga_sebelumnya'].gt(0),
    (tren['harga_rata'] / tren['harga_sebelumnya'] - 1) * 100,
    np.nan)
tren.to_excel(output / '07_tren_dan_perubahan_bulanan.xlsx')
# Analisis tambahan: rata-rata perubahan pada provinsi yang punya dua bulan berurutan
tren_berpasangan = (per_provinsi.groupby(['bulan', 'komoditas'], as_index=False)
    .agg(rata_perubahan_provinsi_persen=('perubahan_bulanan_persen', 'mean'),
         jumlah_provinsi_berpasangan=('perubahan_bulanan_persen', 'count')))
tren_berpasangan.to_excel(output / '07d_perubahan_bulanan_provinsi_sebanding.xlsx', index=False)

# Statistik deskriptif lama, tetap dipertahankan sebagai pendukung.
statistik = analisis.groupby('komoditas')['harga_median'].agg(
    jumlah='count', rata_rata='mean', median='median',
    minimum='min', maksimum='max', simpangan_baku='std').reset_index()
statistik.to_excel(output / '07b_statistik_per_komoditas.xlsx', index=False)
statistik_provinsi = (per_provinsi.groupby(['provinsi', 'komoditas'])['harga']
    .agg(jumlah_bulan='count', rata_rata='mean', minimum='min', maksimum='max')
    .reset_index())
statistik_provinsi.to_excel(output / '07c_statistik_provinsi_komoditas.xlsx', index=False)

# Bandingkan awal-akhir hanya pada provinsi yang tersedia di kedua bulan.
awal, akhir = per_provinsi['bulan'].min(), per_provinsi['bulan'].max()
a = per_provinsi.loc[per_provinsi['bulan'].eq(awal),
                     ['provinsi', 'komoditas', 'harga']].rename(columns={'harga': 'harga_awal'})
b = per_provinsi.loc[per_provinsi['bulan'].eq(akhir),
                     ['provinsi', 'komoditas', 'harga']].rename(columns={'harga': 'harga_akhir'})
pasangan = a.merge(b, on=['provinsi', 'komoditas'], how='inner')
pasangan['perubahan_persen'] = (pasangan['harga_akhir'] / pasangan['harga_awal'] - 1) * 100
pasangan.to_excel(output / '08_perubahan_per_provinsi.xlsx', index=False)
ringkasan = (pasangan.groupby('komoditas', as_index=False)
             .agg(harga_awal=('harga_awal', 'mean'),
                  harga_akhir=('harga_akhir', 'mean'),
                  jumlah_provinsi=('provinsi', 'nunique')))
ringkasan['perubahan_persen'] = (
    (ringkasan['harga_akhir'] / ringkasan['harga_awal'] - 1) * 100)
ringkasan = ringkasan.sort_values('perubahan_persen', ascending=False)
ringkasan.to_excel(output / '09_ringkasan_awal_akhir.xlsx', index=False)
print('\nPerbandingan harga awal dan akhir:')
print(ringkasan.to_string(index=False))

# ==================================================
# 4. VISUALISASI KHUSUS PERUBAHAN HARGA PANGAN
# ==================================================
print('\n=== 4. VISUALISASI PERUBAHAN HARGA PANGAN ===')

# Gunakan provinsi yang berpasangan untuk menghitung perubahan bulanan;
# perbedaan cakupan provinsi tidak otomatis dianggap perubahan harga.
komoditas_urut = sorted(per_provinsi['komoditas'].dropna().unique())

def simpan(nama):
    plt.tight_layout()
    plt.savefig(output / nama, dpi=200, bbox_inches='tight')
    plt.close()

# GRAFIK 01: Indeks harga (awal=100), agar semua komoditas sebanding
# tanpa membandingkan nominal harga lintas satuan komoditas.
indeks = tren.sort_values(['komoditas', 'bulan']).copy()
indeks['harga_awal_indeks'] = indeks.groupby('komoditas')['harga_rata'].transform('first')
indeks['indeks_100'] = indeks['harga_rata'] / indeks['harga_awal_indeks'] * 100
plt.figure(figsize=(12, 6))
for nama, sub in indeks.groupby('komoditas'):
    plt.plot(sub['bulan'], sub['indeks_100'], marker='o', markersize=3, label=nama)
plt.axhline(100, color='gray', linestyle='--', linewidth=1)
plt.title('Indeks Harga Pangan per Komoditas (Bulan Pertama = 100)')
plt.xlabel('Bulan'); plt.ylabel('Indeks harga (100 = bulan awal)')
plt.legend(fontsize=8, ncol=2); plt.grid(alpha=.25)
simpan('01_INDEKS_HARGA_100.png')

# GRAFIK 02: Tren harga rupiah masing-masing komoditas, terpisah.
for i, (nama, sub) in enumerate(tren.groupby('komoditas'), 1):
    plt.figure(figsize=(11, 4.5))
    plt.plot(sub['bulan'], sub['harga_rata'], marker='o', linewidth=2)
    plt.title(f'Tren Rata-Rata Harga Bulanan: {nama}')
    plt.xlabel('Bulan'); plt.ylabel('Harga rata-rata (Rp, satuan sumber)')
    plt.grid(alpha=.25)
    simpan(f'02_TREN_HARGA_{i:02d}.png')

# GRAFIK 03: Heatmap perubahan bulanan, hanya provinsi yang berpasangan.
heat = tren_berpasangan.pivot(index='komoditas', columns='bulan',
                              values='rata_perubahan_provinsi_persen')
if not heat.empty:
    fig, ax = plt.subplots(figsize=(max(12, len(heat.columns)*.55), 5.5))
    batas = np.nanmax(np.abs(heat.to_numpy(dtype=float)))
    if not np.isfinite(batas) or batas == 0: batas = 1
    gambar = ax.imshow(heat.to_numpy(dtype=float), aspect='auto',
                       cmap='RdBu_r', vmin=-batas, vmax=batas)
    ax.set_yticks(range(len(heat.index)), heat.index)
    ax.set_xticks(range(len(heat.columns)),
                  [x.strftime('%Y-%m') for x in heat.columns], rotation=65, ha='right')
    ax.set_title('Heatmap Perubahan Harga Bulanan (%) — Provinsi Berpasangan')
    fig.colorbar(gambar, ax=ax, label='Rata-rata perubahan (%)')
    fig.tight_layout()
    fig.savefig(output / '03_HEATMAP_PERUBAHAN_BULANAN.png', dpi=200, bbox_inches='tight')
    plt.close(fig)

# GRAFIK 04: Perubahan awal-akhir, hanya provinsi yang ada di dua periode.
if not ringkasan.empty:
    urut = ringkasan.sort_values('perubahan_persen')
    plt.figure(figsize=(10, 5))
    plt.barh(urut['komoditas'], urut['perubahan_persen'])
    plt.axvline(0, color='black', linewidth=1)
    plt.title(f'Perubahan Harga Awal–Akhir ({awal:%Y-%m} s.d. {akhir:%Y-%m})')
    plt.xlabel('Perubahan rata-rata harga (%) — provinsi sebanding')
    simpan('04_PERUBAHAN_AWAL_AKHIR.png')

# GRAFIK 05: Sebaran perubahan per provinsi, bukan sekadar rerata nasional.
if not pasangan.empty:
    kelompok = [pasangan.loc[pasangan['komoditas'].eq(n),
                            'perubahan_persen'].dropna().to_numpy() for n in komoditas_urut]
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.boxplot(kelompok, tick_labels=komoditas_urut, showfliers=True)
    ax.axhline(0, color='black', linewidth=1)
    ax.set_title('Sebaran Perubahan Harga Awal–Akhir Antarprovinsi')
    ax.set_ylabel('Perubahan harga (%)')
    ax.tick_params(axis='x', rotation=25)
    ax.grid(axis='y', alpha=.25)
    fig.tight_layout()
    fig.savefig(output / '05_SEBARAN_PERUBAHAN_PROVINSI.png', dpi=200, bbox_inches='tight')
    plt.close(fig)

# GRAFIK 06: Perubahan bulanan menggunakan provinsi berpasangan.
plt.figure(figsize=(12, 6))
for nama, sub in tren_berpasangan.groupby('komoditas'):
    plt.plot(sub['bulan'], sub['rata_perubahan_provinsi_persen'],
             marker='o', markersize=3, label=nama)
plt.axhline(0, color='black', linewidth=1)
plt.title('Perubahan Harga Bulanan (%) Berdasarkan Provinsi Berpasangan')
plt.xlabel('Bulan'); plt.ylabel('Rata-rata perubahan harga (%)')
plt.legend(fontsize=8, ncol=2); plt.grid(alpha=.25)
simpan('06_PERUBAHAN_BULANAN_SEBANDING.png')

# GRAFIK 07: Cakupan data untuk menilai keterbandingan tren.
plt.figure(figsize=(12, 5))
for nama, sub in tren.groupby('komoditas'):
    plt.plot(sub['bulan'], sub['jumlah_provinsi'], label=nama)
plt.title('Jumlah Provinsi yang Tercakup per Bulan')
plt.xlabel('Bulan'); plt.ylabel('Jumlah provinsi')
plt.legend(fontsize=8, ncol=2); plt.grid(alpha=.25)
simpan('07_CAKUPAN_PROVINSI.png')

# GRAFIK 08: Provinsi dengan perubahan harga terbesar/terkecil per komoditas.
# Urutan di sini hanya deskripsi dataset, bukan peringkat kesejahteraan.
for i, nama in enumerate(komoditas_urut, 1):
    sub = pasangan.loc[pasangan['komoditas'].eq(nama)].sort_values('perubahan_persen')
    if sub.empty: continue
    plt.figure(figsize=(10, max(5, len(sub)*.27)))
    plt.barh(sub['provinsi'], sub['perubahan_persen'])
    plt.axvline(0, color='black', linewidth=.8)
    plt.title(f'Perubahan Harga Awal–Akhir per Provinsi: {nama}')
    plt.xlabel('Perubahan harga (%)')
    simpan(f'08_PROVINSI_{i:02d}.png')

# Audit gambar baru: lokasi, waktu, dan checksum.
grafik_terbuat = sorted(output.glob('*.png'))
log = [f'Waktu dijalankan: {datetime.now().isoformat(timespec="seconds")}',
       f'Dataset: {file_data.resolve()}',
       f'Kode: {Path(__file__).resolve()}',
       f'Jumlah grafik: {len(grafik_terbuat)}', '']
for f in grafik_terbuat:
    log.append(f'{f.name} | {f.stat().st_size} byte | {hashlib.sha256(f.read_bytes()).hexdigest()[:16]}')
(output / '00_AUDIT_GRAFIK_TERBARU.txt').write_text('\n'.join(log), encoding='utf-8')
print('Grafik baru dibuat:', len(grafik_terbuat))
print('Folder grafik:', output.resolve())

# ==================================================
# 5. RINGKASAN ANGKA (BUKAN KESIMPULAN TIM)
# ==================================================
ringkasan_proses = f'''RINGKASAN TEKNIS ANALISIS HARGA PANGAN
Sumber data: {file_data.name}
Data mentah: {len(mentah)} baris
Duplikat identik dihapus: {int(duplikat.sum())} baris
Data harga valid: {len(analisis)} baris
Periode data: {awal:%Y-%m} sampai {akhir:%Y-%m}
Jumlah komoditas: {per_provinsi['komoditas'].nunique()}
Jumlah provinsi: {per_provinsi['provinsi'].nunique()}

METODE:
1. Standardisasi nama, bulan, dan angka; audit duplikat dan nilai tidak valid.
2. Agregasi catatan per provinsi-bulan-komoditas terlebih dahulu.
3. Rata-rata antarprovinsi dihitung dengan bobot provinsi yang sama.
4. Perubahan bulanan hanya dihitung untuk dua bulan kalender berurutan.
5. Perbandingan awal-akhir hanya memakai provinsi yang ada di kedua periode.
6. Data harga tidak membuktikan perubahan total biaya hidup.
7. Interpretasi hasil dan rekomendasi wajib disusun oleh tim sendiri.
7a. Tabel 07d menggunakan pasangan provinsi yang sama antarbulan; lebih cocok untuk menguji perubahan ketika cakupan berbeda.
8. Periksa satuan harga dari dokumentasi panitia sebelum membandingkan lintas komoditas.
9. Grafik pendukung lama tidak digunakan untuk menyimpulkan total biaya hidup.
10. Catatan berulang dengan kunci sama dirata-ratakan; audit sebelum interpretasi.
11. Baris asal Excel dipertahankan dalam data audit, bukan mengubah file sumber.
12. Kolom QRIS dan indikator pendukung bukan bukti penyebab perubahan harga.
13. Grafik baru memprioritaskan indeks harga, perubahan bulanan, perbandingan wilayah, dan cakupan data.
'''
(output / '18_ringkasan_teknis.txt').write_text(ringkasan_proses, encoding='utf-8')
print('\n=== SELESAI ===')
print(ringkasan_proses)
print('Semua hasil disimpan di:', output)
