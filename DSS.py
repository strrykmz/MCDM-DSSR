import pandas as pd
import numpy as np
from tabulate import tabulate

# ==========================================
# KONFIGURASI KRITERIA
# ==========================================
# Sesuaikan tipe kriteria berdasarkan data Anda
# C1: Harga (Cost), C2: Luas (Benefit), C3: Kepadatan (Benefit), C4: Jarak (Cost)
KRITERIA_TYPE = ['cost', 'benefit', 'benefit', 'cost']

# ==========================================
# 1. METODE AHP (Untuk Menentukan Bobot)
# ==========================================
def hitung_bobot_ahp():
    """
    Menghitung bobot kriteria menggunakan matriks perbandingan berpasangan AHP.
    Anda bisa mengubah nilai matriks di bawah sesuai preferensi.
    """
    print("\n--- [1] Menghitung Bobot dengan AHP ---")
    
    # Matriks Perbandingan Berpasangan (Pairwise Comparison)
    # Urutan: [Harga, Luas, Kepadatan, Jarak]
    # Contoh: Harga 2x lebih penting dari Jarak, dll.
    # Jika ragu, gunakan matriks identitas (semua 1) untuk bobot sama rata.
    matriks_perbandingan = np.array([
        [1.00, 2.00, 3.00, 1.00],  # Harga vs yang lain
        [0.50, 1.00, 2.00, 0.50],  # Luas vs yang lain
        [0.33, 0.50, 1.00, 0.33],  # Kepadatan vs yang lain
        [1.00, 2.00, 3.00, 1.00]   # Jarak vs yang lain
    ])

    # Hitung Bobot (Normalisasi Kolom -> Rata-rata Baris)
    col_sum = matriks_perbandingan.sum(axis=0)
    normalized_matrix = matriks_perbandingan / col_sum
    bobot_prioritas = normalized_matrix.mean(axis=1)

    print("Matriks Perbandingan AHP:")
    print(matriks_perbandingan)
    print(f"\nBobot Hasil AHP (W): {bobot_prioritas}")
    print(f"Total Bobot: {sum(bobot_prioritas)}") # Harus 1.0
    
    return bobot_prioritas

# ==========================================
# 2. METODE SAW (Simple Additive Weighting)
# ==========================================
def hitung_saw(df, weights, types):
    print("\n--- [2] Proses Perhitungan SAW ---")
    data = df.iloc[:, 3:7].values.astype(float) # Ambil kolom C1-C4
    normalized_data = np.zeros_like(data)

    # Normalisasi
    for j in range(len(types)):
        if types[j] == 'benefit':
            normalized_data[:, j] = data[:, j] / np.max(data[:, j])
        else: # cost
            normalized_data[:, j] = np.min(data[:, j]) / data[:, j]

    # Perankingan
    saw_score = np.sum(normalized_data * weights, axis=1)
    return saw_score

# ==========================================
# 3. METODE WP (Weighted Product)
# ==========================================
def hitung_wp(df, weights, types):
    print("\n--- [3] Proses Perhitungan WP ---")
    data = df.iloc[:, 3:7].values.astype(float)
    
    # Perbaikan Bobot (Total harus 1 - sudah dilakukan di AHP)
    # Pangkat positif untuk Benefit, negatif untuk Cost
    w_pangkat = []
    for i, t in enumerate(types):
        if t == 'benefit':
            w_pangkat.append(weights[i])
        else:
            w_pangkat.append(-1 * weights[i])
    
    # Menghitung Vektor S
    # Rumus: S = Product(Data^Bobot)
    s_vector = np.prod(data ** w_pangkat, axis=1)
    
    # Menghitung Vektor V (Preferensi Relatif)
    v_vector = s_vector / s_vector.sum()
    
    return v_vector

# ==========================================
# 4. METODE TOPSIS
# ==========================================
def hitung_topsis(df, weights, types):
    print("\n--- [4] Proses Perhitungan TOPSIS ---")
    data = df.iloc[:, 3:7].values.astype(float)
    
    # 1. Normalisasi Matriks (Pembagi Akar Kuadrat)
    pembagi = np.sqrt(np.sum(data**2, axis=0))
    norm_matrix = data / pembagi
    
    # 2. Matriks Ternormalisasi Terbobot (Y)
    weighted_matrix = norm_matrix * weights
    
    # 3. Solusi Ideal Positif (A+) dan Negatif (A-)
    a_plus = []
    a_min = []
    
    for i in range(len(types)):
        if types[i] == 'benefit':
            a_plus.append(np.max(weighted_matrix[:, i]))
            a_min.append(np.min(weighted_matrix[:, i]))
        else: # cost
            a_plus.append(np.min(weighted_matrix[:, i])) # Cost terendah adalah ideal positif
            a_min.append(np.max(weighted_matrix[:, i]))  # Cost tertinggi adalah ideal negatif
            
    # 4. Menghitung Jarak (Separation Measure)
    dist_plus = np.sqrt(np.sum((weighted_matrix - a_plus)**2, axis=1))
    dist_min = np.sqrt(np.sum((weighted_matrix - a_min)**2, axis=1))
    
    # 5. Nilai Preferensi (V)
    topsis_score = dist_min / (dist_min + dist_plus)
    
    return topsis_score

# ==========================================
# MAIN EXECUTION
# ==========================================
def main():
    try:
        # Load Data Excel
        file_path = 'DSS.xlsx' 
        df = pd.read_excel(file_path)
        
        print(f"Berhasil memuat data: {len(df)} baris ditemukan.")
        
        # --- PERBAIKAN DI SINI ---
        # Kita tampilkan nama kolom asli untuk pengecekan
        print(f"Nama kolom asli di Excel: {df.columns.tolist()}")
        
        # Kita ubah nama kolom secara paksa agar sesuai dengan script
        # Asumsi urutan kolom di Excel Anda: 
        # [0]Alternatif, [1]Lokasi, [2]Kategori, [3]C1, [4]C2, [5]C3, [6]C4
        # Jika kolom lebih dari 7, kita hanya rename 7 pertama atau sesuaikan list ini
        if len(df.columns) >= 7:
            df.columns = ['Alternatif', 'Lokasi', 'Kategori', 'C1', 'C2', 'C3', 'C4'] + df.columns.tolist()[7:]
            print("Nama kolom telah dinormalisasi menjadi: ", df.columns.tolist())
        else:
            print("PERINGATAN: Jumlah kolom kurang dari 7. Pastikan format Excel sesuai.")
        # -------------------------
        
        # 1. Hitung Bobot AHP
        weights = hitung_bobot_ahp()
        
        # 2. Hitung SAW
        df['SAW_Score'] = hitung_saw(df, weights, KRITERIA_TYPE)
        
        # 3. Hitung WP
        df['WP_Score'] = hitung_wp(df, weights, KRITERIA_TYPE)
        
        # 4. Hitung TOPSIS
        df['TOPSIS_Score'] = hitung_topsis(df, weights, KRITERIA_TYPE)
        
        # --- RANGKUMAN HASIL ---
        print("\n==========================================")
        print("          HASIL PERHITUNGAN AKHIR         ")
        print("==========================================")
        
        # Pilih kolom untuk ditampilkan
        # Sekarang aman karena nama kolom sudah diubah jadi 'Lokasi'
        result_df = df[['Alternatif', 'Lokasi', 'SAW_Score', 'WP_Score', 'TOPSIS_Score']].copy()
        
        # Urutkan berdasarkan TOPSIS tertinggi
        result_df = result_df.sort_values(by='TOPSIS_Score', ascending=False)
        
        # Tampilkan tabel
        print(tabulate(result_df.head(50), headers='keys', tablefmt='pretty', showindex=False))
        print("\n... (Menampilkan Top 50 Data) ...")
        
        # Simpan ke Excel baru
        output_file = 'Hasil_Perhitungan_DSS.xlsx'
        df.to_excel(output_file, index=False)
        print(f"\nHasil lengkap telah disimpan ke file: {output_file}")
        
    except FileNotFoundError:
        print("Error: File 'DSS.xlsx' tidak ditemukan. Pastikan file ada di folder yang sama.")
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")
        # Tambahan debug untuk melihat error lebih detail
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()