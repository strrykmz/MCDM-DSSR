import numpy as np

def hitung_cr(matriks, bobot):
    #  Lambda Max
    n = matriks.shape[0]
    lambda_max = np.mean(np.dot(matriks, bobot) / bobot)
    
    #  Consistency Index (CI)
    ci = (lambda_max - n) / (n - 1)
    
    # Random Index (RI) untuk n=4 adalah 0.90 dalam metode AHP
    ri = 0.90 
    
    #  Consistency Ratio (CR)
    cr = ci / ri
    
    return cr

# Matriks Perbandingan
matriks_perbandingan = np.array([
    [1.,   2.,   3.,   1.],
    [0.5,  1.,   2.,   0.5],
    [0.33, 0.5,  1.,   0.33],
    [1.,   2.,   3.,   1.]
])

bobot = np.array([0.35092087, 0.18934932, 0.10880894, 0.35092087])

nilai_cr = hitung_cr(matriks_perbandingan, bobot)

print(f"Consistency Ratio (CR): {nilai_cr:.5f}")
if nilai_cr < 0.1:
    print("STATUS: KONSISTEN (Bobot Valid)")
else:
    print("STATUS: TIDAK KONSISTEN (Perbaiki Matriks)")