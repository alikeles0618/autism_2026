import csv
import math
import numpy as np

with open('/Users/alik/Desktop/anti_projeler/articles_all/autism_2026_testtekrar/02_data/cleaned_data.csv', 'r', encoding='utf-8') as f:
    data = list(csv.DictReader(f))

n = len(data)

def ols_regression(Y, X, x_names, y_name):
    # Add intercept column
    X_mat = np.column_stack([np.ones(n), X])
    k = X.shape[1] # number of predictors
    
    # OLS coefficients: beta = (X'X)^(-1) X'Y
    XtX = np.dot(X_mat.T, X_mat)
    XtX_inv = np.linalg.inv(XtX)
    beta = np.dot(np.dot(XtX_inv, X_mat.T), Y)
    
    # Predictions and residuals
    Y_pred = np.dot(X_mat, beta)
    residuals = Y - Y_pred
    
    # Sum of squares
    SS_tot = np.sum((Y - np.mean(Y)) ** 2)
    SS_res = np.sum(residuals ** 2)
    SS_reg = SS_tot - SS_res
    
    df_reg = k
    df_res = n - k - 1
    
    R2 = 1.0 - (SS_res / SS_tot)
    R2_adj = 1.0 - ((1.0 - R2) * (n - 1) / df_res)
    
    MS_reg = SS_reg / df_reg
    MS_res = SS_res / df_res
    F_stat = MS_reg / MS_res
    
    # Residual variance and SE of coefficients
    s2 = SS_res / df_res
    cov_beta = s2 * XtX_inv
    se_beta = np.sqrt(np.diagonal(cov_beta))
    t_beta = beta / se_beta
    
    print(f"\n=======================================================")
    print(f"ÇOKLU DOĞRUSAL REGRESYON MODELİ: {y_name}")
    print(f"R² = {R2:.4f} | Düzeltilmiş R² = {R2_adj:.4f} | F({df_reg}, {df_res}) = {F_stat:.3f}")
    print(f"-------------------------------------------------------")
    print(f"{'Değişken':<25} | {'Beta (B)':<10} | {'Standart Hata':<14} | {'t':<8}")
    print(f"-------------------------------------------------------")
    names = ['Sabit (Intercept)'] + x_names
    for name, b, se, t in zip(names, beta, se_beta, t_beta):
        print(f"{name:<25} | {b:>9.4f}  | {se:>12.4f}   | {t:>7.3f}")
    print(f"=======================================================")

# Model 1: Pro-Agility Değişimi (DELTA_CEVIKLIK)
# Y: DELTA_CEVIKLIK
# X: DELTA_GORSEL, DELTA_AKTIF_SIC, YAS_2021, APE_INDEX_PRE
Y1 = np.array([float(r['DELTA_CEVIKLIK']) for r in data])
X1 = np.column_stack([
    [float(r['DELTA_GORSEL']) for r in data],
    [float(r['DELTA_AKTIF_SIC']) for r in data],
    [float(r['YAS_2021']) for r in data],
    [float(r['BOY_PRE']) for r in data]
])
x_names1 = ['Δ Görsel Reaksiyon', 'Δ Aktif Sıçrama', 'Takvim Yaşı (2021)', 'Başlangıç Boyu']
ols_regression(Y1, X1, x_names1, "Δ Pro-Agility Çeviklik Süresi (s)")

# Model 2: Sağ El Kavrama Değişimi (DELTA_SAG_EL)
# Y: DELTA_SAG_EL
# X: KARIS_PRE, KULAC_PRE, VKI_PRE, PHV_OFFSET_PRE
Y2 = np.array([float(r['DELTA_SAG_EL']) for r in data])
X2 = np.column_stack([
    [float(r['KARIS_PRE']) for r in data],
    [float(r['KULAC_PRE']) for r in data],
    [float(r['VKI_PRE']) for r in data],
    [float(r['PHV_OFFSET_PRE']) for r in data]
])
x_names2 = ['Başlangıç Karış (cm)', 'Başlangıç Kulaç (cm)', 'Başlangıç VKİ', 'PHV Maturity Offset']
ols_regression(Y2, X2, x_names2, "Δ Sağ El Kavrama Kuvveti (kg)")
