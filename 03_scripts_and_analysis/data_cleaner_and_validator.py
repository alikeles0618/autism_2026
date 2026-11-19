"""
Otizm Spektrum Bozukluğu (OSB) Retrospektif Veri Seti Temizleme ve Doğrulama Betiği
Tarih Bazı: 2021
Örneklem: N = 57 Erkek Katılımcı
Tasarım: 14 Hafta (1 Hafta Ön Test + 12 Hafta Ek Beceri [15 dk, 2 gün/hf] + 1 Hafta Son Test)
"""

import os
import sys
import openpyxl
import numpy as np

def calculate_mirwald_phv_boys(height_cm, sitting_height_cm, weight_kg, chronological_age):
    """
    Mirwald et al. (2002) - Maturity Offset denklemi (Erkekler için):
    Leg Length = Height - Sitting Height
    Maturity Offset (yıl) = Zirve Boy Hızına (PHV) kalan/geçen süre.
    """
    leg_length_cm = height_cm - sitting_height_cm
    
    # Mirwald 2002 Male Equation
    offset = (
        -9.236
        + 0.0002708 * (leg_length_cm * sitting_height_cm)
        - 0.001663 * (chronological_age * leg_length_cm)
        + 0.007216 * (chronological_age * sitting_height_cm)
        + 0.02292 * (weight_kg / height_cm * 100)
    )
    return offset

def calculate_bmi(weight_kg, height_cm):
    height_m = height_cm / 100.0
    return weight_kg / (height_m ** 2)

def calculate_bilateral_asymmetry(right_grip, left_grip):
    max_grip = max(right_grip, left_grip)
    if max_grip == 0:
        return 0.0
    return (abs(right_grip - left_grip) / max_grip) * 100.0

def calculate_sayers_power(jump_height_cm, weight_kg):
    # Sayers et al. (1999) Peak Anaerobic Power (Watts)
    # Peak Power (W) = 60.7 * Jump Height (cm) + 45.3 * Weight (kg) - 2055
    return 60.7 * jump_height_cm + 45.3 * weight_kg - 2055.0

print("Biyomekanik ve antropometrik doğrulama fonksiyonları hazır.")
