import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import os
from scipy.optimize import curve_fit
import scipy.constants as const
from scipy.integrate import trapezoid

import matplotlib
matplotlib.use('TkAgg') 
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.widgets import SpanSelector

# --- 물리학 상수 ---
KB = const.k
Q = const.e
H = const.h
C = const.c
T = 300
VT = (KB * T) / Q 

# =============================================================================
# [MERGED DATA] Solar Spectrum Data (Integrated from solar_spectrum_data.py)
# =============================================================================

# ASTM G173-03 Reference Spectra Derived from SMARTS v. 2.9.2
# Columns: Wavelength (nm), Spectral Irradiance (W*m-2*nm-1)
# Global Tilt (AM 1.5G)
ASTM_G173_AM15G = {
    'nm': np.array([
        280.0, 285.0, 290.0, 295.0, 300.0, 305.0, 310.0, 315.0, 320.0, 325.0, 330.0, 335.0, 340.0, 345.0, 350.0, 355.0, 360.0, 365.0, 370.0, 375.0,
        380.0, 385.0, 390.0, 395.0, 400.0, 405.0, 410.0, 415.0, 420.0, 425.0, 430.0, 435.0, 440.0, 445.0, 450.0, 455.0, 460.0, 465.0, 470.0, 475.0,
        480.0, 485.0, 490.0, 495.0, 500.0, 505.0, 510.0, 515.0, 520.0, 525.0, 530.0, 535.0, 540.0, 545.0, 550.0, 555.0, 560.0, 565.0, 570.0, 575.0,
        580.0, 585.0, 590.0, 595.0, 600.0, 605.0, 610.0, 615.0, 620.0, 625.0, 630.0, 635.0, 640.0, 645.0, 650.0, 655.0, 660.0, 665.0, 670.0, 675.0,
        680.0, 685.0, 690.0, 695.0, 700.0, 705.0, 710.0, 715.0, 720.0, 725.0, 730.0, 735.0, 740.0, 745.0, 750.0, 755.0, 760.0, 765.0, 770.0, 775.0,
        780.0, 785.0, 790.0, 795.0, 800.0, 805.0, 810.0, 815.0, 820.0, 825.0, 830.0, 835.0, 840.0, 845.0, 850.0, 855.0, 860.0, 865.0, 870.0, 875.0,
        880.0, 885.0, 890.0, 895.0, 900.0, 910.0, 920.0, 930.0, 940.0, 950.0, 960.0, 970.0, 980.0, 990.0, 1000.0,
        1020.0, 1040.0, 1060.0, 1080.0, 1100.0, 1120.0, 1140.0, 1160.0, 1180.0, 1200.0,
        1220.0, 1240.0, 1260.0, 1280.0, 1300.0, 1320.0, 1340.0, 1360.0, 1380.0, 1400.0,
        1450.0, 1500.0, 1550.0, 1600.0, 1650.0, 1700.0, 1750.0, 1800.0, 1850.0, 1900.0,
        1950.0, 2000.0, 2100.0, 2200.0, 2300.0, 2400.0, 2500.0, 2600.0, 2700.0, 2800.0, 2900.0, 3000.0,
        3100.0, 3200.0, 3300.0, 3400.0, 3500.0, 3600.0, 3700.0, 3800.0, 3900.0, 4000.0
    ]),
    'W_m2_nm': np.array([
        4.7309e-09, 8.9151e-09, 2.0531e-06, 2.8229e-05, 5.1419e-04, 3.6517e-03, 1.3444e-02, 4.2965e-02, 1.0559e-01, 2.3789e-01, 4.1481e-01, 5.7656e-01, 6.7725e-01, 7.6200e-01, 8.7077e-01, 9.6953e-01, 1.0182e+00, 1.0945e+00, 1.1633e+00, 1.1762e+00,
        1.1166e+00, 1.0963e+00, 1.0962e+00, 1.1895e+00, 1.4298e+00, 1.6373e+00, 1.7456e+00, 1.7371e+00, 1.6961e+00, 1.6111e+00, 1.4842e+00, 1.3556e+00, 1.5430e+00, 1.7651e+00, 1.9056e+00, 1.9793e+00, 1.9866e+00, 1.9686e+00, 1.8906e+00, 1.8596e+00,
        1.9168e+00, 1.8384e+00, 1.7831e+00, 1.8028e+00, 1.7761e+00, 1.7681e+00, 1.7820e+00, 1.7588e+00, 1.7163e+00, 1.7153e+00, 1.7523e+00, 1.7720e+00, 1.7381e+00, 1.7107e+00, 1.7208e+00, 1.7001e+00, 1.6669e+00, 1.6397e+00, 1.6367e+00, 1.6368e+00,
        1.6449e+00, 1.6450e+00, 1.6027e+00, 1.5833e+00, 1.5583e+00, 1.5562e+00, 1.5658e+00, 1.5645e+00, 1.5528e+00, 1.5298e+00, 1.5035e+00, 1.4891e+00, 1.4744e+00, 1.4651e+00, 1.4468e+00, 1.4276e+00, 1.4326e+00, 1.4429e+00, 1.4338e+00, 1.4190e+00,
        1.3891e+00, 1.3533e+00, 1.2588e+00, 1.3323e+00, 1.3644e+00, 1.3582e+00, 1.3320e+00, 1.2829e+00, 1.0547e+00, 1.0772e+00, 1.2599e+00, 1.2562e+00, 1.2570e+00, 1.2403e+00, 1.2292e+00, 1.2185e+00, 0.7712e+00, 0.9463e+00, 1.1592e+00, 1.1724e+00,
        1.1610e+00, 1.1528e+00, 1.1340e+00, 1.1219e+00, 1.0963e+00, 1.0825e+00, 1.0494e+00, 1.0136e+00, 0.9996e+00, 0.9168e+00, 0.7960e+00, 0.9329e+00, 0.9702e+00, 0.9774e+00, 0.9754e+00, 0.9715e+00, 0.9599e+00, 0.9529e+00, 0.9427e+00, 0.9272e+00,
        0.9161e+00, 0.8931e+00, 0.7951e+00, 0.8711e+00, 0.8140e+00, 0.6908e+00, 0.6402e+00, 0.3546e+00, 0.3340e+00, 0.6120e+00, 0.5055e+00, 0.7658e+00, 0.7306e+00, 0.7186e+00, 0.7410e+00,
        0.7153e+00, 0.6868e+00, 0.6558e+00, 0.6231e+00, 0.5375e+00, 0.2848e+00, 0.1747e+00, 0.5186e+00, 0.4468e+00, 0.5284e+00,
        0.4852e+00, 0.4851e+00, 0.4430e+00, 0.4355e+00, 0.0487e+00, 0.0090e+00, 0.0041e+00, 0.0006e+00, 0.0101e+00, 0.0000e+00,
        0.0465e+00, 0.1610e+00, 0.2709e+00, 0.2505e+00, 0.2323e+00, 0.1983e+00, 0.1340e+00, 0.0093e+00, 0.0000e+00, 0.0016e+00,
        0.0210e+00, 0.0886e+00, 0.0963e+00, 0.0768e+00, 0.0638e+00, 0.0381e+00, 0.0135e+00, 0.0000e+00, 0.0000e+00, 0.0000e+00, 0.0000e+00, 0.0084e+00,
        0.0125e+00, 0.0150e+00, 0.0120e+00, 0.0096e+00, 0.0082e+00, 0.0069e+00, 0.0057e+00, 0.0048e+00, 0.0040e+00, 0.0035e+00
    ])
}

# ASTM E490 Standard Zero Air Mass (AM0) Solar Spectral Irradiance
# Units: W/m2/nm
ASTM_E490_AM0 = {
    'nm': np.array([
        119.5, 120.5, 121.5, 122.5, 123.5, 124.5, 125.5, 126.5, 127.5, 128.5, 
        129.5, 130.5, 131.5, 132.5, 133.5, 134.5, 135.5, 136.5, 137.5, 138.5, 
        139.5, 140.5, 141.5, 142.5, 143.5, 144.5, 145.5, 146.5, 147.5, 148.5, 
        149.5, 150.5, 151.5, 152.5, 153.5, 154.5, 155.5, 156.5, 157.5, 158.5, 
        159.5, 160.5, 161.5, 162.5, 163.5, 164.5, 165.5, 166.5, 167.5, 168.5, 
        169.5, 170.5, 171.5, 172.5, 173.5, 174.5, 175.5, 176.5, 177.5, 178.5, 
        179.5, 180.5, 181.5, 182.5, 183.5, 184.5, 185.5, 186.5, 187.5, 188.5, 
        189.5, 190.5, 191.5, 192.5, 193.5, 194.5, 195.5, 196.5, 197.5, 198.5, 
        199.5, 200.0, 210.0, 220.0, 230.0, 240.0, 250.0, 260.0, 270.0,
        280.0, 290.0, 300.0, 310.0, 320.0, 330.0, 340.0, 350.0, 360.0, 370.0, 380.0, 390.0, 400.0, 410.0, 420.0, 430.0, 440.0, 450.0, 460.0, 470.0,
        480.0, 490.0, 500.0, 510.0, 520.0, 530.0, 540.0, 550.0, 560.0, 570.0, 580.0, 590.0, 600.0, 610.0, 620.0, 630.0, 640.0, 650.0, 660.0, 670.0,
        680.0, 690.0, 700.0, 710.0, 720.0, 730.0, 740.0, 750.0, 760.0, 770.0, 780.0, 790.0, 800.0, 810.0, 820.0, 830.0, 840.0, 850.0, 860.0, 870.0,
        880.0, 890.0, 900.0, 925.0, 950.0, 975.0, 1000.0, 1050.0, 1100.0, 1150.0, 1200.0, 1250.0, 1300.0, 1350.0, 1400.0,
        1450.0, 1500.0, 1550.0, 1600.0, 1650.0, 1700.0, 1750.0, 1800.0, 1850.0, 1900.0, 1950.0, 2000.0, 2100.0, 2200.0, 2300.0, 2400.0, 2500.0,
        2600.0, 2700.0, 2800.0, 2900.0, 3000.0, 3200.0, 3400.0, 3600.0, 3800.0, 4000.0
    ]),
    'W_m2_nm': np.array([
        0.063, 0.465, 5.750, 0.443, 0.040, 0.057, 0.081, 0.069, 0.054, 0.051,
        0.045, 0.043, 0.041, 0.038, 0.039, 0.039, 0.038, 0.038, 0.040, 0.039,
        0.039, 0.038, 0.038, 0.041, 0.040, 0.043, 0.052, 0.055, 0.058, 0.059,
        0.061, 0.065, 0.076, 0.098, 0.117, 0.115, 0.091, 0.076, 0.069, 0.066,
        0.062, 0.054, 0.049, 0.043, 0.038, 0.034, 0.033, 0.033, 0.041, 0.049,
        0.054, 0.063, 0.068, 0.069, 0.057, 0.048, 0.046, 0.060, 0.063, 0.063,
        0.062, 0.054, 0.045, 0.035, 0.026, 0.021, 0.016, 0.012, 0.009, 0.007,
        0.006, 0.005, 0.004, 0.003, 0.003, 0.002, 0.002, 0.002, 0.001, 0.001,
        0.001, 0.001, 0.002, 0.006, 0.007, 0.006, 0.010, 0.032, 0.092,
        0.252, 0.518, 0.514, 0.686, 0.840, 1.096, 1.157, 1.341, 1.350, 1.549, 
        1.642, 1.761, 2.149, 2.378, 2.374, 2.222, 2.373, 2.522, 2.502, 2.450, 
        2.433, 2.304, 2.306, 2.301, 2.155, 2.132, 2.138, 2.100, 2.023, 1.996, 
        2.015, 1.916, 1.889, 1.865, 1.815, 1.758, 1.737, 1.638, 1.603, 1.579, 
        1.536, 1.488, 1.450, 1.407, 1.352, 1.309, 1.272, 1.235, 1.200, 1.168, 
        1.137, 1.107, 1.081, 1.053, 1.026, 1.000, 0.975, 0.950, 0.926, 0.902, 
        0.879, 0.857, 0.835, 0.785, 0.738, 0.697, 0.654, 0.589, 0.525, 0.470, 0.422, 0.377, 0.339, 0.301, 0.267,
        0.238, 0.211, 0.187, 0.165, 0.147, 0.129, 0.113, 0.099, 0.089, 0.079, 0.069, 0.061, 0.050, 0.043, 0.036, 0.033, 0.027,
        0.022, 0.020, 0.017, 0.015, 0.013, 0.010, 0.008, 0.006, 0.005, 0.004
    ])
}

# =============================================================================
# [APP LOGIC] Solar PLQY Calculator
# =============================================================================

SOLAR_SPECTRA = {
    'am15': {'nm': ASTM_G173_AM15G['nm'], 'irr': ASTM_G173_AM15G['W_m2_nm']},
    'am0':  {'nm': ASTM_E490_AM0['nm'], 'irr': ASTM_E490_AM0['W_m2_nm']}
}

# [AUTO-SCALING] Ensure exact standard integrated power
# ASTM G173 AM1.5G target: 1000 W/m2
# ASTM E490 AM0 target: 1366.1 W/m2

_integ_am15 = trapezoid(SOLAR_SPECTRA['am15']['irr'], SOLAR_SPECTRA['am15']['nm'])
_integ_am0 = trapezoid(SOLAR_SPECTRA['am0']['irr'], SOLAR_SPECTRA['am0']['nm'])

SCALE_FACTOR_AM15 = 1000.0 / _integ_am15 if _integ_am15 > 0 else 1.0
SCALE_FACTOR_AM0 = 1366.1 / _integ_am0 if _integ_am0 > 0 else 1.0

# Apply correction
SOLAR_SPECTRA['am15']['irr'] *= SCALE_FACTOR_AM15
SOLAR_SPECTRA['am0']['irr'] *= SCALE_FACTOR_AM0


def get_solar_photon_flux_density(wavelength_nm, spectrum_type='am15'):
    """
    Photon Flux Density [photons/s/m2/nm] 반환
    """
    w_nm = np.array(wavelength_nm)
    
    # 1. 보간 (Interpolation) - Scaled Data 사용
    # Use the appropriate spectrum key
    ref_nm = SOLAR_SPECTRA[spectrum_type]['nm']
    ref_irrad = SOLAR_SPECTRA[spectrum_type]['irr']
    irradiance = np.interp(w_nm, ref_nm, ref_irrad, left=0, right=0)
    
    # 2. 광자 에너지 (J)
    w_m = w_nm * 1e-9
    energy_per_photon = (H * C) / w_m
    
    # 3. Flux 변환
    photon_flux = irradiance / energy_per_photon
    return photon_flux

def calculate_solar_metrics_dual(bandgap_ev, eqe_ratio, plqy_ratio):
    lambda_g = (H * C) / (bandgap_ev * Q) * 1e9
    
    # --- J0,rad (Blackbody 300K) ---
    w_range = np.linspace(200, lambda_g, 2000)
    w_m = w_range * 1e-9
    p1 = 2 * H * C**2
    p2_const = (H * C) / (KB * 300)
    p2 = np.clip(p2_const / w_m, 0, 700)
    bb_irrad_300 = (p1 / (w_m**5)) / (np.exp(p2) - 1) * np.pi
    bb_flux_300 = bb_irrad_300 / ((H * C) / w_m) * 1e-9 
    j0_rad_A = Q * trapezoid(bb_flux_300, w_range)
    if j0_rad_A == 0: j0_rad_A = 1e-25
    
    # --- Calculation ---
    def calc_for_mode(mode_name, pin_mW_cm2):
        if lambda_g < 300: 
            return {'Jsc': 0, 'iVoc': 0, 'iFF': 0, 'iPCE': 0}
            
        # 적분 구간: Data Range Start ~ Bandgap
        start_nm = 200 if mode_name == 'am0' else 280
        calc_waves = np.linspace(start_nm, lambda_g, 2000)
        flux = get_solar_photon_flux_density(calc_waves, mode_name)
        
        # Jsc = q * Integral(Flux) * EQE
        total_photons = trapezoid(flux, calc_waves)
        jsc_ideal_A = total_photons * Q
        jsc_real_A = jsc_ideal_A * eqe_ratio
        
        # mA/cm2 변환
        jsc_mA = jsc_real_A * 0.1 
        
        # iVoc
        voc_rad = VT * np.log(jsc_real_A / j0_rad_A + 1)
        eff_plqy = plqy_ratio if plqy_ratio > 1e-9 else 1e-9
        ivoc = voc_rad + VT * np.log(eff_plqy)
        
        # iFF
        v_oc_norm = ivoc / VT
        if v_oc_norm <= 0: iff = 0
        else: iff = (v_oc_norm - np.log(v_oc_norm + 0.72)) / (v_oc_norm + 1)
        
        # iPCE
        ipce = (ivoc * jsc_mA * iff) / pin_mW_cm2 * 100
        
        return {
            'Jsc': jsc_mA, 'iVoc': ivoc, 'iFF': iff*100, 'iPCE': ipce
        }

    # Calculate Standard Total Irradiance from Spectrum (mW/cm2)
    # AM1.5G should be 100.0, AM0 should be 136.61
    total_irrad_am15 = trapezoid(SOLAR_SPECTRA['am15']['irr'], SOLAR_SPECTRA['am15']['nm']) / 10.0
    total_irrad_am0  = trapezoid(SOLAR_SPECTRA['am0']['irr'], SOLAR_SPECTRA['am0']['nm']) / 10.0

    am15 = calc_for_mode('am15', total_irrad_am15) 
    am0 = calc_for_mode('am0', total_irrad_am0)
    
    return {'AM1.5': am15, 'AM0': am0, 'Eg': bandgap_ev}

def gaussian(x, amp, cen, wid):
    return amp * np.exp(-(x - cen)**2 / (2 * wid**2))

class SolarPLQYApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Solar PLQY Analyzer (Embedded Spectrum)")
        self.root.geometry("1600x950")
        
        self.data_store = {} 
        self.current_file_key = None
        
        self.var_l_start = tk.DoubleVar()
        self.var_l_end = tk.DoubleVar()
        self.var_e_start = tk.DoubleVar()
        self.var_e_end = tk.DoubleVar()
        
        self.var_bg_start = tk.DoubleVar(value=850.0)
        self.var_bg_end = tk.DoubleVar(value=900.0)
        self.use_bg_correction = tk.BooleanVar(value=False)
        self.use_gaussian_fit = tk.BooleanVar(value=False)
        self.view_mode = tk.StringVar(value="net")
        
        self.var_target_eqe = tk.DoubleVar(value=100.0) 
        self.var_detected_eg = tk.StringVar(value="-")
        
        self.setup_ui()
        
    def setup_ui(self):
        main_pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)
        
        left_frame = tk.Frame(main_pane, width=280, bg="#f5f5f5")
        main_pane.add(left_frame)
        
        btn_frame = tk.Frame(left_frame, bg="#f5f5f5")
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Button(btn_frame, text="📂 Add Files", command=self.load_files, bg="#e1e1e1").pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="Clear", command=self.clear_all, bg="#ffcccc").pack(side=tk.LEFT, padx=2)

        self.tree = ttk.Treeview(left_frame, columns=("File", "PLQY", "AM1.5"), show='headings')
        self.tree.heading("File", text="File Name")
        self.tree.heading("PLQY", text="PLQY(%)")
        self.tree.heading("AM1.5", text="AM1.5(%)")
        self.tree.column("File", width=120)
        self.tree.column("PLQY", width=60, anchor="center")
        self.tree.column("AM1.5", width=60, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree.bind("<<TreeviewSelect>>", self.on_file_select)
        
        center_frame = tk.Frame(main_pane, bg="white")
        main_pane.add(center_frame, stretch="always")
        
        ctrl_frame = tk.Frame(center_frame, bd=1, relief=tk.RAISED, bg="#eeeeee")
        ctrl_frame.pack(fill=tk.X)
        
        f1 = tk.Frame(ctrl_frame, bg="#eeeeee")
        f1.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        tk.Label(f1, text="Laser:", fg="red", bg="#eeeeee").pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.var_l_start, width=5).pack(side=tk.LEFT)
        tk.Label(f1, text="~", bg="#eeeeee").pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.var_l_end, width=5).pack(side=tk.LEFT)
        
        tk.Label(f1, text="  Emission:", fg="green", bg="#eeeeee").pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.var_e_start, width=5).pack(side=tk.LEFT)
        tk.Label(f1, text="~", bg="#eeeeee").pack(side=tk.LEFT)
        tk.Entry(f1, textvariable=self.var_e_end, width=5).pack(side=tk.LEFT)
        
        tk.Button(f1, text="Update", command=self.update_current, bg="lightblue").pack(side=tk.LEFT, padx=10)
        tk.Button(f1, text="⚡ Apply All", command=self.apply_to_all, bg="#ffd700", font=("bold")).pack(side=tk.LEFT, padx=5)
        
        f2 = tk.Frame(ctrl_frame, bg="#e6f2ff")
        f2.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        tk.Checkbutton(f2, text="BG Correct", variable=self.use_bg_correction, command=self.update_current, bg="#e6f2ff").pack(side=tk.LEFT)
        tk.Entry(f2, textvariable=self.var_bg_start, width=5).pack(side=tk.LEFT)
        tk.Label(f2, text="~", bg="#e6f2ff").pack(side=tk.LEFT)
        tk.Entry(f2, textvariable=self.var_bg_end, width=5).pack(side=tk.LEFT)
        
        tk.Label(f2, text=" | ", bg="#e6f2ff").pack(side=tk.LEFT)
        tk.Checkbutton(f2, text="Gaussian Fit", variable=self.use_gaussian_fit, command=self.update_current, bg="#e6f2ff", font=("bold")).pack(side=tk.LEFT)
        
        tk.Label(f2, text=" | View:", bg="#e6f2ff").pack(side=tk.LEFT)
        tk.Radiobutton(f2, text="Raw", variable=self.view_mode, value="raw", command=self.redraw_plots, bg="#e6f2ff").pack(side=tk.LEFT)
        tk.Radiobutton(f2, text="Net", variable=self.view_mode, value="net", command=self.redraw_plots, bg="#e6f2ff").pack(side=tk.LEFT)

        self.fig = Figure(figsize=(8, 5), dpi=100)
        self.ax1 = self.fig.add_subplot(121)
        self.ax2 = self.fig.add_subplot(122)
        self.canvas = FigureCanvasTkAgg(self.fig, master=center_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        toolbar = NavigationToolbar2Tk(self.canvas, center_frame)
        toolbar.update()
        
        right_frame = tk.Frame(main_pane, width=380, bg="#FFF8DC", bd=2, relief=tk.GROOVE)
        main_pane.add(right_frame)
        
        tk.Label(right_frame, text="Solar Cell Prediction", font=("Arial", 12, "bold"), bg="#FFF8DC", fg="#8B4500").pack(pady=10)
        
        f_pred = tk.Frame(right_frame, bg="#FFF8DC")
        f_pred.pack(fill=tk.X, padx=10)
        
        tk.Label(f_pred, text="Detected Bandgap (Eg):", bg="#FFF8DC", anchor="w").grid(row=0, column=0, sticky="w")
        tk.Label(f_pred, textvariable=self.var_detected_eg, bg="#FFF8DC", font=("bold")).grid(row=0, column=1, sticky="w")
        
        tk.Label(f_pred, text="Expected EQE (%):", bg="#FFF8DC", anchor="w").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(f_pred, textvariable=self.var_target_eqe, width=8).grid(row=1, column=1, sticky="w")
        
        tk.Button(right_frame, text="Calculate iPCE", command=self.calculate_solar_metrics, bg="#FFA500", fg="white", font=("bold")).pack(pady=10)
        
        res_container = tk.Frame(right_frame, bg="white")
        res_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        f_am15 = tk.LabelFrame(res_container, text="AM 1.5G (1 Sun)", bg="white", fg="blue", font=("bold"))
        f_am15.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
        self.lbl_am15_jsc = tk.Label(f_am15, text="Jsc:\n-", bg="white"); self.lbl_am15_jsc.pack(pady=2)
        self.lbl_am15_ivoc = tk.Label(f_am15, text="iVoc:\n-", bg="white"); self.lbl_am15_ivoc.pack(pady=2)
        self.lbl_am15_iff = tk.Label(f_am15, text="iFF:\n-", bg="white"); self.lbl_am15_iff.pack(pady=2)
        self.lbl_am15_ipce = tk.Label(f_am15, text="iPCE:\n- %", bg="white", font=("bold"), fg="red"); self.lbl_am15_ipce.pack(pady=5)

        f_am0 = tk.LabelFrame(res_container, text="AM 0 (Space)", bg="white", fg="purple", font=("bold"))
        f_am0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
        self.lbl_am0_jsc = tk.Label(f_am0, text="Jsc:\n-", bg="white"); self.lbl_am0_jsc.pack(pady=2)
        self.lbl_am0_ivoc = tk.Label(f_am0, text="iVoc:\n-", bg="white"); self.lbl_am0_ivoc.pack(pady=2)
        self.lbl_am0_iff = tk.Label(f_am0, text="iFF:\n-", bg="white"); self.lbl_am0_iff.pack(pady=2)
        self.lbl_am0_ipce = tk.Label(f_am0, text="iPCE:\n- %", bg="white", font=("bold"), fg="red"); self.lbl_am0_ipce.pack(pady=5)


    def load_files(self):
        filepaths = filedialog.askopenfilenames()
        if not filepaths: return
        for fpath in filepaths:
            fname = os.path.basename(fpath)
            if fname in self.data_store: continue
            try:
                if fpath.endswith('.csv'): df = pd.read_csv(fpath)
                else: df = pd.read_excel(fpath)
                cols_low = [c.lower() for c in df.columns]
                w_idx = next(i for i, c in enumerate(cols_low) if "wave" in c or "nm" in c)
                r_idx = next(i for i, c in enumerate(cols_low) if ("spec" in c and "1" in c) or "ref" in c)
                s_idx = next(i for i, c in enumerate(cols_low) if ("spec" in c and "2" in c) or "sam" in c)
                
                wave = df.iloc[:, w_idx].values
                ref = df.iloc[:, r_idx].values
                sam = df.iloc[:, s_idx].values
                
                peak_idx = np.argmax(ref)
                peak_w = wave[peak_idx]
                
                self.data_store[fname] = {
                    'wave': wave, 'ref': ref, 'sam': sam,
                    'l_range': [peak_w-15, peak_w+15],
                    'e_range': [peak_w+50, wave[-1]],
                    'plqy': 0.0, 'fit_params': None, 'ipce_data': None
                }
                self.calculate_plqy(fname)
                self.tree.insert("", "end", iid=fname, values=(fname, f"{self.data_store[fname]['plqy']:.2f}", "-"))
                
            except Exception as e:
                print(e)
                
        if self.tree.get_children():
            self.tree.selection_set(self.tree.get_children()[0])
            self.on_file_select(None)

    def calculate_plqy(self, fname):
        d = self.data_store[fname]
        wave, ref, sam = d['wave'], d['ref'], d['sam']
        l_min, l_max = d['l_range']
        mask_l = (wave >= l_min) & (wave <= l_max)
        Absorbed = trapezoid(ref[mask_l]-sam[mask_l], wave[mask_l])
        
        e_min, e_max = d['e_range']
        net = sam - ref
        
        if self.use_bg_correction.get():
            bg_s, bg_e = self.var_bg_start.get(), self.var_bg_end.get()
            mask_bg = (wave >= bg_s) & (wave <= bg_e)
            offset = np.mean(net[mask_bg]) if np.any(mask_bg) else 0
        else: offset = 0
            
        d['offset'] = offset
        corr_sig = net - offset
        mask_e = (wave >= e_min) & (wave <= e_max)
        
        if self.use_gaussian_fit.get():
            x = wave[mask_e]
            y = corr_sig[mask_e]
            try:
                p0 = [np.max(y), x[np.argmax(y)], 20]
                popt, _ = curve_fit(gaussian, x, y, p0=p0, maxfev=5000)
                d['fit_params'] = popt
                Emitted = trapezoid(gaussian(x, *popt), x)
            except:
                d['fit_params'] = None
                Emitted = trapezoid(corr_sig[mask_e], wave[mask_e])
        else:
            d['fit_params'] = None
            Emitted = trapezoid(corr_sig[mask_e], wave[mask_e])
            
        plqy = (Emitted / Absorbed * 100) if Absorbed > 0 else 0
        d['plqy'] = plqy
        return plqy

    def calculate_solar_metrics(self):
        if not self.current_file_key: return
        fname = self.current_file_key
        d = self.data_store[fname]
        
        if d['fit_params'] is not None:
            peak_nm = d['fit_params'][1]
        else:
            e_min, e_max = d['e_range']
            mask = (d['wave'] >= e_min) & (d['wave'] <= e_max)
            if np.any(mask):
                peak_idx = np.argmax(d['sam'][mask] - d['ref'][mask])
                peak_nm = d['wave'][mask][peak_idx]
            else: peak_nm = 750
            
        eg_ev = 1240.0 / peak_nm
        self.var_detected_eg.set(f"{eg_ev:.3f} eV ({peak_nm:.1f} nm)")
        
        eqe_val = self.var_target_eqe.get()
        plqy_val = d['plqy']
        
        try:
            res = calculate_solar_metrics_dual(eg_ev, eqe_val/100, plqy_val/100.0)
            d['ipce_data'] = res
            
            am15 = res['AM1.5']
            self.lbl_am15_jsc.config(text=f"Jsc:\n{am15['Jsc']:.2f}")
            self.lbl_am15_ivoc.config(text=f"iVoc:\n{am15['iVoc']:.3f}")
            self.lbl_am15_iff.config(text=f"iFF:\n{am15['iFF']:.1f}")
            self.lbl_am15_ipce.config(text=f"iPCE:\n{am15['iPCE']:.2f} %")

            am0 = res['AM0']
            self.lbl_am0_jsc.config(text=f"Jsc:\n{am0['Jsc']:.2f}")
            self.lbl_am0_ivoc.config(text=f"iVoc:\n{am0['iVoc']:.3f}")
            self.lbl_am0_iff.config(text=f"iFF:\n{am0['iFF']:.1f}")
            self.lbl_am0_ipce.config(text=f"iPCE:\n{am0['iPCE']:.2f} %")
            
            self.tree.set(fname, "AM1.5", f"{am15['iPCE']:.2f}")
        except Exception as e:
            messagebox.showerror("Calc Error", str(e))

    def update_current(self):
        if self.current_file_key:
            self.data_store[self.current_file_key]['l_range'] = [self.var_l_start.get(), self.var_l_end.get()]
            self.data_store[self.current_file_key]['e_range'] = [self.var_e_start.get(), self.var_e_end.get()]
            plqy = self.calculate_plqy(self.current_file_key)
            self.tree.set(self.current_file_key, "PLQY", f"{plqy:.2f}")
            self.redraw_plots()
            if self.data_store[self.current_file_key].get('ipce_data'):
                self.calculate_solar_metrics()

    def apply_to_all(self):
        new_l = [self.var_l_start.get(), self.var_l_end.get()]
        new_e = [self.var_e_start.get(), self.var_e_end.get()]
        
        for fname in self.data_store:
            self.data_store[fname]['l_range'] = new_l
            self.data_store[fname]['e_range'] = new_e
            plqy = self.calculate_plqy(fname)
            self.tree.set(fname, "PLQY", f"{plqy:.2f}")
            
        if self.current_file_key:
            self.redraw_plots()
        messagebox.showinfo("Done", "Settings applied to all files.")

    def on_file_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        fname = sel[0]
        self.current_file_key = fname
        d = self.data_store[fname]
        
        self.var_l_start.set(d['l_range'][0])
        self.var_l_end.set(d['l_range'][1])
        self.var_e_start.set(d['e_range'][0])
        self.var_e_end.set(d['e_range'][1])
        
        self.redraw_plots()
        
        if d.get('ipce_data'):
            res = d['ipce_data']
            self.var_detected_eg.set(f"{res['Eg']:.3f} eV")
            am15 = res['AM1.5']
            self.lbl_am15_jsc.config(text=f"Jsc:\n{am15['Jsc']:.2f}")
            self.lbl_am15_ivoc.config(text=f"iVoc:\n{am15['iVoc']:.3f}")
            self.lbl_am15_iff.config(text=f"iFF:\n{am15['iFF']:.1f}")
            self.lbl_am15_ipce.config(text=f"iPCE:\n{am15['iPCE']:.2f} %")

            am0 = res['AM0']
            self.lbl_am0_jsc.config(text=f"Jsc:\n{am0['Jsc']:.2f}")
            self.lbl_am0_ivoc.config(text=f"iVoc:\n{am0['iVoc']:.3f}")
            self.lbl_am0_iff.config(text=f"iFF:\n{am0['iFF']:.1f}")
            self.lbl_am0_ipce.config(text=f"iPCE:\n{am0['iPCE']:.2f} %")
        else:
            for lbl in [self.lbl_am15_jsc, self.lbl_am15_ivoc, self.lbl_am15_iff, self.lbl_am15_ipce,
                        self.lbl_am0_jsc, self.lbl_am0_ivoc, self.lbl_am0_iff, self.lbl_am0_ipce]:
                lbl.config(text="-\n-")

    def redraw_plots(self):
        if not self.current_file_key: return
        fname = self.current_file_key
        d = self.data_store[fname]
        
        self.ax1.clear(); self.ax2.clear()
        
        self.ax1.plot(d['wave'], d['ref'], 'gray', alpha=0.5)
        self.ax1.plot(d['wave'], d['sam'], 'k')
        self.ax1.set_title("Laser")
        self.ax1.set_xlim(np.mean(d['l_range'])-20, np.mean(d['l_range'])+20)
        self.span1 = SpanSelector(self.ax1, lambda m,x: [self.var_l_start.set(m), self.var_l_end.set(x), self.update_current()], 'horizontal', props=dict(facecolor='red', alpha=0.2), interactive=True)
        self.span1.extents = tuple(d['l_range'])
        
        self.ax2.set_title("Emission")
        mask_view = (d['wave'] > d['e_range'][0]-50) & (d['wave'] < d['e_range'][1]+50)
        
        if self.view_mode.get() == "raw":
            self.ax2.plot(d['wave'], d['ref'], 'gray', alpha=0.5)
            self.ax2.plot(d['wave'], d['sam'], 'k')
            if np.any(mask_view):
                ymax = np.max(d['sam'][mask_view])
                self.ax2.set_ylim(-ymax*0.05, ymax * 1.3)
        else:
            net = (d['sam'] - d['ref']) - d['offset']
            self.ax2.plot(d['wave'], net, 'g', alpha=0.3, label='Net')
            
            if d['fit_params'] is not None:
                x_fit = np.linspace(d['e_range'][0], d['e_range'][1], 500)
                self.ax2.plot(x_fit, gaussian(x_fit, *d['fit_params']), 'm', lw=2, label='Fit')
            
            if np.any(mask_view):
                ymax = np.max(net[mask_view])
                self.ax2.set_ylim(-ymax*0.1, ymax * 1.3)
            
            self.ax2.legend()
            
        self.ax2.set_xlim(d['e_range'][0]-20, d['e_range'][1]+20)
        self.span2 = SpanSelector(self.ax2, lambda m,x: [self.var_e_start.set(m), self.var_e_end.set(x), self.update_current()], 'horizontal', props=dict(facecolor='green', alpha=0.2), interactive=True)
        self.span2.extents = tuple(d['e_range'])
        
        self.canvas.draw()
        
    def clear_all(self):
        self.data_store.clear()
        self.tree.delete(*self.tree.get_children())
        self.ax1.clear(); self.ax2.clear(); self.canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = SolarPLQYApp(root)
    root.mainloop()