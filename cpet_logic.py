import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, List

class CPETAnalyzer:
    def __init__(self, df_5s: pd.DataFrame, df_bp: pd.DataFrame, df_bxb: pd.DataFrame):
        self.df_5s = df_5s
        self.df_bp = df_bp
        self.df_bxb = df_bxb
        self.results = {}
        
    def preprocess(self):
        for df in [self.df_5s, self.df_bp, self.df_bxb]:
            for col in df.columns:
                if col != 'Time':
                    df[col] = pd.to_numeric(df[col], errors='coerce')
        
        def time_to_sec(t):
            if isinstance(t, str):
                parts = t.split(':')
                if len(parts) == 2: return int(parts[0])*60 + int(parts[1])
                if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
            return t

        self.df_5s['Seconds'] = self.df_5s['Time'].apply(time_to_sec)
        self.df_bp['Seconds'] = self.df_bp['Time'].apply(time_to_sec)
        self.df_bxb['Seconds'] = self.df_bxb['Time'].apply(time_to_sec)
        
    def detect_phases(self):
        peak_idx = self.df_5s['Load'].idxmax()
        self.peak_row = self.df_5s.loc[peak_idx]
        
        # Find recovery start (first load=0 after peak load)
        recovery_df = self.df_5s.loc[peak_idx:]
        recovery_start_idx = recovery_df[recovery_df['Load'] == 0].index
        if len(recovery_start_idx) > 0:
            self.recovery_start_idx = recovery_start_idx[0]
            self.recovery_start_time = self.df_5s.loc[self.recovery_start_idx, 'Seconds']
        else:
            self.recovery_start_idx = self.df_5s.index[-1]
            self.recovery_start_time = self.df_5s.loc[self.recovery_start_idx, 'Seconds']

    def calculate_thresholds(self):
        exercise_df = self.df_5s[(self.df_5s['Load'] > 0) & (self.df_5s.index <= self.peak_row.name)].copy()
        if exercise_df.empty: return
        
        exercise_df = exercise_df.dropna(subset=['V\'O2', 'V\'CO2', 'V\'E'])
        if len(exercise_df) < 10: return

        vo2 = exercise_df['V\'O2'].values
        vco2 = exercise_df['V\'CO2'].values
        
        best_rss = float('inf')
        get_idx = None
        
        start_search = int(len(vo2) * 0.2)
        end_search = int(len(vo2) * 0.8)
        
        for i in range(start_search, end_search):
            x1, y1 = vo2[:i], vco2[:i]
            x2, y2 = vo2[i:], vco2[i:]
            try:
                res1 = stats.linregress(x1, y1)
                res2 = stats.linregress(x2, y2)
                rss = np.sum((y1 - (res1.slope * x1 + res1.intercept))**2) + \
                      np.sum((y2 - (res2.slope * x2 + res2.intercept))**2)
                if rss < best_rss:
                    best_rss = rss
                    get_idx = exercise_df.index[i]
            except: continue
                
        if get_idx is not None:
            self.results['GET_idx'] = get_idx
            self.results['GET_row'] = self.df_5s.loc[get_idx]
        
        try:
            smoothed_eq_vco2 = exercise_df['EqCO2'].rolling(window=5, center=True).mean()
            rcp_idx = smoothed_eq_vco2.idxmin()
            if get_idx is not None and rcp_idx <= get_idx:
                 after_get = smoothed_eq_vco2.loc[get_idx:]
                 if not after_get.empty: rcp_idx = after_get.idxmin()
            self.results['RCP_idx'] = rcp_idx
            self.results['RCP_row'] = self.df_5s.loc[rcp_idx]
        except: pass

    def extract_recovery(self):
        rec_start = self.recovery_start_time
        
        hr_1min = self.df_5s[self.df_5s['Seconds'] >= rec_start + 60]['HR'].first_valid_index()
        self.results['HR_recovery_1min_idx'] = hr_1min
        
        hr_2min = self.df_5s[self.df_5s['Seconds'] >= rec_start + 120]['HR'].first_valid_index()
        self.results['HR_recovery_2min_idx'] = hr_2min
        
        last_bp_idx = self.df_bp.dropna(subset=['Psys', 'Pdia']).index[-1] if not self.df_bp.dropna(subset=['Psys', 'Pdia']).empty else None
        self.results['Recovery_BP_idx'] = last_bp_idx

    def generate_9_panel(self, output_path: str):
        plt.style.use('bmh') # Better baseline style
        fig, axes = plt.subplots(3, 3, figsize=(18, 15))
        plt.subplots_adjust(hspace=0.35, wspace=0.25)
        
        df = self.df_5s.copy()
        # Create smoothed data for lines
        df['VO2_s'] = df['V\'O2'].rolling(5, center=True).mean()
        df['VCO2_s'] = df['V\'CO2'].rolling(5, center=True).mean()
        df['VE_s'] = df['V\'E'].rolling(5, center=True).mean()
        df['HR_s'] = df['HR'].rolling(5, center=True).mean()
        df['EqO2_s'] = df['EqO2'].rolling(5, center=True).mean()
        df['EqCO2_s'] = df['EqCO2'].rolling(5, center=True).mean()
        df['PETO2_s'] = df['PETO2'].rolling(5, center=True).mean()
        df['PETCO2_s'] = df['PETCO2'].rolling(5, center=True).mean()

        get_row = self.results.get('GET_row')
        rcp_row = self.results.get('RCP_row')
        peak_row = self.peak_row
        
        def mark(ax, x_col, y_col):
            if get_row is not None:
                ax.scatter(get_row[x_col], get_row[y_col], color='green', s=100, label='GET', edgecolors='black', zorder=10)
            if rcp_row is not None:
                ax.scatter(rcp_row[x_col], rcp_row[y_col], color='blue', s=100, label='RCP', edgecolors='black', zorder=10)
            ax.scatter(peak_row[x_col], peak_row[y_col], color='red', s=120, marker='X', label='Peak', edgecolors='black', zorder=10)

        # Panel 1: VE vs Time
        axes[0,0].scatter(df['Seconds']/60, df['V\'E'], s=15, color='lightgray', alpha=0.5)
        axes[0,0].plot(df['Seconds']/60, df['VE_s'], color='black', linewidth=2)
        axes[0,0].set_title('1. Ventilation', fontweight='bold')
        axes[0,0].set_ylabel('VE (L/min)')
        axes[0,0].set_xlabel('Time (min)')
        mark(axes[0,0], 'Seconds', 'V\'E') # Note: mark function uses Seconds not min, will fix labels
        
        # Adjust mark for time scale
        def mark_t(ax, x_col, y_col, scale=1):
             if get_row is not None: ax.scatter(get_row[x_col]/scale, get_row[y_col], color='green', s=80, edgecolors='black', label='GET', zorder=10)
             if rcp_row is not None: ax.scatter(rcp_row[x_col]/scale, rcp_row[y_col], color='blue', s=80, edgecolors='black', label='RCP', zorder=10)
             ax.scatter(peak_row[x_col]/scale, peak_row[y_col], color='red', s=100, marker='X', edgecolors='black', label='Peak', zorder=10)

        mark_t(axes[0,0], 'Seconds', 'V\'E', 60)

        # Panel 2: HR & O2 pulse vs Time
        ax2 = axes[0,1].twinx()
        axes[0,1].plot(df['Seconds']/60, df['HR_s'], color='red', label='HR')
        ax2.plot(df['Seconds']/60, df['O2pulse'].rolling(5, center=True).mean(), color='blue', label='O2pulse')
        axes[0,1].set_title('2. HR & O2 Pulse', fontweight='bold')
        axes[0,1].set_ylabel('HR (bpm)', color='red')
        ax2.set_ylabel('O2 Pulse (mL/beat)', color='blue')
        mark_t(axes[0,1], 'Seconds', 'HR', 60)
        
        # Panel 3: VO2 & VCO2 vs Time
        axes[0,2].plot(df['Seconds']/60, df['VO2_s'], color='blue', label='VO2')
        axes[0,2].plot(df['Seconds']/60, df['VCO2_s'], color='red', label='VCO2')
        axes[0,2].set_title('3. Gas Exchange', fontweight='bold')
        axes[0,2].set_ylabel('Gas (mL/min)')
        axes[0,2].legend()
        mark_t(axes[0,2], 'Seconds', 'V\'O2', 60)

        # Panel 4: VE vs VCO2
        axes[1,0].scatter(df['V\'CO2'], df['V\'E'], s=15, color='lightgray')
        axes[1,0].plot(df['VCO2_s'], df['VE_s'], color='black')
        axes[1,0].set_title('4. Ventilatory Efficiency', fontweight='bold')
        axes[1,0].set_xlabel('VCO2 (mL/min)')
        axes[1,0].set_ylabel('VE (L/min)')
        mark_t(axes[1,0], 'V\'CO2', 'V\'E', 1)

        # Panel 5: VCO2 vs VO2 (V-slope)
        axes[1,1].scatter(df['V\'O2'], df['V\'CO2'], s=15, color='lightgray')
        lim = max(df['V\'O2'].max(), df['V\'CO2'].max())
        axes[1,1].plot([0, lim], [0, lim], '--', color='black', alpha=0.4)
        axes[1,1].set_title('5. V-Slope (VCO2 vs VO2)', fontweight='bold')
        axes[1,1].set_xlabel('VO2 (mL/min)')
        axes[1,1].set_ylabel('VCO2 (mL/min)')
        mark_t(axes[1,1], 'V\'O2', 'V\'CO2', 1)

        # Panel 6: EqO2 & EqCO2 vs VO2
        axes[1,2].plot(df['V\'O2'], df['EqO2_s'], color='green', label='VE/VO2')
        axes[1,2].plot(df['V\'O2'], df['EqCO2_s'], color='blue', label='VE/VCO2')
        axes[1,2].set_title('6. Vent. Equivalents', fontweight='bold')
        axes[1,2].set_xlabel('VO2 (mL/min)')
        axes[1,2].legend()
        mark_t(axes[1,2], 'V\'O2', 'EqO2', 1)

        # Panel 7: PETO2 & PETCO2 vs VO2
        axes[2,0].plot(df['V\'O2'], df['PETO2_s'], color='green', label='PETO2')
        axes[2,0].plot(df['V\'O2'], df['PETCO2_s'], color='blue', label='PETCO2')
        axes[2,0].set_title('7. End-Tidal Pressures', fontweight='bold')
        axes[2,0].set_xlabel('VO2 (mL/min)')
        axes[2,0].set_ylabel('Pressure (kPa)')
        axes[2,0].legend()
        mark_t(axes[2,0], 'V\'O2', 'PETO2', 1)

        # Panel 8: RER vs Time
        axes[2,1].plot(df['Seconds']/60, df['RER'].rolling(5, center=True).mean(), color='black')
        axes[2,1].axhline(y=1.0, color='red', linestyle='--', alpha=0.6)
        axes[2,1].set_title('8. RER', fontweight='bold')
        axes[2,1].set_ylabel('RER')
        axes[2,1].set_xlabel('Time (min)')
        mark_t(axes[2,1], 'Seconds', 'RER', 60)

        # Panel 9: VO2/kg & HR vs Load
        axes[2,2].plot(df['Load'], df['VO2_s'], color='blue', label='VO2')
        axes[2,2].set_title('9. Load Response', fontweight='bold')
        axes[2,2].set_xlabel('Load (W)')
        axes[2,2].set_ylabel('VO2 (mL/min)', color='blue')
        mark_t(axes[2,2], 'Load', 'V\'O2', 1)

        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
