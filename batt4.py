import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

# =====================
# CONFIG
# =====================
EXCEL_PATH = '210350328803 202505 - 202506 (1).xlsx'   # source Excel
SHEET_NAME = 'Sheet'

# Aggregate May & June
YEAR_MONTHS = [(2025, 5), (2025, 6)]

# 14:00–22:00 half-hour intervals
INTERVAL_HOURS = 0.5

# Battery unit parameters (per 233 kWh unit)
UNIT_KWH = 233.0
# Reserve removed per request: we size on NAMEPLATE energy only
MAX_POWER_KW_PER_UNIT = 105.0      # per 233 kWh unit discharge cap (kW)

# Threshold sweep settings
SWEEP_START_KW = 1100.0
SWEEP_END_KW   = 695.0             # include this and also the first step below it
SWEEP_STEP_KW  = -25.0

# Economics (for Payback)
CAPEX_PER_KWH = 1350.0             # RM/kWh
MD_TARIFF_RM_PER_KW = 97.06        # RM per kW of MD
BILLING_PERIODS_PER_YEAR = 12       # months

# Outputs
SWEEP_CSV = 'ThresholdSweep_MayJun_stats.csv'
VISUALIZATIONS_PDF = 'ThresholdSweep_MayJun_visualizations.pdf'


# =====================
# Data loading
# =====================
def load_month(excel_path: str, sheet_name: str, year: int, month: int):
    """Load one month and return list of (date, demand_kw_array, time_labels) for 14:00–22:00."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name, parse_dates=['start_time'])
    df = df[['start_time', 'kw_import']].copy()

    mask = (df['start_time'].dt.year == year) & (df['start_time'].dt.month == month)
    dfm = df.loc[mask].copy()
    if dfm.empty:
        raise RuntimeError(f'No rows for {year}-{month:02d} in {excel_path}')

    days = []
    for day, g in dfm.sort_values('start_time').groupby(dfm['start_time'].dt.date):
        # 14:00–22:00 (inclusive start, exclusive end)
        g2 = g[(g['start_time'].dt.hour >= 14) & (g['start_time'].dt.hour < 22)]
        if g2.empty:
            continue
        demand_kw = g2['kw_import'].astype(float).to_numpy()
        labels = g2['start_time'].dt.strftime('%H:%M').tolist()
        days.append((day, demand_kw, labels))

    if not days:
        raise RuntimeError('All days empty after applying 14:00–22:00 filter.')
    return days


def load_months(excel_path: str, sheet_name: str, year_months):
    """Load multiple months and concatenate their day-series."""
    all_days = []
    for (y, m) in year_months:
        all_days.extend(load_month(excel_path, sheet_name, y, m))
    all_days.sort(key=lambda x: x[0])  # by date
    return all_days


# =====================
# Threshold sweep (Highest-day aggregation across all days)
# =====================
def build_thresholds(start_kw: float, end_kw: float, step_kw: float):
    """1100 down by -25 to 695, and include the first step below 695 (i.e., 675)."""
    if step_kw == 0:
        raise ValueError('SWEEP_STEP_KW cannot be zero')
    Hs = []
    H = start_kw
    if step_kw < 0:
        while True:
            Hs.append(H)
            nxt = H + step_kw
            if nxt < end_kw:
                Hs.append(nxt)  # include first step below the lower bound
                break
            H = nxt
    else:
        while True:
            Hs.append(H)
            nxt = H + step_kw
            if nxt > end_kw:
                Hs.append(nxt)
                break
            H = nxt
    return Hs


def compute_threshold_sweep_stats_highest(days,
                                          dt_h: float = INTERVAL_HOURS,
                                          start_kw: float = SWEEP_START_KW,
                                          end_kw: float = SWEEP_END_KW,
                                          step_kw: float = SWEEP_STEP_KW):
    """
    For each threshold H, aggregate across *all* days (May + June, 14:00–22:00) and compute:
      - Highest_Energy_kWh: max over days of Σ max(0, D - H) * dt
      - Highest_Energy_Day: which day attained that highest energy
      - Highest_Peak_Shaved_kW: max over days of max_t max(0, D - H)
      - Highest_Peak_Day: which day attained that highest shaved power
    Then compute minimal units needed (nameplate energy only, 233 kWh per unit; 105 kW per unit),
    fit flags up to 4 units, a Limiting_Factor label, and the requested Payback/Efficiency.
    """
    thresholds = build_thresholds(start_kw, end_kw, step_kw)

    e_unit_nom = UNIT_KWH
    p_unit_max = MAX_POWER_KW_PER_UNIT if MAX_POWER_KW_PER_UNIT is not None else float('inf')

    # Cache daily arrays
    day_series = []
    for (day, demand_kw, _labels) in days:
        D = demand_kw.astype(float)
        day_series.append((pd.Timestamp(day).date(), D))

    rows = []
    for H in thresholds:
        highest_energy = 0.0
        highest_energy_day = None

        highest_peak = 0.0
        highest_peak_day = None

        for day, D in day_series:
            shave = np.maximum(D - H, 0.0)
            E = float(np.sum(shave) * dt_h)                # kWh
            P = float(np.max(shave)) if shave.size else 0  # kW

            if E > highest_energy + 1e-12:
                highest_energy = E
                highest_energy_day = day
            if P > highest_peak + 1e-12:
                highest_peak = P
                highest_peak_day = day

        # Minimal units by power & nameplate energy (no reserve)
        units_power  = int(np.ceil(highest_peak   / p_unit_max)) if np.isfinite(p_unit_max) else 0
        units_energy = int(np.ceil(highest_energy / e_unit_nom)) if e_unit_nom > 0 else 0
        units_needed = max(units_power, units_energy)

        # For 'no shaving' cases, set 0 units; otherwise at least 1
        if units_needed == 0 and (highest_peak > 0 or highest_energy > 0):
            units_needed = 1

        # Fits up to 4 units (using nameplate energy + per-unit power)
        fits_units = {}
        for u in range(1, 5):
            stack_power  = u * p_unit_max
            stack_energy = u * e_unit_nom
            fits_units[u] = (highest_peak <= stack_power) and (highest_energy <= stack_energy)

        # Limiting factor label
        if highest_peak == 0 and highest_energy == 0:
            limiting = 'No shaving'
        elif units_power > units_energy:
            limiting = 'Power-limited'
        elif units_energy > units_power:
            limiting = 'Energy-limited'
        else:
            limiting = 'Both energy and power'

        # Battery capacity sized by minimal units (nameplate)
        min_capacity_kwh = int(units_needed * UNIT_KWH)

        # ---- New: Payback & Efficiency ----
        # MD delta = Highest_Peak_Shaved_kW (geometric exceedance above H)
        md_delta_kw = highest_peak
        if md_delta_kw > 0 and min_capacity_kwh > 0:
            payback_years = (CAPEX_PER_KWH * min_capacity_kwh) / (md_delta_kw * MD_TARIFF_RM_PER_KW * BILLING_PERIODS_PER_YEAR)
            efficiency = md_delta_kw / float(min_capacity_kwh)
        else:
            payback_years = np.nan
            efficiency = np.nan

        rows.append({
            'Threshold_kW': round(H, 3),
            'Highest_Energy_kWh': round(highest_energy, 3),
            'Highest_Energy_Day': highest_energy_day,
            'Highest_Peak_Shaved_kW': round(highest_peak, 3),
            'Highest_Peak_Day': highest_peak_day,
            'Min_Units_Required': int(units_needed),
            'Min_Capacity_kWh' : int(min_capacity_kwh),
            'Limiting_Factor': limiting,
            'Payback_years': round(payback_years, 3) if pd.notna(payback_years) else None,
            'Efficiency': round(efficiency, 6) if pd.notna(efficiency) else None,
            'Fits_1x233': bool(fits_units[1]),
            'Fits_2x233(466)': bool(fits_units[2]),
            'Fits_3x233(699)': bool(fits_units[3]),
            'Fits_4x233(932)': bool(fits_units[4]),
        })

    return rows


def create_visualizations(df):
    """Create comprehensive visualizations of the battery analysis results."""
    
    # Filter out rows with NaN values for visualization
    df_viz = df.dropna(subset=['Payback_years', 'Efficiency'])
    
    with PdfPages(VISUALIZATIONS_PDF) as pdf:
        
        # Page 1: Key Analysis Charts
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('Battery Storage Economics Analysis - May & June 2025', fontsize=16, fontweight='bold')
        
        # 1. Payback Period vs Threshold
        ax1 = axes[0, 0]
        ax1.plot(df_viz['Threshold_kW'], df_viz['Payback_years'], 'b-', linewidth=2, marker='o', markersize=4)
        ax1.set_xlabel('Threshold (kW)', fontsize=11)
        ax1.set_ylabel('Payback Period (years)', fontsize=11)
        ax1.set_title('Payback Period vs Demand Threshold', fontsize=12, fontweight='bold')
        ax1.grid(True, alpha=0.3)
        
        # 2. Efficiency vs Threshold
        ax2 = axes[0, 1]
        ax2.plot(df_viz['Threshold_kW'], df_viz['Efficiency'], 'g-', linewidth=2, marker='s', markersize=4)
        ax2.set_xlabel('Threshold (kW)', fontsize=11)
        ax2.set_ylabel('Efficiency (kW saved/kWh capacity)', fontsize=11)
        ax2.set_title('System Efficiency vs Demand Threshold', fontsize=12, fontweight='bold')
        ax2.grid(True, alpha=0.3)
        
        # 3. Units Required Distribution
        ax3 = axes[1, 0]
        units_counts = df_viz['Min_Units_Required'].value_counts().sort_index()
        colors = ['#2ecc71', '#3498db', '#9b59b6', '#e74c3c', '#f39c12']
        bars = ax3.bar(units_counts.index, units_counts.values, color=colors[:len(units_counts)])
        ax3.set_xlabel('Number of 233 kWh Units Required', fontsize=11)
        ax3.set_ylabel('Count of Threshold Scenarios', fontsize=11)
        ax3.set_title('Distribution of Battery Units Required', fontsize=12, fontweight='bold')
        ax3.set_xticks(units_counts.index)
        ax3.grid(True, alpha=0.3, axis='y')
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom', fontsize=10)
        
        # 4. Optimal Configuration Analysis
        ax4 = axes[1, 1]
        
        # Find optimal configurations (best payback for each unit count)
        optimal_configs = []
        for units in df_viz['Min_Units_Required'].unique():
            unit_df = df_viz[df_viz['Min_Units_Required'] == units]
            if not unit_df.empty:
                best_idx = unit_df['Payback_years'].idxmin()
                optimal_configs.append({
                    'Units': units,
                    'Capacity_kWh': unit_df.loc[best_idx, 'Min_Capacity_kWh'],
                    'Payback': unit_df.loc[best_idx, 'Payback_years'],
                    'Threshold': unit_df.loc[best_idx, 'Threshold_kW']
                })
        
        if optimal_configs:
            opt_df = pd.DataFrame(optimal_configs)
            x_pos = np.arange(len(opt_df))
            bars = ax4.bar(x_pos, opt_df['Payback'], color='#2ecc71', alpha=0.7)
            ax4.set_xlabel('Configuration', fontsize=11)
            ax4.set_ylabel('Best Payback Period (years)', fontsize=11)
            ax4.set_title('Optimal Payback by Battery Configuration', fontsize=12, fontweight='bold')
            ax4.set_xticks(x_pos)
            ax4.set_xticklabels([f'{int(u)}×233kWh\n({int(c)}kWh total)' 
                                for u, c in zip(opt_df['Units'], opt_df['Capacity_kWh'])], 
                               fontsize=9)
            ax4.grid(True, alpha=0.3, axis='y')
            
            # Add value labels
            for bar, payback, threshold in zip(bars, opt_df['Payback'], opt_df['Threshold']):
                height = bar.get_height()
                ax4.text(bar.get_x() + bar.get_width()/2., height,
                        f'{payback:.1f}y\n@{threshold:.0f}kW', 
                        ha='center', va='bottom', fontsize=9)
        
        plt.tight_layout()
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()
        
        # Page 2: Comprehensive Summary Statistics
        fig = plt.figure(figsize=(14, 10))
        fig.suptitle('Summary Statistics and Key Findings', fontsize=16, fontweight='bold')
        
        # Create a text summary
        ax = fig.add_subplot(111)
        ax.axis('off')
        
        # Calculate key statistics
        valid_df = df_viz[df_viz['Payback_years'] <= 10]  # Focus on reasonable payback periods
        
        if not valid_df.empty:
            best_payback_idx = valid_df['Payback_years'].idxmin()
            best_efficiency_idx = valid_df['Efficiency'].idxmax()
            
            summary_text = f"""
KEY FINDINGS - Battery Storage Analysis for May-June 2025

OPTIMAL CONFIGURATIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Best Payback Period:
  • Threshold: {valid_df.loc[best_payback_idx, 'Threshold_kW']:.0f} kW
  • Payback: {valid_df.loc[best_payback_idx, 'Payback_years']:.2f} years
  • Battery Size: {valid_df.loc[best_payback_idx, 'Min_Capacity_kWh']:.0f} kWh ({valid_df.loc[best_payback_idx, 'Min_Units_Required']:.0f} units)
  • Peak Reduction: {valid_df.loc[best_payback_idx, 'Highest_Peak_Shaved_kW']:.1f} kW
  • Efficiency: {valid_df.loc[best_payback_idx, 'Efficiency']:.4f} kW/kWh

Highest Efficiency:
  • Threshold: {valid_df.loc[best_efficiency_idx, 'Threshold_kW']:.0f} kW
  • Efficiency: {valid_df.loc[best_efficiency_idx, 'Efficiency']:.4f} kW/kWh
  • Payback: {valid_df.loc[best_efficiency_idx, 'Payback_years']:.2f} years
  • Battery Size: {valid_df.loc[best_efficiency_idx, 'Min_Capacity_kWh']:.0f} kWh ({valid_df.loc[best_efficiency_idx, 'Min_Units_Required']:.0f} units)

ECONOMIC PARAMETERS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  • Capital Cost: RM {CAPEX_PER_KWH:,.0f} per kWh
  • MD Tariff: RM {MD_TARIFF_RM_PER_KW:.2f} per kW
  • Battery Unit Size: {UNIT_KWH:.0f} kWh
  • Max Power per Unit: {MAX_POWER_KW_PER_UNIT:.0f} kW

ANALYSIS RANGE:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  • Threshold Range: {SWEEP_START_KW:.0f} to {SWEEP_END_KW:.0f} kW
  • Step Size: {abs(SWEEP_STEP_KW):.0f} kW
  • Time Window: 14:00 - 22:00 daily
  • Data Period: May - June 2025

RECOMMENDATIONS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
            
            # Add recommendations based on payback targets
            configs_under_5y = valid_df[valid_df['Payback_years'] <= 5]
            if not configs_under_5y.empty:
                summary_text += f"""
  ✓ {len(configs_under_5y)} configurations achieve < 5-year payback
  ✓ Recommended threshold range: {configs_under_5y['Threshold_kW'].min():.0f} - {configs_under_5y['Threshold_kW'].max():.0f} kW
  ✓ Typical battery size needed: {configs_under_5y['Min_Capacity_kWh'].min():.0f} - {configs_under_5y['Min_Capacity_kWh'].max():.0f} kWh
"""
            else:
                summary_text += """
  ⚠ No configurations achieve < 5-year payback with current economics
  ⚠ Consider reviewing capital costs or exploring additional revenue streams
"""
            
        else:
            summary_text = "No valid configurations found in the analysis range."
        
        ax.text(0.05, 0.95, summary_text, transform=ax.transAxes, fontsize=11,
                verticalalignment='top', fontfamily='monospace',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.3))
        
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()
    
    print(f"Saved visualizations to: {VISUALIZATIONS_PDF}")


# =====================
# Main
# =====================
def main():
    # Load May+June, 14:00–22:00 windows
    days = load_months(EXCEL_PATH, SHEET_NAME, YEAR_MONTHS)

    # Compute Highest-day aggregation across all days for each threshold
    rows = compute_threshold_sweep_stats_highest(days)
    df = pd.DataFrame(rows)

    # Include the new columns in the export
    col_order = [
        'Threshold_kW',
        'Highest_Energy_kWh', 'Highest_Energy_Day',
        'Highest_Peak_Shaved_kW', 'Highest_Peak_Day',
        'Min_Units_Required', 'Min_Capacity_kWh',
        'Limiting_Factor', 'Payback_years', 'Efficiency',
        'Fits_1x233', 'Fits_2x233(466)', 'Fits_3x233(699)', 'Fits_4x233(932)'
    ]
    df = df[col_order]

    # Save CSV
    df.to_csv(SWEEP_CSV, index=False)
    print(f"Saved stats CSV: {SWEEP_CSV}")
    
    # Create visualizations instead of PDF table
    create_visualizations(df)


if __name__ == '__main__':
    main()