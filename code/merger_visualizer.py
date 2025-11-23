"""
Academic visualization of interval analysis data.
Reads ../data/interval_data/interval_analysis_applied.xlsx and creates a 2x3 figure.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams
from pathlib import Path

rcParams['font.family'] = 'serif'
rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
rcParams['font.size'] = 11
rcParams['axes.linewidth'] = 0.8
rcParams['grid.linewidth'] = 0.5
rcParams['lines.linewidth'] = 2
rcParams['figure.dpi'] = 120

FILE_PATH = Path("../data/interval_data/interval_analysis_applied.xlsx")
FIGSIZE = (16, 8.33)

COLS = {
    'participants': {
        'current': 'current_participants',
        'ci_low': 'participation_ci_low',
        'ci_high': 'participation_ci_high',
        'min': 'participants_pred_min',
        'mean': 'participants_pred_mean',
        'max': 'participants_pred_max'
    },
    'duration': {
        'current': 'current_duration',
        'ci_low': 'duration_ci_low',
        'ci_high': 'duration_ci_high',
        'min': 'duration_pred_min',
        'mean': 'duration_pred_mean',
        'max': 'duration_pred_max'
    },
    'budget': {
        'current': 'current_budget',
        'ci_low': 'budget_ci_low',
        'ci_high': 'budget_ci_high',
        'min': 'budget_pred_min',
        'mean': 'budget_pred_mean',
        'max': 'budget_pred_max'
    }
}

YLIM_UPPER_THRESHOLD = 2200
YLIM_LOWER = 0

COLOR_OBSERVED = '#2E86AB'
COLOR_CI_LOW = '#A23B72'
COLOR_CI_HIGH = '#F18F01'
COLOR_MIN = '#C73E1D'
COLOR_MEAN = '#6A994E'
COLOR_MAX = '#BC4B51'
COLOR_ZERO = '#333333'

def to_numeric_safe(series):
    """Convert series to numeric, handling commas."""
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False), errors='coerce')

def set_ylim_if_exceeds(ax, series_list, upper_threshold=YLIM_UPPER_THRESHOLD, lower_bound=YLIM_LOWER):
    """Set y-limits if max value exceeds threshold."""
    max_vals = []
    for s in series_list:
        if s is None:
            continue
        arr = np.asarray(s, dtype=float)
        if arr.size > 0 and np.isfinite(arr).any():
            max_vals.append(np.nanmax(arr[np.isfinite(arr)]))
    if not max_vals:
        return
    overall_max = max(max_vals)
    if overall_max > upper_threshold:
        ax.set_ylim(lower_bound, upper_threshold)

def plot_top(ax, df, mapping, title):
    """Plot observed values and confidence intervals."""
    s_cur = to_numeric_safe(df[mapping['current']]) if mapping['current'] in df.columns else pd.Series([np.nan]*len(df))
    s_low = to_numeric_safe(df[mapping['ci_low']]) if mapping['ci_low'] in df.columns else pd.Series([np.nan]*len(df))
    s_high = to_numeric_safe(df[mapping['ci_high']]) if mapping['ci_high'] in df.columns else pd.Series([np.nan]*len(df))

    if s_cur.notna().any():
        order = s_cur.sort_values(ascending=True, na_position='last').index
    else:
        order = df.index

    s_cur = s_cur.loc[order].reset_index(drop=True)
    s_low = s_low.loc[order].reset_index(drop=True)
    s_high = s_high.loc[order].reset_index(drop=True)

    x = np.arange(len(s_cur))

    artists = []

    if s_cur.notna().any():
        h_cur, = ax.plot(x, s_cur, color=COLOR_OBSERVED, linewidth=2, label='Observed')
        artists.append((h_cur, 'Observed'))

    if s_low.notna().any():
        h_low, = ax.plot(x, s_low, color=COLOR_CI_LOW, linestyle='--', linewidth=1.5, label='CI Lower')
        artists.append((h_low, 'CI Lower'))

    if s_high.notna().any():
        h_high, = ax.plot(x, s_high, color=COLOR_CI_HIGH, linestyle=':', linewidth=1.5, label='CI Upper')
        artists.append((h_high, 'CI Upper'))

    ax.axhline(0, color=COLOR_ZERO, linewidth=0.8, linestyle='-', alpha=0.5)

    set_ylim_if_exceeds(ax, [s_cur, s_low, s_high])

    ax.set_title(f'{title}', fontweight='bold', fontsize=12, pad=12)
    ax.set_ylabel('Value', fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_xticks([])

    return artists

def plot_bottom(ax, df, mapping, title):
    """Plot predicted values (min, mean, max)."""
    s_min = to_numeric_safe(df[mapping['min']]) if mapping['min'] in df.columns else pd.Series([np.nan]*len(df))
    s_mean = to_numeric_safe(df[mapping['mean']]) if mapping['mean'] in df.columns else pd.Series([np.nan]*len(df))
    s_max = to_numeric_safe(df[mapping['max']]) if mapping['max'] in df.columns else pd.Series([np.nan]*len(df))

    key = s_mean if s_mean.notna().any() else s_min
    if key.notna().any():
        order = key.sort_values(ascending=True, na_position='last').index
    else:
        order = df.index

    s_min = s_min.loc[order].reset_index(drop=True)
    s_mean = s_mean.loc[order].reset_index(drop=True)
    s_max = s_max.loc[order].reset_index(drop=True)

    x = np.arange(len(s_min))

    artists = []

    if s_min.notna().any():
        h_min, = ax.plot(x, s_min, color=COLOR_MIN, linestyle='--', linewidth=1.5, label='Pred. Min')
        artists.append((h_min, 'Pred. Min'))

    if s_mean.notna().any():
        h_mean, = ax.plot(x, s_mean, color=COLOR_MEAN, linewidth=2, label='Pred. Mean')
        artists.append((h_mean, 'Pred. Mean'))

    if s_max.notna().any():
        h_max, = ax.plot(x, s_max, color=COLOR_MAX, linestyle=':', linewidth=1.5, label='Pred. Max')
        artists.append((h_max, 'Pred. Max'))

    ax.axhline(0, color=COLOR_ZERO, linewidth=0.8, linestyle='-', alpha=0.5)

    set_ylim_if_exceeds(ax, [s_min, s_mean, s_max])

    ax.set_ylabel('Value', fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.set_xticks([])

    return artists

def main():
    df = pd.read_excel(FILE_PATH)

    fig, axes = plt.subplots(2, 3, figsize=FIGSIZE)
    fig.suptitle('Interval Analysis: Observed vs. Predicted Values',
                 fontsize=14, fontweight='bold', y=0.97)

    try:
        fig.canvas.manager.set_window_title('Interval Analysis Visualization')
    except Exception:
        pass

    collected_handles = []

    titles = ['(a) Participants', '(b) Duration', '(c) Budget']
    for idx, (ax, key, title) in enumerate(zip(axes[0], ['participants', 'duration', 'budget'], titles)):
        mapping = COLS[key]
        handles = plot_top(ax, df, mapping, title)
        if idx == 0:
            collected_handles.extend(handles)

    titles = ['(d) Participants', '(e) Duration', '(f) Budget']
    for idx, (ax, key, title) in enumerate(zip(axes[1], ['participants', 'duration', 'budget'], titles)):
        mapping = COLS[key]
        handles = plot_bottom(ax, df, mapping, title)
        if idx == 0:
            collected_handles.extend(handles)

    seen = set()
    unique_handles = []
    for art, lab in collected_handles:
        if lab not in seen:
            seen.add(lab)
            unique_handles.append((art, lab))

    if unique_handles:
        artists, labels = zip(*unique_handles)
        fig.legend(artists, labels, loc='lower center', ncol=len(labels),
                  frameon=True, fontsize=10, bbox_to_anchor=(0.5, -0.01),
                  framealpha=0.9, edgecolor='black')

    plt.subplots_adjust(top=0.88, bottom=0.10, hspace=0.40, wspace=0.20)
    plt.show()

if __name__ == "__main__":
    main()