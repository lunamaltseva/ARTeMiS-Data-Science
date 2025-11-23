import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

rcParams['font.family'] = 'serif'
rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
rcParams['font.size'] = 10
rcParams['axes.linewidth'] = 0.8
rcParams['grid.linewidth'] = 0.5
rcParams['figure.dpi'] = 120

df = pd.read_excel('../data/artemis/artemis_data_cleaned.xlsx')

rating_columns = ['Relevance Rating', 'Performance Rating', 'Planning Rating',
                  'Necessities Rating', 'Budget Rating', 'Experience Rating']
df['Average Rating'] = df[rating_columns].mean(axis=1)

fig, axes = plt.subplots(1, 4, figsize=(18, 5))

variables = [
    ('Target Audience Size', 'Participation'),
    ('Budget Requested', 'Budget (Requested)'),
    ('Project Duration', 'Duration'),
    ('Average Rating', 'Average Rating')
]

for idx, (ax, (col_name, label)) in enumerate(zip(axes, variables)):
    if col_name in df.columns:
        data = df[col_name].dropna()

        counts, bins = np.histogram(data, bins=15)
        bin_centers = (bins[:-1] + bins[1:]) / 2

        bars = ax.bar(bin_centers, counts, width=(bins[1] - bins[0]) * 0.9,
                      color='#4A90E2', edgecolor='black', linewidth=0.8, alpha=0.8)

        mean_val = data.mean()
        median_val = data.median()

        ax.axvline(mean_val, color='#D32F2F', linestyle='--',
                   linewidth=2, label=f'Mean: {mean_val:.1f}')
        ax.axvline(median_val, color='#388E3C', linestyle=':',
                   linewidth=2, label=f'Median: {median_val:.1f}')

        ax.set_xlabel(label, fontsize=11, fontweight='bold')
        ax.set_ylabel('Frequency', fontsize=11, fontweight='bold')
        ax.set_title(f'({chr(97 + idx)}) {label}',
                     fontweight='bold', fontsize=11, loc='left', pad=10)
        ax.legend(loc='upper right', fontsize=9, frameon=True, framealpha=0.9)
        ax.grid(True, alpha=0.3, axis='y', linestyle='--', linewidth=0.5)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        for bar in bars:
            height = bar.get_height()
            if height > 0:
                ax.text(bar.get_x() + bar.get_width() / 2., height,
                        f'{int(height)}',
                        ha='center', va='bottom', fontsize=8)

fig.suptitle('Distribution of Key Project Variables',
             fontsize=13, fontweight='bold', y=0.98)
fig.tight_layout()

plt.show()