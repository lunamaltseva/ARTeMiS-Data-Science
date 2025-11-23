import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

rcParams['font.family'] = 'serif'
rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
rcParams['font.size'] = 10
rcParams['axes.linewidth'] = 0.8
rcParams['figure.dpi'] = 120

df = pd.read_excel('../data/artemis/artemis_data_cleaned.xlsx')

fig, axes = plt.subplots(1, 4, figsize=(18, 5))
fig.subplots_adjust(wspace=0.1)

df_greenlit = df[df['Status'] == 'Greenlit']

plot_configs = [
    (df, 'Project Country', 'Project Country (All)'),
    (df, 'Theme', 'Theme (All)'),
    (df_greenlit, 'Project Country', 'Project Country (Greenlit)'),
    (df_greenlit, 'Theme', 'Theme (Greenlit)')
]

colors = ['#4A90E2', '#E57373', '#81C784', '#FFD54F', '#BA68C8',
          '#4DB6AC', '#FF8A65', '#9575CD', '#7986CB', '#AED581',
          '#90CAF9', '#A1887F', '#EF9A9A', '#80CBC4', '#C5E1A5']

for idx, (ax, (data_df, col_name, label)) in enumerate(zip(axes, plot_configs)):
    if col_name in data_df.columns:
        data = data_df[col_name].value_counts()

        if len(data) > 10:
            other_count = data.iloc[10:].sum()
            data = data.iloc[:10]
            if other_count > 0:
                data['Other'] = other_count

        truncated_labels = [str(label)[:10] + '...' if len(str(label)) > 10 else str(label)
                            for label in data.index]

        wedges, texts, autotexts = ax.pie(data.values,
                                          labels=truncated_labels,
                                          autopct='%1.1f%%',
                                          startangle=90,
                                          colors=colors[:len(data)],
                                          wedgeprops=dict(edgecolor='black', linewidth=0.8),
                                          textprops=dict(fontsize=8, clip_on=False))

        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
            autotext.set_fontsize(7)
            autotext.set_clip_on(False)

        for text in texts:
            text.set_fontsize(8)
            text.set_fontweight('normal')
            text.set_clip_on(False)

        ax.set_title(f'({chr(97 + idx)}) {label}',
                     fontweight='bold', fontsize=10, pad=12)

fig.suptitle('Distribution of Projects: All vs. Greenlit',
             fontsize=13, fontweight='bold', y=0.98)
fig.tight_layout()

plt.show()