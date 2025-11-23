import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

rcParams['font.family'] = 'serif'
rcParams['font.serif'] = ['Times New Roman', 'DejaVu Serif']
rcParams['font.size'] = 10
rcParams['axes.linewidth'] = 0.8
rcParams['grid.linewidth'] = 0.5
rcParams['lines.linewidth'] = 1.5

df = pd.read_excel('../data/artemis/artemis_data_for_regression.xlsx')
num_df = df.select_dtypes(include=[np.number])
corr = num_df.corr().abs()

fig1, ax1 = plt.subplots(figsize=(8, 7), dpi=120)
im = ax1.imshow(corr, cmap='RdYlBu_r', interpolation='nearest', vmin=0, vmax=1)

ax1.set_xticks(range(len(corr.columns)))
ax1.set_yticks(range(len(corr.columns)))
ax1.set_xticklabels(corr.columns, rotation=45, ha='right')
ax1.set_yticklabels(corr.columns)

for i in range(corr.shape[0]):
    for j in range(corr.shape[1]):
        color = 'white' if corr.iloc[i, j] > 0.5 else 'black'
        ax1.text(j, i, f"{corr.iloc[i, j]:.2f}",
                 ha='center', va='center', color=color, fontsize=8)

cbar = plt.colorbar(im, ax=ax1, fraction=0.046, pad=0.04)
cbar.set_label('|Correlation Coefficient|', rotation=270, labelpad=20)

ax1.set_title("Correlation Matrix of Regression Variables",
              fontsize=12, fontweight='bold', pad=15)
fig1.tight_layout()

df = pd.read_excel('../data/artemis/artemis_data_for_regression.xlsx')

y = df["Rating"]
X_vars = {
    "Budget": df["Budget"],
    "Target Audience": df["Target Audience"],
    "Staff": df["Staff"],
    "Duration": df["Duration"]
}

degree = 2

fig2, axes = plt.subplots(2, 2, figsize=(10, 8), dpi=120)
axes = axes.flatten()

for ax, (name, x) in zip(axes, X_vars.items()):
    mask = ~(x.isna() | y.isna())
    x_clean = x[mask]
    y_clean = y[mask]

    coef = np.polyfit(x_clean, y_clean, degree)
    poly = np.poly1d(coef)

    y_pred = poly(x_clean)
    ss_res = np.sum((y_clean - y_pred) ** 2)
    ss_tot = np.sum((y_clean - np.mean(y_clean)) ** 2)
    r_squared = 1 - (ss_res / ss_tot)

    xs = np.linspace(x_clean.min(), x_clean.max(), 300)

    ax.scatter(x_clean, y_clean, alpha=0.6, s=30,
               edgecolors='black', linewidths=0.5, color='#2E86AB')

    ax.plot(xs, poly(xs), color='#A23B72', linewidth=2,
            label=f'$R^2 = {r_squared:.3f}$')

    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)

    ax.set_xlabel(name, fontsize=10, fontweight='bold')
    ax.set_ylabel("Rating", fontsize=10, fontweight='bold')
    ax.set_title(f"({chr(97 + list(X_vars.keys()).index(name))}) Rating vs. {name}",
                 fontsize=10, fontweight='bold', loc='left')
    ax.legend(loc='best', frameon=True, fontsize=8)

    ax.set_xlim(x_clean.min() - 0.05 * (x_clean.max() - x_clean.min()),
                x_clean.max() + 0.05 * (x_clean.max() - x_clean.min()))

fig2.suptitle("Polynomial Regression Analysis (Degree 2)",
              fontsize=12, fontweight='bold', y=0.995)
fig2.tight_layout()
plt.show()