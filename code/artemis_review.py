import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (10, 6)

df = pd.read_excel('../data/artemis/artemis_data.xlsx')

print("Dataset Shape:", df.shape)
print("\n" + "=" * 50)
print("Column Names and Types:")
print("=" * 50)
print(df.dtypes)
print("\n" + "=" * 50)
print("First Few Rows:")
print("=" * 50)
print(df.head())

rating_columns = [col for col in df.columns if 'Rating' in col]
print("\n" + "=" * 50)
print("Rating Columns Found:")
print("=" * 50)
for col in rating_columns:
    print(f"  - {col}")

for col in rating_columns:
    if df[col].dtype == 'object':
        df[col] = df[col].str.rstrip('%').astype('float')

print("\n" + "=" * 50)
print("Rating Statistics:")
print("=" * 50)
print(df[rating_columns].describe())

categorical_columns = df.select_dtypes(include=['object']).columns.tolist()
print("\n" + "=" * 50)
print("Categorical Columns Available:")
print("=" * 50)
for col in categorical_columns:
    print(f"  - {col} ({df[col].nunique()} unique values)")