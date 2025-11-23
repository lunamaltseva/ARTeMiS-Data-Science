import random
import numpy as np
import pandas as pd

df = pd.read_excel('../data/comparison_data/previous_projects_data.xlsx')
print(f"Rows incoming: {len(df)}")

df = df.dropna(subset=['participants'])
df = df.dropna(subset=['country'])

for col in ['budget']:
    df[col] = df[col].fillna(df['budget'].mean()).apply(lambda x: x * random.uniform(0.8, 1.1))
    df[col] = df[col].clip(lower=0, upper=1000)

numeric_cols = [
    #'participants', 'duration'
]

for col in numeric_cols:
    q1 = df[col].quantile(0.15)
    q3 = df[col].quantile(0.85)
    iqr = q3 - q1

    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr

    df[col] = df[col].clip(lower=lower, upper=upper)

df.to_excel('../data/comparison_data/previous_projects_data_cleaned.xlsx', index=False)

print("Data cleaning complete!")
print(f"Rows remaining: {len(df)}")
