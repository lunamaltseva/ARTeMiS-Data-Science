import pandas as pd
import numpy as np

df_artemis = pd.read_excel('../data/artemis/artemis_data.xlsx')
df_pred    = pd.read_excel('../data/interval_data/interval_analysis_applied.xlsx')

df_artemis = df_artemis.drop(columns=[
    "Student ID",
    "Rating", "Rating.1", "Rating.2", "Rating.3",
    "Cumulative GPA",
    "Budget Requested (USD)"
], errors="ignore")

df_artemis = df_artemis.rename(columns={
    "Measurable: How many people do you plan to affect?": "participants",
    "Timebound: How many operating days will the project have?": "duration",
    "How many people will you need on your team? (staff, volunteers)": "staff",
    "Greenlit Budget": "budget",
    "Sum": "rating",
    "What is the thematic premise of your project?": "theme",
    "In which country will your project take place?": "country"
})

numeric_df = df_artemis.select_dtypes(include=["number"]).drop(columns=["staff"], errors="ignore")
df_clean = pd.concat([df_artemis[["country", "theme"]], numeric_df], axis=1)

df = df_pred.copy()

min_c, mean_c, max_c = (
    'participants_pred_min',
    'participants_pred_mean',
    'participants_pred_max'
)

orig = df['current_participants']

df[mean_c] = df[mean_c].fillna(orig)
df[min_c]  = df[min_c].fillna(df[mean_c])
df[max_c]  = df[max_c].fillna(df[mean_c])

df[min_c] = np.minimum(df[min_c], df[mean_c])
df[max_c] = np.maximum(df[max_c], df[mean_c])

df[[min_c, mean_c, max_c]] = df[[min_c, mean_c, max_c]].clip(lower=0)

spread = df[max_c] - df[min_c]

MIN_SPREAD_FACTOR = 0.05
MAX_SPREAD_FACTOR = 1.00

min_spread = (df[mean_c] * MIN_SPREAD_FACTOR).clip(lower=1e-6)
max_spread = (df[mean_c] * MAX_SPREAD_FACTOR).clip(lower=min_spread)

spread = spread.clip(lower=min_spread, upper=max_spread)

lower = (df[mean_c] - spread/2).clip(lower=0)
upper = df[mean_c] + spread/2

RATING_INFLECTION = 70

norm = (df_artemis['rating'] - RATING_INFLECTION) / (1 - RATING_INFLECTION)
norm = norm.clip(-1, 1).fillna(0)

ci_conf = spread / df[mean_c].replace(0, np.nan)
ci_conf = 1 / (1 + ci_conf)
ci_conf = ci_conf.fillna(0.05)
ci_conf = ci_conf.clip(0.05, 0.8)

ci_target = df[mean_c] + np.where(
    norm >= 0,
    norm * (upper - df[mean_c]),
    norm * (df[mean_c] - lower)
)

df['new_participants'] = (
    orig * (1 - ci_conf) +
    ci_target * ci_conf
).round().astype(int)

df_clean['participants'] = df['new_participants']

out_path = "../data/artemis/artemis_data_for_DP.xlsx"
df_clean.to_excel(out_path, index=False)

df_clean.head(), out_path
