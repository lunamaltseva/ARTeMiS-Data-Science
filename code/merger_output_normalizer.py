import pandas as pd
import numpy as np

df_advanced = pd.read_excel("../data/interval_data/interval_analysis.xlsx")
df_simple = pd.read_excel("../data/interval_data/interval_analysis_simple.xlsx")

budget_ci_low = df_simple["current_budget"] - df_simple["current_budget"] * np.maximum(df_advanced["budget_part_p10"], df_advanced["budget_duration_p10"])
budget_ci_high = df_simple["current_budget"] + df_simple["current_budget"] * np.minimum(df_advanced["budget_part_p10"], df_advanced["budget_duration_p10"])
participation_ci_low = df_simple["current_participants"] - df_simple["current_participants"] * np.maximum(df_advanced["part_duration_p10"], df_advanced["part_budget_p10"])
participation_ci_high = df_simple["current_participants"] + df_simple["current_participants"] * np.minimum(df_advanced["part_duration_p10"], df_advanced["part_budget_p10"])
duration_ci_low = df_simple["current_duration"] - df_simple["current_duration"] * np.maximum(df_advanced["duration_budget_p10"], df_advanced["duration_part_p10"])
duration_ci_high = df_simple["current_duration"] + df_simple["current_duration"] * np.minimum(df_advanced["duration_budget_p10"], df_advanced["duration_part_p10"])

df_simple["budget_ci_low"] = budget_ci_low
df_simple["budget_ci_high"] = budget_ci_high
df_simple["participation_ci_low"] = participation_ci_low
df_simple["participation_ci_high"] = participation_ci_high
df_simple["duration_ci_low"] = duration_ci_low
df_simple["duration_ci_high"] = duration_ci_high

df_simple.to_csv("../data/interval_data/interval_analysis_applied.csv", index=False)

print("Normalized data saved to 'interval_analysis_applied.xlsx'")
print(df_simple.head())