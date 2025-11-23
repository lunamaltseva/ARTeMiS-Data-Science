import pandas as pd
import re

df = pd.read_excel('../data/artemis/artemis_data.xlsx')

df = df.drop(columns=[
    "Student ID",
    "Rating", "Rating.1", "Rating.2", "Rating.3",
    "Cumulative GPA",
    "Budget Requested (USD)"
], errors="ignore")

df = df.rename(columns={
    "Measurable: How many people do you plan to affect?": "Target Audience",
    "Timebound: How many operating days will the project have?": "Duration",
    "How many people will you need on your team? (staff, volunteers)": "Staff",
    "Greenlit Budget": "Budget",
    "Sum": "Rating",
    "What is the thematic premise of your project?": "Theme"
})

theme = df["Theme"]
numeric_df = df.select_dtypes(include=['number']).copy()
numeric_df.drop(columns=["Sum_clean"], errors='ignore', inplace=True)
df_clean = pd.concat([theme, numeric_df], axis=1)

out_path = "../data/artemis/artemis_data_for_regression.xlsx"
df_clean.to_excel(out_path, index=False)

df_clean.head(), out_path
