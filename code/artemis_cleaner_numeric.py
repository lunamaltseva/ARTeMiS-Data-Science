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
    "Measurable: How many people do you plan to affect?": "participants",
    "Timebound: How many operating days will the project have?": "duration",
    "How many people will you need on your team? (staff, volunteers)": "staff",
    "Greenlit Budget": "budget",
    "Sum": "rating",
    "What is the thematic premise of your project?": "theme",
    "In which country will your project take place?": "country"
})

theme = df["theme"]
country = df["country"]

numeric_df = df.select_dtypes(include=["number"]).copy()
numeric_df.drop(columns=["staff"], errors="ignore", inplace=True)
df_clean = pd.concat([country, theme, numeric_df], axis=1)

out_path = "../data/artemis/artemis_data_numeric.xlsx"
df_clean.to_excel(out_path, index=False)

df_clean.head(), out_path
