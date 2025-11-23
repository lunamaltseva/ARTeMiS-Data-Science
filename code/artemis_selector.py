import math
import pandas as pd
from typing import List
import matplotlib.pyplot as plt
import numpy as np
from datetime import timedelta
import matplotlib.dates as mdates

def select_projects_dp(
    filepath: str,
    max_budget: int,
    theme_diversity_factor: float,
    country_diversity_factor: float,
    max_states: int = 200_000,
    verbose: bool = False
) -> List[int]:
    df = pd.read_excel(filepath)
    required_cols = {"ID", "country", "theme", "participants", "budget", "rating"}
    if not required_cols.issubset(set(df.columns)):
        raise ValueError(f"input file must contain columns: {required_cols}")

    df = df.copy()
    df['budget'] = df['budget'].fillna(0).astype(int)
    df['participants'] = df['participants'].fillna(0).astype(int)
    df['rating'] = df['rating'].fillna(0.0).astype(float)
    df['ID'] = df['ID'].astype(object)

    themes = sorted(df['theme'].astype(str).unique())
    countries = sorted(df['country'].astype(str).unique())
    theme_index = {t: i for i, t in enumerate(themes)}
    country_index = {c: i for i, c in enumerate(countries)}
    T = len(themes)
    C = len(countries)
    counts_len = T + C

    items = []
    for _, row in df.iterrows():
        items.append({
            'id': row['ID'],
            'theme_idx': theme_index[str(row['theme'])],
            'country_idx': country_index[str(row['country'])],
            'participants': int(row['participants']),
            'budget': int(row['budget']),
            'rating': float(row['rating'])
        })

    n = len(items)

    start_counts = tuple([0] * counts_len)
    dp = {
        (0, 0, start_counts): ((0.0, 0, 0), tuple())
    }

    for idx, it in enumerate(items):
        snapshot = list(dp.items())
        for (sel_count, budget_used, counts_tuple), (obj_tuple, sel_ids) in snapshot:
            new_budget = budget_used + it['budget']
            if new_budget <= max_budget:
                new_sel = sel_count + 1

                counts_list = list(counts_tuple)
                counts_list[it['theme_idx']] += 1
                counts_list[T + it['country_idx']] += 1
                new_counts = tuple(counts_list)

                new_rating = obj_tuple[0] + it['rating']
                new_participants = obj_tuple[1] + it['participants']
                new_neg_budget = obj_tuple[2] - it['budget']

                new_obj = (new_rating, new_participants, new_neg_budget)
                new_sel_ids = sel_ids + (it['id'],)

                key = (new_sel, new_budget, new_counts)
                prev = dp.get(key)
                if (prev is None) or (new_obj > prev[0]):
                    dp[key] = (new_obj, new_sel_ids)

        if len(dp) > max_states:
            scored = [ (v[0], k, v[1]) for k, v in dp.items() ]
            scored.sort(reverse=True, key=lambda x: x[0])
            kept = scored[:max_states]
            dp = { k: (obj, sel_ids) for (obj, k, sel_ids) in kept }

    best_solution = None
    for (sel_count, budget_used, counts_tuple), (obj_tuple, sel_ids) in dp.items():
        k = sel_count
        if k == 0:
            continue
        if budget_used > max_budget:
            continue

        theme_cap = math.floor(theme_diversity_factor * k)
        country_cap = math.floor(country_diversity_factor * k)

        theme_counts = counts_tuple[:T]
        country_counts = counts_tuple[T:]
        max_theme = max(theme_counts) if theme_counts else 0
        max_country = max(country_counts) if country_counts else 0

        if (max_theme <= theme_cap) and (max_country <= country_cap):
            if (best_solution is None) or (obj_tuple > best_solution[0]):
                best_solution = (obj_tuple, sel_ids, k, budget_used, max_theme, max_country)

        return []

    obj, sel_ids, k, total_budget_used, max_theme, max_country = best_solution
    print(f"Selected {k} projects, total_budget={total_budget_used}, max_theme={max_theme}, max_country={max_country}")
    return list(sel_ids)


def plot_participation_timeline(data_path, selected_ids):
    plt.rcParams['font.family'] = 'serif'
    plt.rcParams['font.serif'] = ['Times New Roman']

    df = pd.read_excel(data_path)

    df_selected = df[df['ID'].isin(selected_ids)].copy()
    df_selected['debut'] = pd.to_datetime(df_selected['debut'])
    df_selected['end_date'] = df_selected['debut'] + pd.to_timedelta(df_selected['duration'], unit='D')

    start_date = df_selected['debut'].min()
    end_date = df_selected['end_date'].max()

    april_extension = pd.Timestamp(year=end_date.year if end_date.month <= 4 else end_date.year + 1,
                                   month=4, day=30)
    end_date = max(end_date, april_extension)

    time_points = pd.date_range(start=start_date, end=end_date, freq='D')
    participation_matrix = np.zeros((len(time_points), len(df_selected)))
    df_selected = df_selected.sort_values('debut', ascending=False).reset_index(drop=True)
    for idx, (_, project) in enumerate(df_selected.iterrows()):
        project_start = project['debut']
        project_end = project['end_date']
        project_participants = project['participants']

        for t_idx, t in enumerate(time_points):
            if t < project_start:
                participation_matrix[t_idx, idx] = 0
            elif t <= project_end:
                days_elapsed = (t - project_start).days
                total_days = (project_end - project_start).days
                if total_days > 0:
                    participation_matrix[t_idx, idx] = project_participants * (days_elapsed / total_days)
                else:
                    participation_matrix[t_idx, idx] = project_participants
            else:
                participation_matrix[t_idx, idx] = project_participants

    fig, ax = plt.subplots(figsize=(10, 6), dpi=120)
    line_color = '#2C3E50'

    cumulative_participation = np.cumsum(participation_matrix, axis=1)

    for idx in range(len(df_selected)):
        ax.plot(time_points, cumulative_participation[:, idx],
                color=line_color, linewidth=1.5, alpha=0.7)

    ax.set_xlabel('Date', fontsize=12)
    ax.set_ylabel('Cumulative Participation', fontsize=12)
    ax.set_title('Direct engagement over time', fontsize=13, pad=15)

    ax.xaxis.set_major_formatter(mdates.DateFormatter('%B %Y'))
    ax.xaxis.set_major_locator(mdates.MonthLocator())

    ax.tick_params(axis='both', which='major', labelsize=10)
    ax.tick_params(axis='both', which='minor', labelsize=8)

    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')

    ax.grid(True, alpha=0.2, linestyle='-', linewidth=0.5, color='gray')
    ax.set_axisbelow(True)

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_linewidth(0.8)
    ax.spines['bottom'].set_linewidth(0.8)

    plt.tight_layout()

    return fig, ax


def overlap_coefficient(list1, list2):
    set1, set2 = set(list1), set(list2)
    intersection = len(set1 & set2)
    return intersection / min(len(set1), len(set2)) if min(len(set1), len(set2)) > 0 else 0

if __name__ == "__main__":
    path = "../data/artemis/artemis_data_for_DP.xlsx"
    max_budget = 9700
    theme_diversity_factor = 0.8
    country_diversity_factor = 0.95
    incoming = [1, 2, 3, 4, 11, 13, 22, 24, 29, 32, 36]

    selected = select_projects_dp(path, max_budget, theme_diversity_factor, country_diversity_factor,
                                  max_states=200_000, verbose=True)
    print("Selected IDs:", selected)

    print("Overlap percentage:" + overlap_coefficient(incoming, selected))

    fig, ax = plot_participation_timeline(
        data_path="../data/artemis/artemis_data_for_DP.xlsx",
        selected_ids=selected
    )

    fig, ax = plot_participation_timeline(
        data_path="../data/artemis/artemis_data_for_DP.xlsx",
        selected_ids=incoming
    )

    plt.show()