import pandas as pd

# Read the existing Excel file into a DataFrame
file_path = Path("data/sample_mlb_10yrs.xlsx")
df = pd.read_excel(file_path)

# Create readable result labels
df['Covered Spread?'] = df['ATS Margin'].apply(lambda x: 'Push' if x == 0 else ('Yes' if x > 0 else 'No'))
df['Over_Under'] = df['O/U Margin'].apply(lambda x: 'Push' if x == 0 else ('Over' if x > 0 else 'Under'))

# Function to analyze past streaks in a column
def analyze_past_streaks(series, streak_value, streak_length):
    past_streaks = []
    current_streak = 0
    for value in series:
        if value == streak_value:
            current_streak += 1
        else:
            if current_streak >= streak_length:
                past_streaks.append(current_streak)
            current_streak = 0
    if current_streak >= streak_length:
        past_streaks.append(current_streak)

    ended_at_length = sum(1 for s in past_streaks if s == streak_length)
    extended_streaks = [s for s in past_streaks if s > streak_length]

    return len(past_streaks), ended_at_length, extended_streaks

# Initialize streak storage
team_streaks = {}
teams = df['Team'].unique()

# Iterate over each team
for team in teams:
    team_data = df[df['Team'] == team].sort_values(by='Date', ascending=False).reset_index(drop=True)

    # --- COVERED SPREAD STREAK ---
    covered_streak = 0
    covered_type = None
    for result in team_data['Covered Spread?']:
        if covered_streak == 0:
            covered_type = result
            covered_streak = 1
        elif result == covered_type:
            covered_streak += 1
        else:
            break

    if covered_streak >= 4:
        label = "covered the spread" if covered_type == "Yes" else "failed to cover the spread"
        recent_data = team_data.iloc[covered_streak:covered_streak + 100]['Covered Spread?']
        total, ended_at, extended = analyze_past_streaks(recent_data, covered_type, covered_streak)

        if team not in team_streaks:
            team_streaks[team] = {}

        key = 'Cover' if covered_type == 'Yes' else 'NoCover'
        team_streaks[team][key] = {
            'current_length': covered_streak,
            'previous_streaks': total,
            'ended_at': ended_at,
            'extended_lengths': extended
        }

    # --- OVER/UNDER STREAK ---
    ou_streak = 0
    ou_type = None
    for result in team_data['Over_Under']:
        if ou_streak == 0:
            ou_type = result
            ou_streak = 1
        elif result == ou_type:
            ou_streak += 1
        else:
            break

    if ou_streak >= 4:
        recent_data = team_data.iloc[ou_streak:ou_streak + 100]['Over_Under']
        total, ended_at, extended = analyze_past_streaks(recent_data, ou_type, ou_streak)

        if team not in team_streaks:
            team_streaks[team] = {}

        team_streaks[team][ou_type] = {
            'current_length': ou_streak,
            'previous_streaks': total,
            'ended_at': ended_at,
            'extended_lengths': extended
        }

# === Build Output Summary Column ===
output_rows = []

for team, data in team_streaks.items():
    for streak_type, streak_info in data.items():
        current_len = streak_info['current_length']
        prev_streaks = streak_info['previous_streaks']
        ended_at = streak_info['ended_at']
        extended_lengths = streak_info['extended_lengths']

        # Sentence 1: Current streak description
        if streak_type == 'Over':
            summary_sentence = f"{team} has gone Over for the last {current_len} straight games."
        elif streak_type == 'Under':
            summary_sentence = f"{team} has gone Under for the last {current_len} straight games."
        elif streak_type == 'Cover':
            summary_sentence = f"{team} has covered the spread for the last {current_len} straight games."
        elif streak_type == 'NoCover':
            summary_sentence = f"{team} has failed to cover the spread for the last {current_len} straight games."

        # Sentence 2: Similar streaks from past 100 games
        if prev_streaks == 0:
            summary_sentence += f" No other streak(s) of at least {current_len} games {streak_type.lower()} in last 100 games."
        else:
            all_lengths = [current_len] * ended_at + extended_lengths
            all_lengths.sort()
            length_str = ', '.join(str(x) for x in all_lengths)
            summary_sentence += f" {prev_streaks} other similar streak(s) in the last 100 games ({length_str})."

        output_rows.append({'Summary': summary_sentence})

# Create DataFrame from summaries
output_df = pd.DataFrame(output_rows)

# === Export to Excel with Formatting ===
output_file_path = Path("outputs/2025.8.16_mlb_streaks.xlsx")
current_date = pd.Timestamp.now().strftime('%Y-%m-%d')

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    output_df.to_excel(writer, index=False, startrow=0, sheet_name='Trends & Streaks')
    workbook = writer.book
    worksheet = writer.sheets['Trends & Streaks']

    # Add a title at the top
    #title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    #worksheet.merge_range('A1:A1', f'MLB Current Trends & Streaks for {current_date}', title_format)

    # Format header
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'middle',
        'align': 'center',
        'bg_color': '#003366',
        'font_color': 'white',
        'border': 1
    })
    worksheet.write(1, 0, "Summary", header_format)

    # Apply borders to entire table
    num_rows = len(output_df) + 2
    table_range = f"A2:A{num_rows}"
    border_format = workbook.add_format({'border': 1})
    worksheet.conditional_format(table_range, {'type': 'no_blanks', 'format': border_format})
    worksheet.conditional_format(table_range, {'type': 'blanks', 'format': border_format})

print("Excel export complete with condensed summary column.")
