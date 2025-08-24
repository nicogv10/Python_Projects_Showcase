import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
import os
import shutil

# Clear the undetected-chromedriver cache
cache_dir = os.path.join(os.path.expanduser("~"), ".undetected_chromedriver")
if os.path.exists(cache_dir):
    shutil.rmtree(cache_dir)
    print("Cleared undetected-chromedriver cache")

# Set Pandas option to display all columns
pd.set_option('display.max_columns', None)

#this function detects reverse line movement based on opening line vs current line (odds) for plays receiving majority of the $
def reverse_line_mov(open_line, current_line, percentage_money):
    try:
        open_value = float(open_line.replace('+', '').strip())
        current_value = float(current_line.replace('+', '').strip())
        if percentage_money > 60:
            if current_value > open_value:
                return "Potential reverse line movement"
    except ValueError:
        pass
    return ""

#this function detects reverse line movement for game totals (overs & unders)
def check_reverse_line(open_line, current_line, percentage_money):
    try:
        open_val = float(open_line.replace('o', '').replace('u', '').replace('+', '').strip())
        current_val = float(current_line.replace('o', '').replace('u', '').replace('+', '').strip())
        if percentage_money > 60:
            if open_line.startswith('o') and current_val < open_val:
                return "Potential reverse line movement"
            elif open_line.startswith('u') and current_val > open_val:
                return "Potential reverse line movement"
    except ValueError:
        pass
    return ""

#this function extracts all the game data we will export to Excel and use for analysis
def extract_game_data(rows):
    # Initialize a dictionary to store all game data for a single game
    game_data = {}

    if len(rows) < 3:
        print("Insufficient data in rows to extract game information.")
        return game_data

    try:
        # Extract team names
        first_team_name = rows[0].find_element(By.XPATH, ".//div[@class='game-info__team-info']/div[@class='game-info__team--desktop']").text
        second_team_name = rows[0].find_elements(By.XPATH, ".//div[@class='game-info__team-info']/div[@class='game-info__team--desktop']")[1].text

        game_data['Team 1'] = first_team_name
        game_data['Team 2'] = second_team_name

        # Extract Spread data for Team 1 and Team 2
        game_data['Spread Team 1'] = {
            'Open': rows[0].find_element(By.XPATH, ".//div[@class='public-betting__open-container']/div[1]").text,
            'Current Line': rows[0].find_element(By.XPATH, ".//div[@class='book-cell__odds']/span[1]").text,
            'Current Odds': rows[0].find_element(By.XPATH, ".//div[@class='book-cell__odds']/span[2]").text,
            '% Bets': rows[0].find_element(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][1]//span[@class='highlight-text__children']").text,
            '% Money': rows[0].find_elements(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][1]//span[@class='highlight-text__children']")[1].text
        }

        game_data['Spread Team 2'] = {
            'Open': rows[0].find_element(By.XPATH, ".//div[@class='public-betting__open-container']/div[2]").text,
            'Current Line': rows[0].find_elements(By.XPATH, ".//div[@class='book-cell__odds']/span[1]")[1].text,
            'Current Odds': rows[0].find_elements(By.XPATH, ".//div[@class='book-cell__odds']/span[2]")[1].text,
            '% Bets': rows[0].find_element(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][2]//span[@class='highlight-text__children']").text,
            '% Money': rows[0].find_elements(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][2]//span[@class='highlight-text__children']")[1].text
        }

        # Extract Over/Under data
        game_data['Over'] = {
            'Open': rows[1].find_element(By.XPATH, ".//div[@class='public-betting__open-container']/div[1]").text,
            'Current Line': rows[1].find_element(By.XPATH, ".//div[@class='book-cell__odds']/span[1]").text,
            'Current Odds': rows[1].find_element(By.XPATH, ".//div[@class='book-cell__odds']/span[2]").text,
            '% Bets': rows[1].find_element(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][1]//span[@class='highlight-text__children']").text,
            '% Money': rows[1].find_elements(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][1]//span[@class='highlight-text__children']")[1].text
        }

        game_data['Under'] = {
            'Open': rows[1].find_element(By.XPATH, ".//div[@class='public-betting__open-container']/div[2]").text,
            'Current Line': rows[1].find_elements(By.XPATH, ".//div[@class='book-cell__odds']/span[1]")[1].text,
            'Current Odds': rows[1].find_elements(By.XPATH, ".//div[@class='book-cell__odds']/span[2]")[1].text,
            '% Bets': rows[1].find_element(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][2]//span[@class='highlight-text__children']").text,
            '% Money': rows[1].find_elements(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][2]//span[@class='highlight-text__children']")[1].text
        }

        # Extract Moneyline data for Team 1 and Team 2
        game_data['Moneyline Team 1'] = {
            'Open': rows[2].find_element(By.XPATH, ".//div[@class='public-betting__open-container']/div[1]").text,
            'Current Odds': rows[2].find_element(By.XPATH, ".//div[@class='book-cell__odds']/span[1]").text,
            '% Bets': rows[2].find_element(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][1]//span[@class='highlight-text__children']").text,
            '% Money': rows[2].find_elements(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][1]//span[@class='highlight-text__children']")[1].text
        }

        game_data['Moneyline Team 2'] = {
            'Open': rows[2].find_element(By.XPATH, ".//div[@class='public-betting__open-container']/div[2]").text,
            'Current Odds': rows[2].find_elements(By.XPATH, ".//div[@class='book-cell__odds']/span[1]")[1].text,
            '% Bets': rows[2].find_element(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][2]//span[@class='highlight-text__children']").text,
            '% Money': rows[2].find_elements(By.XPATH, ".//div[@class='public-betting__percent-and-bar'][2]//span[@class='highlight-text__children']")[1].text
        }
    except Exception as e:
        print(f"Error extracting game data: {e}")

    return game_data

#once we've extracted the data we move it into a formatted dataframe
def construct_dataframe(all_game_data):
    df_rows = []
    for game in all_game_data:
        try:
            game_name = f"{game['Team 1']} vs {game['Team 2']}"

            spread_diff_team_1 = int(game['Spread Team 1']['% Money'].strip('%')) - int(game['Spread Team 1']['% Bets'].strip('%'))
            spread_diff_team_2 = int(game['Spread Team 2']['% Money'].strip('%')) - int(game['Spread Team 2']['% Bets'].strip('%'))

            spread_potential_sharp_team_1 = f"Potential sharp play ({spread_diff_team_1}% diff)" if spread_diff_team_1 > 10 else ""
            spread_potential_sharp_team_2 = f"Potential sharp play ({spread_diff_team_2}% diff)" if spread_diff_team_2 > 10 else ""

            ml_diff_team_1 = int(game['Moneyline Team 1']['% Money'].strip('%')) - int(game['Moneyline Team 1']['% Bets'].strip('%'))
            ml_diff_team_2 = int(game['Moneyline Team 2']['% Money'].strip('%')) - int(game['Moneyline Team 2']['% Bets'].strip('%'))

            ml_potential_sharp_team_1 = f"Potential sharp play ({ml_diff_team_1}% diff)" if ml_diff_team_1 > 10 else ""
            ml_potential_sharp_team_2 = f"Potential sharp play ({ml_diff_team_2}% diff)" if ml_diff_team_2 > 10 else ""

            ou_diff_over = int(game['Over']['% Money'].strip('%')) - int(game['Over']['% Bets'].strip('%'))
            ou_diff_under = int(game['Under']['% Money'].strip('%')) - int(game['Under']['% Bets'].strip('%'))

            ou_potential_sharp_over = f"Potential sharp play ({ou_diff_over}% diff)" if ou_diff_over > 10 else ""
            ou_potential_sharp_under = f"Potential sharp play ({ou_diff_under}% diff)" if ou_diff_under > 10 else ""

            spread_reverse_line_team_1 = reverse_line_mov(
                game['Spread Team 1']['Open'], game['Spread Team 1']['Current Line'],
                int(game['Spread Team 1']['% Money'].strip('%'))
            )
            spread_reverse_line_team_2 = reverse_line_mov(
                game['Spread Team 2']['Open'], game['Spread Team 2']['Current Line'],
                int(game['Spread Team 2']['% Money'].strip('%'))
            )

            ml_reverse_line_team_1 = reverse_line_mov(
                game['Moneyline Team 1']['Open'], game['Moneyline Team 1']['Current Odds'],
                int(game['Moneyline Team 1']['% Money'].strip('%'))
            )
            ml_reverse_line_team_2 = reverse_line_mov(
                game['Moneyline Team 2']['Open'], game['Moneyline Team 2']['Current Odds'],
                int(game['Moneyline Team 2']['% Money'].strip('%'))
            )

            ou_reverse_line_over = check_reverse_line(
                game['Over']['Open'], game['Over']['Current Line'],
                int(game['Over']['% Money'].strip('%'))
            )
            ou_reverse_line_under = check_reverse_line(
                game['Under']['Open'], game['Under']['Current Line'],
                int(game['Under']['% Money'].strip('%'))
            )

            df_rows.append([
                game_name,
                f"{game['Team 1']} {game['Spread Team 1']['Open']}",
                f"{game['Team 1']} {game['Spread Team 1']['Current Line']}",
                game['Spread Team 1']['Current Odds'],
                game['Spread Team 1']['% Bets'],
                game['Spread Team 1']['% Money'],
                spread_potential_sharp_team_1,
                spread_reverse_line_team_1,
                game['Moneyline Team 1']['Open'],
                game['Moneyline Team 1']['Current Odds'],
                game['Moneyline Team 1']['% Bets'],
                game['Moneyline Team 1']['% Money'],
                ml_potential_sharp_team_1,
                ml_reverse_line_team_1,
                game['Over']['Open'],
                game['Over']['Current Line'],
                game['Over']['Current Odds'],
                game['Over']['% Bets'],
                game['Over']['% Money'],
                ou_potential_sharp_over,
                ou_reverse_line_over
            ])
            df_rows.append([
                game_name,
                f"{game['Team 2']} {game['Spread Team 2']['Open']}",
                f"{game['Team 2']} {game['Spread Team 2']['Current Line']}",
                game['Spread Team 2']['Current Odds'],
                game['Spread Team 2']['% Bets'],
                game['Spread Team 2']['% Money'],
                spread_potential_sharp_team_2,
                spread_reverse_line_team_2,
                game['Moneyline Team 2']['Open'],
                game['Moneyline Team 2']['Current Odds'],
                game['Moneyline Team 2']['% Bets'],
                game['Moneyline Team 2']['% Money'],
                ml_potential_sharp_team_2,
                ml_reverse_line_team_2,
                game['Under']['Open'],
                game['Under']['Current Line'],
                game['Under']['Current Odds'],
                game['Under']['% Bets'],
                game['Under']['% Money'],
                ou_potential_sharp_under,
                ou_reverse_line_under
            ])
        except KeyError as e:
            print(f"Missing data for game '{game.get('Team 1', 'Unknown')} vs {game.get('Team 2', 'Unknown')}': {e}")

    columns = [
        'Game', 'Opening Spread', 'Current Spread', 'Spread Current Odds',
        'Spread % of bets', 'Spread % of $', 'Spread Potential Sharp',
        'Spread Reverse Line Movement', 'ML opening odds', 'ML current odds',
        'ML % of bets', 'ML % of $', 'ML Potential Sharp', 'ML Reverse Line Movement',
        'Opening O/U line', 'Current O/U line', 'Current O/U Odds', 'O/U % of bets',
        'O/U % of $', 'O/U Potential Sharp', 'O/U Reverse Line Movement'
    ]
    df = pd.DataFrame(df_rows, columns=columns)
    return df

#next we create a summary dataframe that only details potential sharp plays and RLM plays
def construct_summary_dataframes(all_game_data):
    # Lists to store rows for sharp plays and reverse line movement
    sharp_plays_rows = []
    reverse_line_movement_rows = []

    for game in all_game_data:
        try:
            game_name = f"{game['Team 1']} vs {game['Team 2']}"

            # Calculate differences and determine sharp plays for Spread and Over/Under
            spread_diff_team_1 = int(game['Spread Team 1']['% Money'].strip('%')) - int(game['Spread Team 1']['% Bets'].strip('%'))
            spread_diff_team_2 = int(game['Spread Team 2']['% Money'].strip('%')) - int(game['Spread Team 2']['% Bets'].strip('%'))

            ou_diff_over = int(game['Over']['% Money'].strip('%')) - int(game['Over']['% Bets'].strip('%'))
            ou_diff_under = int(game['Under']['% Money'].strip('%')) - int(game['Under']['% Bets'].strip('%'))

            # Add sharp plays with a threshold of 13% or greater for Spread and Over/Under
            if spread_diff_team_1 >= 13:
                # Move spread after team name
                sharp_plays_rows.append([
                    f"{game['Team 1']} {game['Spread Team 1']['Current Line']}",
                    game['Spread Team 1']['% Bets'],
                    game['Spread Team 1']['% Money'],
                    f'{spread_diff_team_1}% diff'
                ])
            if spread_diff_team_2 >= 13:
                sharp_plays_rows.append([
                    f"{game['Team 2']} {game['Spread Team 2']['Current Line']}",
                    game['Spread Team 2']['% Bets'],
                    game['Spread Team 2']['% Money'],
                    f'{spread_diff_team_2}% diff'
                ])
            if ou_diff_over >= 13:
                sharp_plays_rows.append([
                    f"Over {game['Over']['Current Line']} {game_name}",
                    game['Over']['% Bets'],
                    game['Over']['% Money'],
                    f'{ou_diff_over}% diff'
                ])
            if ou_diff_under >= 13:
                sharp_plays_rows.append([
                    f"Under {game['Under']['Current Line']} {game_name}",
                    game['Under']['% Bets'],
                    game['Under']['% Money'],
                    f'{ou_diff_under}% diff'
                ])

            # Check for reverse line movement for any bets
            spread_reverse_line_team_1 = reverse_line_mov(
                game['Spread Team 1']['Open'], game['Spread Team 1']['Current Line'],
                int(game['Spread Team 1']['% Money'].strip('%'))
            )
            spread_reverse_line_team_2 = reverse_line_mov(
                game['Spread Team 2']['Open'], game['Spread Team 2']['Current Line'],
                int(game['Spread Team 2']['% Money'].strip('%'))
            )

            ml_reverse_line_team_1 = reverse_line_mov(
                game['Moneyline Team 1']['Open'], game['Moneyline Team 1']['Current Odds'],
                int(game['Moneyline Team 1']['% Money'].strip('%'))
            )
            ml_reverse_line_team_2 = reverse_line_mov(
                game['Moneyline Team 2']['Open'], game['Moneyline Team 2']['Current Odds'],
                int(game['Moneyline Team 2']['% Money'].strip('%'))
            )

            ou_reverse_line_over = check_reverse_line(
                game['Over']['Open'], game['Over']['Current Line'],
                int(game['Over']['% Money'].strip('%'))
            )
            ou_reverse_line_under = check_reverse_line(
                game['Under']['Open'], game['Under']['Current Line'],
                int(game['Under']['% Money'].strip('%'))
            )

            # Add to reverse line movement list if any are found
            if spread_reverse_line_team_1:
                reverse_line_movement_rows.append([f"{game['Team 1']} Spread", "RLM (potentially avoid)"])
            if spread_reverse_line_team_2:
                reverse_line_movement_rows.append([f"{game['Team 2']} Spread", "RLM (potentially avoid)"])
            if ml_reverse_line_team_1:
                reverse_line_movement_rows.append([f"{game['Team 1']} Moneyline", "RLM (potentially avoid)"])
            if ml_reverse_line_team_2:
                reverse_line_movement_rows.append([f"{game['Team 2']} Moneyline", "RLM (potentially avoid)"])
            if ou_reverse_line_over:
                reverse_line_movement_rows.append([f"{game_name} Over", "RLM (potentially avoid)"])
            if ou_reverse_line_under:
                reverse_line_movement_rows.append([f"{game_name} Under", "RLM (potentially avoid)"])
        except KeyError as e:
            print(f"Missing data for game '{game.get('Team 1', 'Unknown')} vs {game.get('Team 2', 'Unknown')}': {e}")

    # Define the new column headers for the sharp plays and reverse line movements
    sharp_plays_columns = ['Bet', '% of Bets', '% of $', 'Sharp % Difference']
    reverse_line_movement_columns = ['Bet', 'Reverse Line Movement']

    # Create DataFrames with the updated columns
    sharp_plays_df = pd.DataFrame(sharp_plays_rows, columns=sharp_plays_columns)
    reverse_line_movement_df = pd.DataFrame(reverse_line_movement_rows, columns=reverse_line_movement_columns)

    # Sort the sharp plays DataFrame by the 'Sharp % Difference' column in descending order
    sharp_plays_df['Sharp % Difference Value'] = sharp_plays_df['Sharp % Difference'].str.extract(r'(\d+)').astype(int)
    sharp_plays_df = sharp_plays_df.sort_values(by='Sharp % Difference Value', ascending=False).drop('Sharp % Difference Value', axis=1)

    return sharp_plays_df, reverse_line_movement_df

#this is the actual web scrape function
def action_scrape():
    driver = None

    # First attempt: Auto-detection
    try:
        print("Attempting to create Chrome driver with auto-detection...")
        options = uc.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        driver = uc.Chrome(options=options, version_main=None)
        print("Driver created successfully with auto-detection!")
    except Exception as e:
        print(f"Auto-detection failed: {e}")

        # Second attempt: Explicit Chrome version 138
        try:
            print("Trying with explicit Chrome version 138...")
            options = uc.ChromeOptions()  # Create fresh options
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")

            driver = uc.Chrome(options=options, version_main=138)
            print("Driver created successfully with explicit version!")
        except Exception as e2:
            print(f"Explicit version failed: {e2}")

            # Third attempt: No version specification
            try:
                print("Trying without version specification...")
                options = uc.ChromeOptions()  # Create fresh options again
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")

                driver = uc.Chrome(options=options)
                print("Driver created successfully without version specification!")
            except Exception as e3:
                print(f"All undetected-chromedriver attempts failed: {e3}")
                print("Let's try with regular ChromeDriver...")

                # Fourth attempt: Regular ChromeDriver with webdriver-manager
                try:
                    from selenium.webdriver.chrome.service import Service
                    from webdriver_manager.chrome import ChromeDriverManager

                    options = webdriver.ChromeOptions()
                    options.add_argument("--disable-blink-features=AutomationControlled")
                    options.add_argument("--no-sandbox")
                    options.add_argument("--disable-dev-shm-usage")

                    service = Service(ChromeDriverManager().install())
                    driver = webdriver.Chrome(service=service, options=options)
                    print("Driver created successfully with regular ChromeDriver!")
                except Exception as e4:
                    print(f"Regular ChromeDriver also failed: {e4}")
                    return

    # Execute script to remove webdriver property
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    wait = WebDriverWait(driver, 10)
    try:
        driver.get("https://URL SITE HERE/public-betting")
        input("Please log in manually and then press Enter...")
        time.sleep(5)

        bet_type_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@data-testid='odds-tools-sub-nav__odds-type']//select")))

        select = Select(bet_type_dropdown)
        select.select_by_value("combined")
        time.sleep(3)

        game_rows = driver.find_elements(By.XPATH, "//table[@role='table']/tbody/tr")

        number_of_rows_per_game = 3
        all_game_data = []

        for i in range(0, len(game_rows), number_of_rows_per_game):
            rows = game_rows[i:i + number_of_rows_per_game]
            if len(rows) == number_of_rows_per_game:
                game_data = extract_game_data(rows)
                all_game_data.append(game_data)

        # Main DataFrame for detailed sheet
        df_main = construct_dataframe(all_game_data)

        # Separate DataFrames for summary sheet
        sharp_plays_df, reverse_line_movement_df = construct_summary_dataframes(all_game_data)

        output_file_path = "/Users/nicog-v/Documents/sharp_action_RLM.xlsx"

        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            df_main.to_excel(writer, index=False, startrow=1, sheet_name='Public Betting')

            workbook = writer.book
            worksheet_main = writer.sheets['Public Betting']

            # Apply the original formatting to the 'Public Betting' sheet
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'middle',
                'align': 'center',
                'bg_color': '#003366',
                'font_color': 'white',
                'border': 1
            })

            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
            current_date = pd.Timestamp.now().strftime('%Y-%m-%d')
            worksheet_main.merge_range('A1:U1', f'MLB Public Betting & Line Movement Analysis for {current_date}', title_format)

            for col_num, value in enumerate(df_main.columns.values):
                worksheet_main.write(1, col_num, value, header_format)

            worksheet_main.set_column('A:A', 25)
            worksheet_main.set_column('B:B', 20)
            worksheet_main.set_column('C:C', 22)
            worksheet_main.set_column('D:D', 15)
            worksheet_main.set_column('E:E', 15)
            worksheet_main.set_column('F:F', 15)
            worksheet_main.set_column('G:U', 25)

            table_range = f"A2:U{len(df_main)+2}"
            worksheet_main.conditional_format(table_range, {
                'type': 'no_blanks',
                'format': workbook.add_format({'border': 1})
            })
            worksheet_main.conditional_format(table_range, {
                'type': 'blanks',
                'format': workbook.add_format({'border': 1})
            })

            reverse_line_format = workbook.add_format({'bg_color': '#FF6666'})
            worksheet_main.conditional_format(
                f'G3:U{len(df_main)+2}',  # Apply from row 3 (after title and header) to end
                {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': 'Potential reverse line',
                    'format': reverse_line_format
                }
            )

            sharp_play_format = workbook.add_format({'bg_color': '#66FF66'})
            worksheet_main.conditional_format(
                f'G3:U{len(df_main)+2}',
                {
                    'type': 'text',
                    'criteria': 'containing',
                    'value': 'Potential sharp play',
                    'format': sharp_play_format
                }
            )

            # Concatenating 'Sharp Plays' and 'Reverse Line Movements' to write them side by side
            combined_df = pd.concat([sharp_plays_df, reverse_line_movement_df], axis=1)
            combined_df.to_excel(writer, index=False, sheet_name='Summary')

            worksheet_summary = writer.sheets['Summary']

            # Format headers for the combined data
            for col_num, value in enumerate(combined_df.columns.values):
                worksheet_summary.write(0, col_num, value, header_format)

            worksheet_summary.set_column(0, len(combined_df.columns) - 1, 25)

            # Apply borders to combined sheet
            table_range_summary = f"A2:{chr(65 + len(combined_df.columns) - 1)}{len(combined_df) + 1}"
            worksheet_summary.conditional_format(table_range_summary, {
                'type': 'no_blanks',
                'format': workbook.add_format({'border': 1})
            })
            worksheet_summary.conditional_format(table_range_summary, {
                'type': 'blanks',
                'format': workbook.add_format({'border': 1})
            })

    finally:
        if driver:
            driver.quit()

action_scrape()


