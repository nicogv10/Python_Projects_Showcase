from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from bs4 import BeautifulSoup
import time

def parse_insights(html_content):
    """
    Parse the HTML content for betting insights.
    Extracts team, game, trend, bet, hit rate, and odds information.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    insights = []

    insight_elements = soup.find_all('div', class_='[ trending-insight ] flex flex-col group')

    for element in insight_elements:
        try:
            match_info_div = element.find('div', class_='[ match-info ] flex flex-col grow')
            if not match_info_div:
                continue

            team_name_div = match_info_div.find_all('div', recursive=False)
            if not team_name_div:
                continue
            team = team_name_div[0].text.strip()  # Not used later, but kept for clarity

            game_span = match_info_div.find('span', {'data-sentry-component': 'FormatEventLabel'})
            if not game_span:
                continue
            game = game_span.text.strip()

            date_time_span = match_info_div.find('span', text=lambda t: ' AM' in t or ' PM' in t)
            date_time = date_time_span.text.strip() if date_time_span else "No date/time"

            trend_div = element.find('div', class_='headline-2 bold')
            trend = trend_div.text.strip() if trend_div else "No trend"

            bet_link = element.find('a', {'data-sentry-component': 'InsightViewLink'})
            bet = bet_link.text.strip() if bet_link else "No bet"

            hit_rate_div = element.find('div', class_='caption-2 bold text-accent-green-50b')
            hit_rate = hit_rate_div.text.strip() if hit_rate_div else "No hit rate"

            odds_span = element.find('button', class_='group').find('span')
            odds = odds_span.text.strip() if odds_span else "No odds"

            insights.append({
                'Trend': trend,
                'Bet': bet,
                'Hit_Rate': hit_rate,
                'Odds': odds,
            })
        except Exception as e:
            print(f"Failed to parse insight: {e}")
            continue

    return insights

def calculate_implied_probability_and_gain(df):
    """
    Calculate implied probability from odds and compare to hit rate.
    Adds a 'Gain' column showing the difference between hit rate and implied probability.
    """
    df['Odds'] = df['Odds'].astype(float, errors='ignore')
    df['Implied Probability'] = df['Odds'].apply(lambda x: (-x / (-x + 100)) if x < 0 else (100 / (x + 100)))
    df['Hit_Rate'] = df['Hit_Rate'].str.rstrip('%').astype(float) / 100
    df['Gain (Hit Rate - Implied Probability)'] = df['Hit_Rate'] - df['Implied Probability']
    df = df[df['Odds'] >= -225]
    df = df.sort_values(by='Gain (Hit Rate - Implied Probability)', ascending=False)

    # Keep only top 12 bets
    df = df[['Trend', 'Bet', 'Hit_Rate', 'Odds', 'Gain (Hit Rate - Implied Probability)']].head(12)
    return df

def scrape_insights():
    """
    Launches a browser session, navigates to the insights page (URL required),
    waits for manual login, and scrapes insights game by game.
    Outputs results to Excel with formatting applied.
    """
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)

    all_insights = []

    try:
        # Replace this with the actual URL when running locally
        driver.get("SCRAPE_URL_HERE")
        input("Please log in manually and then press Enter...")
        time.sleep(5)

        try:
            # Example of applying filters — may vary depending on page layout
            insight_type_button_xpath = "//button[@data-flip-id='insightType' and contains(@class, 'flex')]"
            insight_type_button = wait.until(EC.element_to_be_clickable((By.XPATH, insight_type_button_xpath)))
            insight_type_button.click()
            time.sleep(2)

            team_option_xpath = "//div[normalize-space()='Team' and contains(@class, 'truncate')]"
            team_option = wait.until(EC.element_to_be_clickable((By.XPATH, team_option_xpath)))
            team_option.click()
            time.sleep(2)

            done_button_xpath = "//button[div[normalize-space()='Done']]"
            done_button = wait.until(EC.element_to_be_clickable((By.XPATH, done_button_xpath)))
            done_button.click()
            time.sleep(2)
        except Exception as e:
            print(f"Filter application error: {e}")

        while True:
            proceed = input("Manually filter the next game, then press Enter to scrape (or type 'done' to finish): ")
            if proceed.lower() == 'done':
                print("Finished scraping games.")
                break

            try:
                html_content = driver.page_source
                insights = parse_insights(html_content)
                all_insights.extend(insights)
                print(f"Scraped {len(insights)} insights from current game.\n")
            except Exception as e:
                print(f"An error occurred while scraping current game: {e}")

    finally:
        driver.quit()

    if all_insights:
        df = pd.DataFrame(all_insights)
        df = calculate_implied_probability_and_gain(df)

        with pd.ExcelWriter("output/TOP12_insights.xlsx", engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=1, sheet_name='Top Bets')

            workbook = writer.book
            worksheet = writer.sheets['Top Bets']

            current_date = pd.Timestamp.now().strftime('%Y-%m-%d')
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
            worksheet.merge_range('A1:E1', f'Top 12 Insights for {current_date}', title_format)

            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'middle',
                'align': 'center',
                'bg_color': '#003366',
                'font_color': 'white',
                'border': 1
            })
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(1, col_num, value, header_format)

            worksheet.set_column('A:A', 32)  # Trend
            worksheet.set_column('B:B', 24)  # Bet
            worksheet.set_column('C:C', 12, workbook.add_format({'num_format': '0%'}))  # Hit Rate
            worksheet.set_column('D:D', 10)  # Odds
            worksheet.set_column('E:E', 24, workbook.add_format({'num_format': '0.00'}))  # Gain

            data_range = f"A2:E{len(df) + 2}"
            worksheet.conditional_format(data_range, {
                'type': 'no_blanks',
                'format': workbook.add_format({'border': 1})
            })

            worksheet.conditional_format(
                f'E2:E{len(df) + 2}',
                {
                    'type': 'data_bar',
                    'bar_color': '#FF9900'
                }
            )

        print("Top 12 condensed insights have been exported to output/TOP12_insights.xlsx")
    else:
        print("No data to export.")

if __name__ == "__main__":
    scrape_insights()
