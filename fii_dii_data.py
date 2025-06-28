import os
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from PIL import Image
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def generate_fii_dii_summary():
    def arrow(val: float) -> str:
        return f"{val:,.2f} {'🟢↑' if val >= 0 else '🔴↓'}"

    def fetch_all_pages(min_unique_days=15, max_pages=4):
        all_data = []
        page = 1
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Connection": "keep-alive",
        }
        print("\nFetching FII/DII data...")

        while page <= max_pages:
            url = "https://groww.in/fii-dii-data" if page == 1 else f"https://groww.in/fii-dii-data?page={page}"
            print(f"Fetching page {page}...")
            try:
                resp = requests.get(url, headers=headers, timeout=10)
                resp.raise_for_status()
            except requests.RequestException as e:
                print(f"Error fetching page {page}: {str(e)}")
                break
            soup = BeautifulSoup(resp.text, "html.parser")
            table = soup.find("table")
            if not table:
                print(f"No table found on page {page}")
                break
            
            print(f"Processing data from page {page}...")
            rows = table.select("tbody tr")
            for row in rows:
                cols = [td.get_text(strip=True).replace(",", "") for td in row.find_all("td")]
                if len(cols) >= 7:
                    date_str = cols[0]
                    fii_net = float(cols[3].replace("+", "").replace("−", "-").replace("–", "-"))
                    dii_net = float(cols[6].replace("+", "").replace("−", "-").replace("–", "-"))
                    parsed_date = datetime.strptime(date_str, "%d %b %Y")
                    all_data.append((parsed_date, fii_net, dii_net))
            page += 1

        df = pd.DataFrame(all_data, columns=["Date", "FII_Net", "DII_Net"])
        df = df.drop_duplicates(subset=['Date'], keep='first')
        df = df.sort_values("Date", ascending=False).reset_index(drop=True)
        
        if len(df) < min_unique_days:
            print(f"\nWarning: Could only fetch {len(df)} days of data, needed {min_unique_days}")
        print(f"\nTotal trading days collected: {len(df)}")
        print("\nAll dates in data:")
        for date in df['Date']:
            print(date.strftime("%d-%m-%Y"))
        return df

    df = fetch_all_pages(min_unique_days=10)
    results = []

    # Last 3 individual days
    for i in range(3):
        row = df.iloc[i]
        date_label = row['Date'].strftime("%d-%m-%Y")
        desc = "Last" if i == 0 else f"{i+1}ᵗʰ"
        label = f"{date_label}"
        results.append({
            "Date": label,
            "FII": arrow(row['FII_Net']),
            "DII": arrow(row['DII_Net'])
        })

    # Last 7 days cumulative
    last_7 = df.head(7)
    results.append({
        "Date": "Last 7 Days",
        "FII": arrow(last_7['FII_Net'].sum()),
        "DII": arrow(last_7['DII_Net'].sum())
    })

    # Last 10 days
    last_10 = df.head(10)
    results.append({
        "Date": "Last 10 Days",
        "FII": arrow(last_10['FII_Net'].sum()),
        "DII": arrow(last_10['DII_Net'].sum())
    })

    return results

def get_fii_dii_chart(output_filename="fii_dii_chart.png"):
    """
    Launches a stealth browser, navigates to the Groww FII/DII data page,
    and takes a screenshot of the main activity chart.
    """
    print("Initializing stealth browser to capture FII/DII chart...")
    
    driver = None
    try:
        options = uc.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1200")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        driver = uc.Chrome(version_main=137, options=options, use_subprocess=True)
        url = 'https://groww.in/fii-dii-data'
        
        print(f"Navigating to {url}...")
        driver.get(url)
        
        print("Waiting for the chart element to load...")
        wait = WebDriverWait(driver, 30)
        chart_element_selector = "svg.recharts-surface"
        
        chart_element = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, chart_element_selector))
        )
        
        time.sleep(3)
        print("Chart element found. Taking screenshot...")
        
        print("Changing page zoom level to 80%...")
        driver.execute_script("document.body.style.zoom='80%'")
        time.sleep(2)

        temp_screenshot = "temp_full_page.png"
        driver.save_screenshot(temp_screenshot)
        
        img = Image.open(temp_screenshot)
        width, height = img.size
        
        left = width * 0.25
        right = width * 0.75
        top = height * 0.35
        bottom = height * 0.75
        
        chart_area = img.crop((left, top, right, bottom))
        chart_area.save(output_filename)
        
        os.remove(temp_screenshot)
        print(f"Successfully saved chart to {output_filename}")
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        return False
        
    finally:
        if driver:
            print("Closing browser.")
            driver.quit()

def main():
    try:
        print("\n=== FII/DII Data and Chart Generation ===")
        
        # Generate FII/DII Summary
        print("\nGenerating FII/DII summary...")
        summary = generate_fii_dii_summary()
        
        if summary:
            print("\nFII/DII Summary:")
            print("=" * 80)
            for row in summary:
                print(f"{row['Date']}: FII = {row['FII']}, DII = {row['DII']}")
            print("=" * 80)
        else:
            print("Failed to generate FII/DII summary.")

        # Capture FII/DII Chart
        print("\nCapturing FII/DII chart...")
        if get_fii_dii_chart():
            print("\nChart generation completed successfully!")
        else:
            print("\nChart generation failed.")

    except Exception as e:
        print(f"\nError occurred: {type(e).__name__}")
        print(f"Error details: {str(e)}")
        print("\nPlease check your internet connection and try again.")

if __name__ == "__main__":
    main()
