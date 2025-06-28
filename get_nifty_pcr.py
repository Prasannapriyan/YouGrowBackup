import os
import time
from PIL import Image
import pandas as pd
from nsepython import nse_optionchain_scrapper
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

HISTORY_FILE = "pcr_history.csv"
CHART_URL = "https://upstox.com/fno-discovery/open-interest-analysis/nifty-pcr/"
SCREENSHOT_FILE = "nifty_pcr_chart.png"
CROPPED_SCREENSHOT = "nifty_pcr_chart_cropped.png"

def get_nifty_pcr_and_history():
    """
    Fetches the latest NIFTY PCR value, displays the last two sessions from history,
    and takes a screenshot of the PCR chart from a website.
    """
    # --- Fetch and display current PCR ---
    try:
        print("Fetching latest NIFTY option chain data...")
        chain = nse_optionchain_scrapper('NIFTY')
        if not chain:
            print("Could not fetch NIFTY option chain data.")
        else:
            total_pe_oi = sum(float(strike.get('PE', {}).get('openInterest', 0))
                             for strike in chain['records']['data'] if 'PE' in strike)
            total_ce_oi = sum(float(strike.get('CE', {}).get('openInterest', 0))
                             for strike in chain['records']['data'] if 'CE' in strike)

            if total_ce_oi == 0:
                print("Warning: Total Call OI is zero, cannot calculate PCR.")
            else:
                current_pcr = total_pe_oi / total_ce_oi
                today_str = datetime.now().strftime('%d-%m-%Y')
                print(f"\nNIFTY PCR for the latest trading session ({today_str}): {current_pcr:.2f}")

    except Exception as e:
        print(f"An error occurred while fetching live NIFTY PCR: {e}")

    # --- Read and display historical data ---
    if os.path.exists(HISTORY_FILE):
        print("\n--- Last Two Trading Sessions ---")
        try:
            history_df = pd.read_csv(HISTORY_FILE, parse_dates=['Date'])
            history_df = history_df.sort_values('Date', ascending=False)
            
            today_date = pd.to_datetime(datetime.now().date())
            
            # Exclude today's data to get previous sessions
            previous_sessions = history_df[history_df['Date'] < today_date].head(2)

            if not previous_sessions.empty:
                for index, row in previous_sessions.iterrows():
                    date_str = row['Date'].strftime('%d-%m-%Y')
                    pcr_val = row['PCR']
                    print(f"Date: {date_str}, PCR: {pcr_val:.2f}")
            else:
                print("Not enough historical data to show previous sessions.")

        except Exception as e:
            print(f"An error occurred while reading PCR history: {e}")
    else:
        print(f"\nHistory file '{HISTORY_FILE}' not found. Cannot display previous sessions.")

def capture_pcr_chart():
    """
    Captures a screenshot of the NIFTY PCR chart from the specified URL.
    """
    print(f"\nCapturing PCR chart from: {CHART_URL}")
    driver = None
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--no-sandbox")
        
        driver = webdriver.Chrome(options=chrome_options)
        driver.get(CHART_URL)
        
        # Wait for the chart container to be present
        wait = WebDriverWait(driver, 20)
        print("Waiting for chart to load...")
        
        # Wait for the main container and ensure chart is loaded
        wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, "recharts-responsive-container"))
        )
        
        # Set a larger window size for better chart visibility
        driver.set_window_size(1920, 1080)
        
        # Hide any overlays and set theme
        driver.execute_script("""
            // Hide any overlays or popups
            document.querySelectorAll('.modal, .overlay').forEach(el => el.style.display = 'none');
            // Set theme
            document.documentElement.setAttribute('data-theme', 'light');
        """)
        
        # Give chart time to stabilize
        time.sleep(3)
        
        # Take full page screenshot
        temp_screenshot = "temp_full_page.png"
        driver.save_screenshot(temp_screenshot)
        
        # Open and crop the image
        img = Image.open(temp_screenshot)
        width, height = img.size
        
        # Calculate crop dimensions for PCR chart
        left = width * 0.15    # 20% from left
        right = width * 0.67   # 80% from left
        top = height * 0.30    # 15% from top
        bottom = height * 0.81  # 55% from top
        
        # Crop and save
        chart_area = img.crop((left, top, right, bottom))
        chart_area.save(CROPPED_SCREENSHOT)
        
        # Clean up temporary file
        os.remove(temp_screenshot)
        
        print(f"Chart screenshot saved to '{CROPPED_SCREENSHOT}'")
        return True
        
    except Exception as e:
        print(f"An error occurred while capturing the chart: {e}")
        return False
        
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    get_nifty_pcr_and_history()
    capture_pcr_chart()
