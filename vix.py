import os
import time
from PIL import Image
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ANSI color codes
GREEN = '\033[32m'
RED = '\033[31m'
RESET = '\033[0m'


def get_vix_data_and_chart(output_filename="india_vix_chart.png"):
    """
    Gets both VIX data and chart in a single browser session.
    """
    print("Fetching VIX data and chart...")
    driver = None
    try:
        # Initialize browser
        options = uc.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        
        # Ensure compatibility by using the correct ChromeDriver version
        driver = uc.Chrome(version_main=137, options=options, use_subprocess=True)
        page_url = 'https://groww.in/indices/india-vix'
        driver.get(page_url)
        
        wait = WebDriverWait(driver, 30)
        time.sleep(5)
        
        # First get the VIX data
        vix_data = None
        try:
            # Find VIX value
            table_rows = driver.find_elements(By.TAG_NAME, "tr")
            for row in table_rows:
                row_text = row.text
                if 'INDIA VIX' in row_text:
                    parts = row_text.split('\n') if '\n' in row_text else row_text.split()
                    for i, part in enumerate(parts):
                        if part == 'INDIA VIX':
                            current_value = float(parts[i+1])
                            break
            
            # Find Prev.Close and Open values
            all_text = driver.find_element(By.TAG_NAME, "body").text
            
            # Find Prev.Close
            prev_close_idx = all_text.find("Prev. Close")
            if prev_close_idx != -1:
                lines = all_text[prev_close_idx:].split('\n')
                for line in lines[:3]:
                    if line.replace('.', '').replace(',', '').strip().isdigit():
                        prev_close = float(line.strip())
                        break
            
            # Find Open
            open_idx = all_text.find("Open")
            if open_idx != -1:
                lines = all_text[open_idx:].split('\n')
                for line in lines[:3]:
                    if line.replace('.', '').replace(',', '').strip().isdigit():
                        open_value = float(line.strip())
                        break
            
            print(f"Prev Close: {prev_close}, Open: {open_value}")
            
            # Calculate changes using (Current Value - Open)/Open formula
            change_value = current_value - open_value
            change_percentage = (change_value / open_value) * 100
            vix_data = (current_value, change_value, change_percentage)
            
        except Exception as e:
            print(f"Error getting VIX data: {e}")
        
        # Then capture the chart
        chart_success = False
        try:
            # Switch to 1M view
            driver.execute_script("""
                document.querySelectorAll('.modal, .overlay').forEach(el => el.style.display = 'none');
                document.documentElement.setAttribute('data-theme', 'light');
            """)
            
            timeframe_buttons = driver.find_elements(By.XPATH, "//button[text()='1M']")
            if timeframe_buttons:
                driver.execute_script("arguments[0].click();", timeframe_buttons[0])
                time.sleep(2)
            
            # Take and crop screenshot
            temp_screenshot = "temp_full_page.png"
            driver.save_screenshot(temp_screenshot)
            
            img = Image.open(temp_screenshot)
            width, height = img.size
            
            left = width * 0.15
            right = width * 0.65
            top = height * 0.25
            bottom = height * 0.80
            
            chart_area = img.crop((left, top, right, bottom))
            chart_area.save(output_filename)
            
            os.remove(temp_screenshot)
            chart_success = True
            
        except Exception as e:
            print(f"Error capturing chart: {e}")
        
        return vix_data, chart_success
        
    except Exception as e:
        print(f"An error occurred: {e}")
        return None, False
        
    finally:
        if driver:
            driver.quit()

# --- Main Execution Block for Demonstration ---
if __name__ == "__main__":
    vix_data, chart_success = get_vix_data_and_chart()
    
    if vix_data:
        current_value, change_value, change_percentage = vix_data
        
        # Format the change with color
        if change_value > 0:
            change_text = f"{GREEN}+{change_value:.2f} (+{change_percentage:.2f}%){RESET}"
        else:
            change_text = f"{RED}{change_value:.2f} ({change_percentage:.2f}%){RESET}"
            
        print("\nIndia VIX Current Status")
        print("=" * 30)
        print(f"Current Value: {current_value:.2f}")
        print(f"Change: {change_text}")
        print("-" * 30)
        
        if chart_success:
            print("Chart capture was successful!")
        else:
            print("Chart capture failed.")
    else:
        print("Failed to get VIX data.")
