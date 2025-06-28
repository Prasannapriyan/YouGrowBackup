import os
import time
from PIL import Image
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def get_nifty_oi_data_and_chart(chart_filename="nifty_oi_chart.png"):
    """
    Scrapes the Upstox Nifty OI page, intelligently waiting for the data
    to populate before extracting it.

    Fetches:
    - Nifty Spot Price
    - Total Calls and Puts OI
    - Screenshots the OI chart and saves it as a PNG.

    Returns:
        A dictionary with the scraped data and chart filename, or None on failure.
    """
    print("Initializing browser to fetch Nifty OI data from Upstox...")
    driver = None
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1200")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        url = 'https://upstox.com/fno-discovery/open-interest-analysis/nifty-oi/'
        print(f"Navigating to {url}...")
        driver.get(url)

        print("Waiting for page content to load...")
        wait = WebDriverWait(driver, 30)
        
        # --- THE GUARANTEED FIX: Wait for REAL data to appear ---
        # This XPath finds the <p> tag for the Spot price value.
        spot_price_xpath = "//div[p[text()='Spot']]/p[2]"
        
        # We will wait until the text in this element is NOT '--/--'.
        # This is an explicit wait for the JavaScript data to load.
        print("Waiting for live Spot Price data to populate...")
        wait.until(
            lambda driver: driver.find_element(By.XPATH, spot_price_xpath).text != '--/--'
        )
        print("Live data detected.")
        # --------------------------------------------------------

        # Now that we know the data is real, we can extract it.
        summary_container = driver.find_element(By.CSS_SELECTOR, "div.gap-8.py-4")
        data_sections = summary_container.find_elements(By.CSS_SELECTOR, "div.flex-col")
        
        extracted_data = {}
        for section in data_sections:
            p_tags = section.find_elements(By.TAG_NAME, 'p')
            if len(p_tags) == 2:
                label = p_tags[0].text
                value = p_tags[1].text
                extracted_data[label] = value

        spot_price_str = extracted_data.get('Spot', '0')
        total_calls_oi = extracted_data.get('Total Calls', '0 L')
        total_puts_oi = extracted_data.get('Total Puts', '0 L')
        
        spot_price = float(spot_price_str.replace(',', ''))
        
        print("Numerical data extracted successfully.")
        
        # --- Screenshot the Chart ---
        print("Waiting for chart element...")
        chart_container_selector = "div.recharts-responsive-container"
        chart_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, chart_container_selector))
        )
        
        time.sleep(2) # Allow for final render
        
        # Take full page screenshot
        print("Taking full page screenshot...")
        temp_screenshot = "temp_full_page.png"
        driver.save_screenshot(temp_screenshot)
        
        # Open and crop the image
        img = Image.open(temp_screenshot)
        width, height = img.size
        
        # Calculate crop dimensions for the chart
        left = width * 0.15    # 15% from left
        right = width * 0.67   # 67% from left
        top = height * 0.20    # 20% from top
        bottom = height * 0.72  # 72% from top
        
        # Crop and save
        chart_area = img.crop((left, top, right, bottom))
        chart_area.save(chart_filename)
        
        # Clean up temporary file
        os.remove(temp_screenshot)
        print(f"Chart saved to '{chart_filename}'")
        
        final_data = {
            "spot_price": spot_price,
            "total_calls_oi": total_calls_oi,
            "total_puts_oi": total_puts_oi,
            "chart_filepath": chart_filename
        }
        
        return final_data

    except Exception as e:
        print(f"An error occurred: {e}")
        return None
    finally:
        if driver:
            print("Closing browser.")
            driver.quit()

# --- Main Execution Block ---
if __name__ == "__main__":
    oi_data = get_nifty_oi_data_and_chart()
    
    if oi_data:
        print("\n" + "="*50)
        print("          NIFTY OPEN INTEREST ANALYSIS")
        print("="*50)
        print(f"  Spot Price:      {oi_data['spot_price']:,.2f}")
        print(f"  Total Call OI:   {oi_data['total_calls_oi']}")
        print(f"  Total Put OI:    {oi_data['total_puts_oi']}")
        #print(f"  Chart saved to:  '{oi_data['chart_filepath']}'")
        print("="*50)
    else:
        print("\nFailed to retrieve Nifty OI data.")
