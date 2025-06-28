import time
import io
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image

def get_sgx_nifty_snapshot(output_filename="sgx_nifty_snapshot.png"):
    """
    Launches a browser, screenshots a stable container element, and then
    programmatically crops the image to the exact required section,
    excluding the chart and ads at the bottom.

    Args:
        output_filename (str): The name of the output PNG file.

    Returns:
        bool: True if successful, False otherwise.
    """
    print("Initializing browser to capture SGX Nifty snapshot...")
    
    driver = None
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--window-size=1280,1000")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        
        url = 'https://sgxnifty.org/'
        
        print(f"Navigating to {url}...")
        driver.get(url)
        
        wait = WebDriverWait(driver, 30)
        
        # --- 1. Find the main container that holds all desired content ---
        container_selector = ".grid.col-940.align-center"
        print("Waiting for the main data container to load...")
        container_element = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, container_selector))
        )
        time.sleep(2) # Allow dynamic values to load

        # --- 2. Screenshot just this container ---
        print("Taking initial screenshot of the container...")
        screenshot_data = container_element.screenshot_as_png

        # --- 3. Find the LAST element we want to KEEP ---
        # The second 'main-table-div' is the last piece of content we want.
        tables = container_element.find_elements(By.CLASS_NAME, "main-table-div")
        if len(tables) < 2:
            raise Exception("Could not find both data tables for cropping.")
            
        last_element_to_keep = tables[1]
        
        # --- 4. Calculate the precise crop height ---
        # The height is its y-position relative to the container plus its own height.
        # Subtracting a padding of 130 pixels for a clean bottom edge.
        crop_height = last_element_to_keep.location['y'] + last_element_to_keep.size['height'] - 130
        
        # --- 5. Crop and Save ---
        print(f"Cropping image to a calculated height of {crop_height} pixels...")
        img = Image.open(io.BytesIO(screenshot_data))
        
        # Crop from top-left (0,0) to the full width and the calculated height
        crop_box = (0, 0, img.width, crop_height)
        cropped_img = img.crop(crop_box)

        cropped_img.save(output_filename, "PNG")

        print(f"Successfully saved cropped snapshot to {output_filename}")
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        return False
        
    finally:
        if driver:
            print("Closing browser.")
            driver.quit()

# --- Main Execution Block ---
if __name__ == "__main__":
    success = get_sgx_nifty_snapshot()
    if success:
        print("\nSGX Nifty snapshot capture was successful!")
    else:
        print("\nSGX Nifty snapshot capture failed.")