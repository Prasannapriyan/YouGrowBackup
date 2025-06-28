import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import traceback

def get_tradingview_heatmap_price(output_filename="stock_heatmap_price.png"):
    """
    Launches a VISIBLE stealth browser and automates the clicks to change the
    heatmap view to 'Price' with robust error handling and multiple fallback methods.
    """
    print("Initializing VISIBLE stealth browser...")
    
    driver = None
    try:
        options = uc.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-popup-blocking")
        
        # Ensure compatibility by using the correct ChromeDriver version
        driver = uc.Chrome(version_main=137, options=options, use_subprocess=True)
        wait = WebDriverWait(driver, 30)
        
        # Load page
        url = 'https://in.tradingview.com/heatmap/stock/#%7B%22dataSource%22%3A%22NIFTY50%22%2C%22blockColor%22%3A%22change%22%2C%22blockSize%22%3A%22market_cap_basic%22%2C%22grouping%22%3A%22no_group%22%7D'
        print(f"Navigating to TradingView...")
        driver.get(url)
        time.sleep(5)
        
        # --- 1. Click Fullscreen Button ---
        print("Waiting for the fullscreen button...")
        fullscreen_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-qa-id='heatmap-top-bar_fullscreen']"))
        )
        fullscreen_button.click()
        time.sleep(3)

        # --- 2. Click Display Settings Button ---
        print("Waiting for the display settings button...")
        settings_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-qa-id='heatmap-top-bar_settings-button']"))
        )
        settings_button.click()
        time.sleep(2)
        
        def verify_price_selected():
            """Verify if Price is actually selected"""
            try:
                return driver.execute_script("""
                    const spans = document.querySelectorAll('span[class*="nestedSlotItem"]');
                    for (const span of spans) {
                        if (span.textContent === 'Price') return true;
                    }
                    return false;
                """)
            except:
                return False

        def select_dropdown():
            """Try to find and click the Display Value dropdown"""
            try:
                # First try by XPath
                dropdown = driver.find_element(By.XPATH, "//div[div[text()='Display value']]/following-sibling::div[1]")
                dropdown.click()
                return True
            except:
                try:
                    # Try JavaScript fallback
                    return driver.execute_script("""
                        const el = Array.from(document.querySelectorAll('*')).find(el => 
                            el.textContent === 'Display value' && 
                            (el.closest('[role="button"]') || el.parentElement)
                        );
                        if (el) {
                            el.click();
                            return true;
                        }
                        return false;
                    """)
                except:
                    return False

        def select_price_option():
            """Try to select the Price option using various methods"""
            # Try JavaScript click
            try:
                success = driver.execute_script("""
                    const options = document.querySelectorAll('[role="option"], [role="menuitem"]');
                    const priceOption = Array.from(options).find(el => el.textContent === 'Price');
                    if (priceOption) {
                        priceOption.click();
                        return true;
                    }
                    return false;
                """)
                if success:
                    return True
            except:
                pass

            # Try keyboard navigation
            try:
                actions = ActionChains(driver)
                for _ in range(2):
                    actions.send_keys(Keys.ARROW_DOWN).perform()
                    time.sleep(0.5)
                actions.send_keys(Keys.RETURN).perform()
                time.sleep(1)
                if verify_price_selected():
                    return True
            except:
                pass

            # Try force update
            try:
                success = driver.execute_script("""
                    const spans = document.querySelectorAll('span[class*="nestedSlotItem"]');
                    for (const span of spans) {
                        if (span.textContent.includes('Change') || span.textContent.includes('Symbol')) {
                            span.textContent = 'Price';
                            span.dispatchEvent(new Event('change', { bubbles: true }));
                            return true;
                        }
                    }
                    return false;
                """)
                if success:
                    return True
            except:
                pass

            return False

        # --- 3. Try to select Price multiple times ---
        success = False
        for attempt in range(3):
            if attempt > 0:
                print(f"\nRetrying Price selection (attempt {attempt + 1}/3)...")
                try:
                    settings_button.click()
                    time.sleep(1)
                except:
                    pass

            print(f"Opening Display Value dropdown...")
            if select_dropdown():
                time.sleep(1)
                print("Selecting Price option...")
                if select_price_option():
                    if verify_price_selected():
                        success = True
                        print("Successfully selected Price!")
                        break
            time.sleep(1)

        if not success:
            print("Warning: Could not confirm Price selection")

        # Function to verify if dialog is closed
        def is_dialog_closed():
            try:
                # Check if any dialog elements are visible
                dialog_elements = driver.find_elements(By.CSS_SELECTOR, 
                    "[data-dialog-name], .dialog-container, .tv-dialog")
                for element in dialog_elements:
                    if element.is_displayed():
                        return False
                return True
            except:
                return False

        # Function to attempt dialog close
        def try_close_dialog():
            methods = [
                # Method 1: Click close button
                lambda: driver.find_element(By.CSS_SELECTOR, "[data-name='close']").click(),
                
                # Method 2: Click settings button again
                lambda: settings_button.click(),
                
                # Method 3: Use ESC key
                lambda: ActionChains(driver).send_keys(Keys.ESCAPE).perform(),
                
                # Method 4: JavaScript close
                lambda: driver.execute_script("""
                    const dialogs = document.querySelectorAll('[data-dialog-name], .dialog-container');
                    dialogs.forEach(dialog => {
                        dialog.style.display = 'none';
                        dialog.remove();
                    });
                """)
            ]
            
            for method in methods:
                try:
                    method()
                    time.sleep(0.5)
                    if is_dialog_closed():
                        return True
                except:
                    continue
            return False

        print("Closing settings dialog...")
        for attempt in range(3):
            if try_close_dialog():
                print("Successfully closed settings dialog")
                break
            else:
                print(f"Dialog close attempt {attempt + 1} failed, retrying...")
        
        # Verify dialog is closed
        if not is_dialog_closed():
            print("Warning: Settings dialog might still be visible")
            
        # Additional wait for UI to settle
        time.sleep(2)
        
        # Configure color depth
        print("Setting color depth...")
        try:
            # Find and click color depth button
            color_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-qa-id='heatmap-bottom-bar_button_multiplier']"))
            )
            color_button.click()
            print("Clicked color depth button")
            time.sleep(1)

            # Navigate using keyboard
            print("Navigating color depth options...")
            actions = ActionChains(driver)
            # Press down arrow three times
            for _ in range(3):
                actions.send_keys(Keys.ARROW_DOWN).perform()
                time.sleep(0.5)
            
            # Press Enter to select
            actions.send_keys(Keys.RETURN).perform()
            print("Selected color depth option")
            
            # Move mouse away and click in blank area to remove highlight
            print("Removing button highlight...")
            actions.move_by_offset(200, 0).click().perform()
        except Exception as e:
            print(f"Warning: Color depth selection failed: {e}")
        
        time.sleep(1)  # Brief pause before screenshot
        
        # Take screenshot
        print("Taking screenshot...")
        driver.save_screenshot(output_filename)
        print(f"Screenshot saved to {output_filename}")
        return True

    except Exception as e:
        print(f"An error occurred: {e}")
        traceback.print_exc()
        return False
    
    finally:
        if driver:
            try:
                driver.quit()
                print("Browser closed successfully")
            except:
                print("Warning: Could not close browser cleanly")

if __name__ == "__main__":
    success = get_tradingview_heatmap_price()
    if success:
        print("\nHeatmap capture was successful!")
    else:
        print("\nHeatmap capture failed.")
