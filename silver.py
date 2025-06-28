import requests
from bs4 import BeautifulSoup
import re

def get_chennai_silver_rates():
    """
    Scrapes the GoodReturns website for silver rates in Chennai, using
    header matching to find the correct historical data table.

    Returns:
        A dictionary containing the structured data, or None on failure.
    """
    print("Fetching silver rates for Chennai from goodreturns.in...")
    
    try:
        url = 'https://www.goodreturns.in/silver-rates/chennai.html'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')
        
        # --- 1. Extract Today's Silver Prices (This part is correct) ---
        print("Extracting today's prices...")
        price_container = soup.find('div', class_='gold-rate-container')
        if not price_container:
            raise ValueError("Could not find the main price container 'gold-rate-container'.")
        
        todays_rates_divs = price_container.find_all('div', class_='gold-each-container')
        if len(todays_rates_divs) < 2:
            raise ValueError("Could not find enough price boxes for silver.")

        def parse_price_box(box):
            bottom_div = box.find('div', class_='gold-bottom')
            p_tags = bottom_div.find_all('p')
            price_str = p_tags[0].text if len(p_tags) > 0 else "0"
            change_str = p_tags[1].text if len(p_tags) > 1 else "0"
            price = float(re.sub(r'[₹,]', '', price_str))
            change_val = float(re.sub(r'[₹,+-]', '', change_str))
            if '-' in change_str:
                change_val = -change_val
            return {"price": price, "change": change_val}

        today_per_gram = parse_price_box(todays_rates_divs[0])
        today_per_kg = parse_price_box(todays_rates_divs[1])

        # --- 2. Extract Historical Data with Robust Table Finding ---
        print("Extracting historical data...")
        
        # ***** THE FIX IS HERE: Find all tables and identify the correct one by its headers *****
        all_tables = soup.find_all('table')
        history_table = None
        
        for table in all_tables:
            headers = [th.get_text(strip=True) for th in table.find_all('th')]
            # Check if this is the table we want
            if headers == ['Date', '10 gram', '100 gram', '1 Kg']:
                history_table = table
                print("Found correct historical silver data table.")
                break
        
        if not history_table:
            raise ValueError("Could not find the historical data table with the expected headers.")
        # *****************************************************************************************
            
        history_rows = history_table.find('tbody').find_all('tr')
        if not history_rows:
            raise ValueError("No rows found in history table")
        
        historical_data = []
        for row in history_rows:
            cols = row.find_all('td')
            if len(cols) == 4:
                date = cols[0].get_text(strip=True)
                price_10g = cols[1].get_text(strip=True)
                price_100g = cols[2].get_text(strip=True)
                price_1kg_full = ' '.join(cols[3].text.split())

                historical_data.append({
                    "date": date,
                    "price_10g": price_10g,
                    "price_100g": price_100g,
                    "price_1kg": price_1kg_full
                })
        
        # --- 3. Assemble the final data structure ---
        final_data = {
            "today_per_gram": today_per_gram,
            "today_per_kg": today_per_kg,
            "last_10_days": historical_data
        }
        
        print("Data extraction successful.")
        return final_data

    except Exception as e:
        print(f"An error occurred: {e}")
        return None
    

# --- Main Execution Block for Demonstration ---
if __name__ == "__main__":
    silver_data = get_chennai_silver_rates()

    if silver_data:
        print("\n" + "="*60)
        print(" " * 18 + "CHENNAI SILVER RATE REPORT")
        print("="*60)

        # Print Today's Data
        print("\n[+] Today's Silver Rate")
        
        price_g = silver_data['today_per_gram']['price']
        change_g = silver_data['today_per_gram']['change']
        arrow_g = "▲" if change_g >= 0 else "▼"
        print(f"  - Per Gram: ₹{price_g:,.2f} (Change: ₹{abs(change_g):.2f} {arrow_g})")

        price_kg = silver_data['today_per_kg']['price']
        change_kg = silver_data['today_per_kg']['change']
        arrow_kg = "▲" if change_kg >= 0 else "▼"
        print(f"  - Per Kg:   ₹{price_kg:,.0f} (Change: ₹{abs(change_kg):,.0f} {arrow_kg})")
        
        # Print Historical Data
        print("\n[+] Silver Rate in Chennai for Last 10 Days")
        header = f"{'Date':<15} | {'10 gram':<15} | {'100 gram':<15} | {'1 Kg'}"
        print("-" * len(header))
        print(header)
        print("-" * len(header))
        
        for record in silver_data['last_10_days']:
            print(f"{record['date']:<15} | {record['price_10g']:<15} | {record['price_100g']:<15} | {record['price_1kg']}")
        print("-" * len(header))