import requests
from bs4 import BeautifulSoup
import re

def get_chennai_gold_rates():
    """
    Scrapes the GoodReturns website for gold rates in Chennai using the
    correct HTML selectors for all sections.

    Fetches:
    - Today's 24K and 22K price and change.
    - Last 10 days of historical data for 24K and 22K gold.

    Returns:
        A dictionary containing the structured data, or None on failure.
    """
    print("Fetching gold rates for Chennai from goodreturns.in...")
    
    try:
        url = 'https://www.goodreturns.in/gold-rates/chennai.html'
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')
        
        # --- 1. Extract Today's Gold Prices ---
        print("Extracting today's prices...")
        price_container = soup.find('div', class_='gold-rate-container')
        if not price_container:
            raise ValueError("Could not find the main price container 'gold-rate-container'.")
        
        todays_rates_divs = price_container.find_all('div', class_='gold-each-container')
        if len(todays_rates_divs) < 2:
            raise ValueError("Could not find enough price boxes for 24K and 22K gold.")

        def parse_price_box(box):
            bottom_div = box.find('div', class_='gold-bottom')
            p_tags = bottom_div.find_all('p')
            price_str = p_tags[0].text if len(p_tags) > 0 else "0"
            change_str = p_tags[1].text if len(p_tags) > 1 else "0"
            price = float(re.sub(r'[₹,]', '', price_str))
            try:
                # Handle negative values properly
                change_str = re.sub(r'[₹,]', '', change_str).strip()
                # First remove all spaces
                change_str = ''.join(change_str.split())
                # Handle plus/minus signs
                if change_str.startswith('+'):
                    change_str = change_str[1:]
                elif change_str.startswith('- ') or change_str.startswith('-'):
                    change_str = f"-{change_str.replace('-', '').strip()}"
                print(f"Debug - Change string after cleanup: '{change_str}'")  # Debug output
                change = float(change_str)
            except ValueError as e:
                print(f"Debug - Error parsing change value: '{change_str}', Error: {e}")
                change = 0.0  # Default to 0 if parsing fails
            return {"price": price, "change": change}

        today_24k = parse_price_box(todays_rates_divs[0])
        today_22k = parse_price_box(todays_rates_divs[1])

        # --- 2. Extract Historical Data ---
        print("Extracting historical data...")
        
        # Find the historical data table
        print("Looking for tables...")
        tables = soup.find_all('table')
        print(f"Found {len(tables)} tables")
        
        # Debug print all tables
        '''for i, table in enumerate(tables):
            print(f"\nTable {i+1} classes:", table.get('class', []))
            print(f"Table {i+1} HTML:")
            print(table.prettify()[:200] + "...\n")'''
        
        # Find the correct table (look for the one with Date, 24K, 22K headers)
        history_table = None
        for table in tables:
            headers = [th.get_text(strip=True) for th in table.find_all('th')]
            print(f"Found headers: {headers}")
            if 'Date' in headers and '24K' in headers and '22K' in headers:
                history_table = table
                print("Found target table with correct headers")
                break
        
        if not history_table:
            raise ValueError("Could not find the historical data table")
            
        print("\nExtracting rows...")
        history_rows = history_table.find_all('tr')
        if not history_rows:
            raise ValueError("No rows found in table")
            
        print(f"Found {len(history_rows)} rows")
        
        # Skip header row if present
        if history_rows[0].find('th'):
            history_rows = history_rows[1:]
            print(f"Skipped header row, processing {len(history_rows)} data rows")
        
        historical_data = []
        for row in history_rows:
            cols = row.find_all('td')
            if len(cols) == 3:
                # Get date
                date = cols[0].text.strip()
                
                # Process 24K data
                td_24k = cols[1]
                cell_content = td_24k.get_text(strip=True)
                
                # Split into price and change values
                price_24k = cell_content.split('(')[0].strip() if '(' in cell_content else cell_content
                
                # Then get the change value from span
                change_span_24k = td_24k.find('span')
                change_24k = ""
                change_color_24k = ""
                if change_span_24k:
                    change_text = change_span_24k.text.strip()
                    change_24k = change_text.strip('()')
                    # Remove any existing plus signs to avoid duplicates
                    change_24k = change_24k.replace('+', '')
                    if 'green-span' in change_span_24k.get('class', []):
                        change_color_24k = 'green'
                    elif 'red-span' in change_span_24k.get('class', []):
                        change_color_24k = 'red'
                
                # Process 22K data
                td_22k = cols[2]
                cell_content = td_22k.get_text(strip=True)
                
                # Split into price and change values
                price_22k = cell_content.split('(')[0].strip() if '(' in cell_content else cell_content
                
                # Then get the change value from span
                change_span_22k = td_22k.find('span')
                change_22k = ""
                change_color_22k = ""
                if change_span_22k:
                    change_text = change_span_22k.text.strip()
                    change_22k = change_text.strip('()')
                    # Remove any existing plus signs to avoid duplicates
                    change_22k = change_22k.replace('+', '')
                    if 'green-span' in change_span_22k.get('class', []):
                        change_color_22k = 'green'
                    elif 'red-span' in change_span_22k.get('class', []):
                        change_color_22k = 'red'
                
                # Add only if we have valid data
                if date and price_24k and price_22k:
                    historical_data.append({
                    "date": date,
                    "price_24k": price_24k,
                    "change_24k": change_24k,
                    "change_color_24k": change_color_24k,
                    "price_22k": price_22k,
                    "change_22k": change_22k,
                    "change_color_22k": change_color_22k
                })
        
        # --- 3. Assemble the final data structure ---
        final_data = {
            "today_24k": today_24k,
            "today_22k": today_22k,
            "last_10_days": historical_data
        }
        
        print("Data extraction successful.")
        return final_data

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# --- Main Execution Block for Demonstration ---
# (This part of the code does not need to be changed)
if __name__ == "__main__":
    gold_data = get_chennai_gold_rates()

    if gold_data:
        print("\n" + "="*60)
        print(" " * 18 + "CHENNAI GOLD RATE REPORT")
        print("="*60)

        # Print Today's Data
        print("\n[+] Today's Gold Rate (per gram)")
        
        price_24k = gold_data['today_24k']['price']
        change_24k = gold_data['today_24k']['change']
        arrow_24k = "▲" if change_24k >= 0 else "▼"
        print(f"  - 24K Gold: ₹{price_24k:,.0f} (Change: ₹{abs(change_24k):.0f} {arrow_24k})")

        price_22k = gold_data['today_22k']['price']
        change_22k = gold_data['today_22k']['change']
        arrow_22k = "▲" if change_22k >= 0 else "▼"
        print(f"  - 22K Gold: ₹{price_22k:,.0f} (Change: ₹{abs(change_22k):.0f} {arrow_22k})")
        
        # --- THE FIX IS HERE: Corrected Table Printing ---
        # Define column widths and spacing
        col1_width = 16  # Date
        col2_width = 42  # 24K
        col3_width = 42  # 22K
        divider = "-" * 104  # Adjusted divider length
        
        # Print section header with space above
        print("\n[+] Gold Rate in Chennai for Last 10 Days (1 gram)\n")
        
        # Print table header
        header = f"{'Date':<{col1_width}} | {'24K Price':<{col2_width}} | {'22K Price':<{col3_width}}"
        print(divider)
        print(header)
        print(divider)
        
        # ANSI color codes
        GREEN = '\033[32m'
        RED = '\033[31m'
        RESET = '\033[0m'
        
        # Loop through the records and print them with correct padding
        for record in gold_data['last_10_days']:
            # Format 24K price with change
            price_24k = record['price_24k']
            change_24k = record['change_24k']
            base_text = f"{price_24k}  "
            if change_24k:
                if change_24k == "0":
                    change_text = f"({change_24k})"
                elif record['change_color_24k'] == 'green':
                    change_text = f"({GREEN}+{change_24k}{RESET})"
                elif record['change_color_24k'] == 'red':
                    change_text = f"({RED}-{change_24k.lstrip('-')}{RESET})"
                else:
                    change_text = f"({change_24k})"
            else:
                change_text = ""
            
            # Calculate padding based on visible length (excluding ANSI codes)
            base_length = len(base_text)
            change_length = len(change_text.replace(GREEN, "").replace(RED, "").replace(RESET, ""))
            visible_length = base_length + change_length
            
            # Apply consistent formatting for all values
            if change_24k == "0":
                # Format zero values with consistent spacing
                price_24k_full = f"{price_24k}  ({change_24k})"  # Same format as non-zero values
                padding_length = 35 - len(price_24k_full)
                padding = " " * max(0, padding_length)
                price_24k_full = base_text + change_text + padding
            else:
                padding_length = 35 - visible_length
                padding = " " * max(0, padding_length)
                price_24k_full = base_text + change_text + padding
            
            # Format 22K price with change
            price_22k = record['price_22k']
            change_22k = record['change_22k']
            base_text = f"{price_22k}  "
            if change_22k:
                if change_22k == "0":
                    change_text = f"({change_22k})"
                elif record['change_color_22k'] == 'green':
                    change_text = f"({GREEN}+{change_22k}{RESET})"
                elif record['change_color_22k'] == 'red':
                    change_text = f"({RED}-{change_22k.lstrip('-')}{RESET})"
                else:
                    change_text = f"({change_22k})"
            else:
                change_text = ""
            
            # Calculate padding based on visible length (excluding ANSI codes)
            base_length = len(base_text)
            change_length = len(change_text.replace(GREEN, "").replace(RED, "").replace(RESET, ""))
            visible_length = base_length + change_length
            
            # Apply consistent formatting for all values
            if change_22k == "0":
                # Format zero values with consistent spacing
                price_22k_full = f"{price_22k}  ({change_22k})"  # Same format as non-zero values
                padding_length = 35 - len(price_22k_full)
                padding = " " * max(0, padding_length)
                price_22k_full = base_text + change_text + padding
            else:
                padding_length = 35 - visible_length
                padding = " " * max(0, padding_length)
                price_22k_full = base_text + change_text + padding
            
            # Print the formatted line
            print(f"{record['date']:<{col1_width}} | {price_24k_full:<{col2_width}} | {price_22k_full:<{col3_width}}")
        
        print(divider)
        print()  # Add extra line at the end
        # ---------------------------------------------------
