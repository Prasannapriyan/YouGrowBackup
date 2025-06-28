import yfinance as yf

def get_currency_exchange_rates():
    """
    Fetches the latest exchange rates for popular currencies against the INR.

    Returns:
        A list of dictionaries, where each dictionary represents a currency
        with its name, country, and value against INR. Returns None on failure.
    """
    print("Fetching latest currency exchange rates against INR...")

    # --- THE FIX IS HERE: Replaced KWD with AED ---
    # A dictionary mapping currency codes to their full name, country, and yfinance ticker
    currencies = {
        "USD": {"name": "US Dollar", "country": "United States", "ticker": "USDINR=X"},
        "EUR": {"name": "Euro", "country": "Eurozone", "ticker": "EURINR=X"},
        "GBP": {"name": "British Pound", "country": "United Kingdom", "ticker": "GBPINR=X"},
        "SGD": {"name": "Singapore Dollar", "country": "Singapore", "ticker": "SGDINR=X"},
        "JPY": {"name": "Japanese Yen", "country": "Japan", "ticker": "JPYINR=X"},
        "AED": {"name": "UAE Dirham", "country": "U.A.E.", "ticker": "AEDINR=X"},
    }
    # ----------------------------------------------------
    
    results = []
    
    try:
        # Loop through each currency defined in our dictionary
        for code, details in currencies.items():
            try:
                print(f"  > Fetching {details['name']} ({code})...")
                
                ticker = yf.Ticker(details['ticker'])
                info = ticker.info
                
                value = info.get('regularMarketPrice') or info.get('previousClose')
                
                if value is None:
                    print(f"    - Could not find price for {code}. Skipping.")
                    continue

                results.append({
                    "Code": code,
                    "Name": details['name'],
                    "Country": details['country'],
                    "Value": value
                })

            except Exception as e:
                print(f"    - Failed to fetch data for {code}: {e}. Skipping.")
                continue

        if not results:
            print("Could not fetch any currency data.")
            return None

        print("Successfully fetched currency data.")
        return results

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

# --- Main Execution Block for Demonstration ---
if __name__ == "__main__":
    currency_data = get_currency_exchange_rates()

    if currency_data:
        # Define the desired order for the final output
        desired_order = ["USD", "JPY", "EUR", "SGD", "GBP", "AED"]
        
        data_map = {item['Code']: item for item in currency_data}
        sorted_data = [data_map[code] for code in desired_order if code in data_map]

        print("\n" + "="*70)
        print(" " * 22 + "POPULAR CURRENCIES vs. INR")
        print("="*70)
        print(f"{'Code':<5} | {'Country':<25} | {'Value (1 unit in INR)'}")
        print("-" * 70)

        for currency in sorted_data:
            code_str = currency['Code']
            country_str = currency['Country']
            value_str = f"₹{currency['Value']:.2f}"
            
            print(f"{code_str:<5} | {country_str:<25} | {value_str}")
        
        print("-" * 70)