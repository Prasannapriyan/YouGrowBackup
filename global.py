import yfinance as yf
import pandas as pd

def get_global_indices_data():
    """
    Fetches the latest data for key global indices, now robustly handling
    potential missing data points from the source.

    Returns:
        A list of dictionaries, where each dictionary represents an index
        with its name, LTP, change, and percentage change. Returns None on failure.
    """
    print("Fetching data for key global indices...")

    indices = {
        "Dow Jones": "^DJI",
        "S&P 500": "^GSPC",
        "Nasdaq": "^IXIC",
        "FTSE 100": "^FTSE",
        "Hang Seng": "^HSI"
    }
    
    tickers_list = list(indices.values())

    try:
        # --- THE FIX IS HERE: Fetch a longer period to be safe ---
        # We fetch 5 days of data to have a buffer against missing data for one day.
        data = yf.download(tickers=tickers_list, period="5d", auto_adjust=True)

        if data.empty:
            print("Could not download indices data."); return None

        results = []
        ticker_to_name = {v: k for k, v in indices.items()}

        for ticker in tickers_list:
            try:
                # --- THE FIX IS HERE: Clean the data for each ticker ---
                # 1. Select the data for the current ticker
                ticker_data = data['Close'][ticker]
                # 2. Drop any rows with NaN (missing) values
                ticker_data_cleaned = ticker_data.dropna()
                
                # 3. Check if we have at least 2 valid data points left
                if len(ticker_data_cleaned) < 2:
                    print(f"  > Warning: Not enough valid data for {ticker} after cleaning. Skipping.")
                    continue
                
                # 4. Get the last two valid closing prices
                latest_close = ticker_data_cleaned.iloc[-1]
                prev_close = ticker_data_cleaned.iloc[-2]
                # -----------------------------------------------------

                change = latest_close - prev_close
                change_percent = (change / prev_close) * 100

                results.append({
                    "Name": ticker_to_name[ticker],
                    "LTP": latest_close,
                    "Change": change,
                    "Change %": change_percent
                })
            except (KeyError, IndexError):
                print(f"  > Warning: Could not process data for ticker {ticker}.")
                continue

        print("Successfully fetched and processed global indices data.")
        return results

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# --- Main Execution Block for Demonstration ---
# (This part of the code does not need to be changed)
if __name__ == "__main__":
    indices_data = get_global_indices_data()

    if indices_data:
        desired_order = ["Dow Jones", "Nasdaq", "S&P 500", "Hang Seng", "FTSE 100"]
        sorted_data = sorted(indices_data, key=lambda x: desired_order.index(x['Name']))

        print("\n" + "="*65)
        print(" " * 24 + "KEY GLOBAL INDICES")
        print("="*65)
        print(f"{'Name':<15} | {'LTP':>15} | {'Change':>12} | {'Change %':>12}")
        print("-" * 65)

        for index in sorted_data:
            ltp_str = f"{index['LTP']:,.2f}"
            change_str = f"{index['Change']:+.2f}"
            change_pct_str = f"{index['Change %']:+.2f}%"
            
            print(f"{index['Name']:<15} | {ltp_str:>15} | {change_str:>12} | {change_pct_str:>12}")
        
        print("-" * 65)