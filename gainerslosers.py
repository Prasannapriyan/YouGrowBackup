import yfinance as yf
import pandas as pd

def get_nifty50_movers():
    """
    Fetches the top 5 gainers and top 5 losers from the NIFTY 50 index
    for the last trading session.

    Returns:
        A dictionary containing two lists: 'gainers' and 'losers'.
        Returns None if data fetching fails.
    """
    print("Fetching Nifty 50 constituents' data...")
    
    # A reasonably recent list of Nifty 50 tickers. 
    # For a production system, this should be fetched dynamically.
    nifty50_tickers = [
        "ADANIENT.NS", "ADANIPORTS.NS", "APOLLOHOSP.NS", "ASIANPAINT.NS", 
        "AXISBANK.NS", "BAJAJ-AUTO.NS", "BAJFINANCE.NS", "BAJAJFINSV.NS", 
        "BPCL.NS", "BHARTIARTL.NS", "BRITANNIA.NS", "CIPLA.NS", "COALINDIA.NS", 
        "DIVISLAB.NS", "DRREDDY.NS", "EICHERMOT.NS", "GRASIM.NS", "HCLTECH.NS", 
        "HDFCBANK.NS", "HDFCLIFE.NS", "HEROMOTOCO.NS", "HINDALCO.NS", 
        "HINDUNILVR.NS", "ICICIBANK.NS", "ITC.NS", "INDUSINDBK.NS", "INFY.NS", 
        "JSWSTEEL.NS", "KOTAKBANK.NS", "LTIM.NS", "LT.NS", "M&M.NS", 
        "MARUTI.NS", "NTPC.NS", "NESTLEIND.NS", "ONGC.NS", "POWERGRID.NS", 
        "RELIANCE.NS", "SBILIFE.NS", "SBIN.NS", "SUNPHARMA.NS", "TCS.NS", 
        "TATACONSUM.NS", "TATAMOTORS.NS", "TATASTEEL.NS", "TECHM.NS", 
        "TITAN.NS", "ULTRACEMCO.NS", "WIPRO.NS", "SHRIRAMFIN.NS"
    ]

    try:
        # Download data for the last 2 trading days for all tickers at once
        data = yf.download(nifty50_tickers, period="2d", auto_adjust=True)

        if data.empty:
            print("Could not download stock data. Check tickers or network.")
            return None

        # We only need the 'Close' prices
        close_prices = data['Close']
        
        if len(close_prices) < 2:
            print("Not enough data to calculate change (less than 2 trading days).")
            return None

        # Get the latest and previous closing prices
        latest_prices = close_prices.iloc[-1]
        previous_prices = close_prices.iloc[-2]

        # Calculate percentage change
        percent_change = ((latest_prices - previous_prices) / previous_prices) * 100
        
        # Create a list of dictionaries to hold the results
        results = []
        for ticker in percent_change.index:
            # Skip if data is missing for a particular stock
            if pd.notna(percent_change[ticker]) and pd.notna(latest_prices[ticker]):
                results.append({
                    "stock": ticker.replace(".NS", ""),
                    "price": latest_prices[ticker],
                    "change": percent_change[ticker]
                })

        # Sort by percentage change
        sorted_results = sorted(results, key=lambda x: x['change'], reverse=True)
        
        # Get top 5 gainers and losers
        top_gainers = sorted_results[:5]
        top_losers = sorted_results[-5:]
        # Reverse the losers list to show the biggest loser first
        top_losers.reverse()

        print("Successfully processed data.")
        results = {
            "gainers": top_gainers,
            "losers": top_losers
        }

        # Print gainers and losers in a table format
        print("\nTop 5 Gainers:")
        print("=" * 40)
        print(f"{'Stock':<15}{'Price':<12}{'Change (%)':<10}")
        print("-" * 40)
        for gainer in results['gainers']:
            print(f"{gainer['stock']:<15}{gainer['price']:<12.2f}+{gainer['change']:.2f}%")
        
        print("\nTop 5 Losers:")
        print("=" * 40)
        print(f"{'Stock':<15}{'Price':<12}{'Change (%)':<10}")
        print("-" * 40)
        for loser in results['losers']:
            print(f"{loser['stock']:<15}{loser['price']:<12.2f}{loser['change']:.2f}%")

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# --- Main Execution Block ---
if __name__ == "__main__":
    nifty_movers = get_nifty50_movers()
    if nifty_movers:
        print("\nNIFTY 50 Movers Report")
        print("=" * 40)
        print(f"{'Category':<15}{'Stock':<12}{'Change (%)':<10}")
        print("-" * 40)
        for category, stocks in nifty_movers.items():
            print(f"\n{category.capitalize()}:")
            for stock in stocks:
                print(f"{stock['stock']:<15}{stock['price']:<12.2f}{stock['change']:<10.2f}")
    else:
        print("Failed to fetch NIFTY 50 movers data.")
