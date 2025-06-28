import yfinance as yf
import pandas as pd
import mplfinance as mpf
import os

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph

# --- Configuration ---
NIFTY_TICKER = "^NSEI"
CHART_FILE = "nifty_1hr_chart.png"
PDF_FILE = "Market_Report_Segment_Nifty50.pdf"

# --- 1. Data Fetching and Analysis ---
def get_nifty_data():
    """
    Fetches Nifty 50 data and prepares it for analysis.
    This version includes the definitive fix for multi-level column headers.
    """
    print("Fetching Nifty 50 data...")
    # Download data
    nifty_daily_data = yf.download(NIFTY_TICKER, period="90d", interval="1d", auto_adjust=True)
    nifty_hourly_data = yf.download(NIFTY_TICKER, period="15d", interval="1h", auto_adjust=True)

    # ***** DEFINITIVE FIX 1: FLATTEN MULTI-LEVEL COLUMNS *****
    # yfinance can return multi-level columns e.g., ('Open', '^NSEI').
    # We flatten them to a single level e.g., 'Open' to ensure simple access.
    if isinstance(nifty_daily_data.columns, pd.MultiIndex):
        nifty_daily_data.columns = nifty_daily_data.columns.get_level_values(0)
    if isinstance(nifty_hourly_data.columns, pd.MultiIndex):
        nifty_hourly_data.columns = nifty_hourly_data.columns.get_level_values(0)
    # **********************************************************

    if nifty_daily_data.empty or nifty_hourly_data.empty:
        print("Could not download data. Check ticker or internet connection.")
        return None

    # --- Analysis ---
    nifty_daily_data['SMA50'] = nifty_daily_data['Close'].rolling(window=50).mean()
    latest_close = nifty_daily_data['Close'].iloc[-1]
    latest_sma50 = nifty_daily_data['SMA50'].iloc[-1]
    
    recent_period = nifty_daily_data.tail(15)
    
    # ***** DEFINITIVE FIX 2: SILENCE FUTUREWARNING *****
    # Use .iloc[0] to get the scalar value, as recommended by the warning.
    resistance_level = float(recent_period['High'].max())
    support_level = float(recent_period['Low'].min())
    # Note: The above float() is often sufficient, but for absolute certainty,
    # the most robust way (though more verbose) would be:
    # resistance_level = recent_period['High'].max().iloc[0] if isinstance(recent_period['High'].max(), pd.Series) else recent_period['High'].max()
    # support_level = recent_period['Low'].min().iloc[0] if isinstance(recent_period['Low'].min(), pd.Series) else recent_period['Low'].min()
    # For now, the simple float() cast after flattening columns will work correctly.
    
    analysis_data = {
        "daily_data": nifty_daily_data,
        "hourly_data": nifty_hourly_data,
        "latest_close": latest_close,
        "latest_sma50": latest_sma50,
        "resistance_1": resistance_level,
        "resistance_2": resistance_level * 1.005,
        "support_1": support_level,
        "support_2": support_level * 0.995,
    }
    
    print("Data analysis complete.")
    return analysis_data

# --- 2. Chart Generation (Now works with clean data) ---
def create_nifty_chart(data):
    if not data:
        return False

    hourly_df = data['hourly_data'].tail(7*8).copy()
    
    # This cleaning step is now redundant because the data is clean at the source,
    # but it's good practice to keep it as a safeguard.
    cols_to_check = ['Open', 'High', 'Low', 'Close']
    for col in cols_to_check:
        hourly_df[col] = pd.to_numeric(hourly_df[col], errors='coerce')
    hourly_df.dropna(inplace=True)

    if hourly_df.empty:
        print("No valid hourly data to plot after cleaning.")
        return False

    hlines = dict(hlines=[data['resistance_1'], data['support_1']], colors=['r', 'g'], linestyle='--')
    mc = mpf.make_marketcolors(up='#00b746', down='#ef403c', inherit=True)
    style = mpf.make_mpf_style(marketcolors=mc, base_mpf_style='nightclouds')

    print(f"Generating chart and saving to {CHART_FILE}...")
    try:
        mpf.plot(
            hourly_df, type='candle', style=style, title='Nifty 50 - 1 Hour Chart',
            ylabel='Price (INR)', volume=False, hlines=hlines, figratio=(16, 9),
            savefig=dict(fname=CHART_FILE, dpi=150, pad_inches=0.1)
        )
        print("Chart saved successfully.")
        return True
    except Exception as e:
        print(f"An error occurred during chart plotting: {e}")
        return False

# --- 3. PDF Generation (No changes needed) ---
def generate_pdf_report(data):
    if not data:
        print("No data to generate report.")
        return

    print(f"Creating PDF report: {PDF_FILE}...")
    c = canvas.Canvas(PDF_FILE, pagesize=letter)
    width, height = letter

    title_color = HexColor("#1E3A8A")
    text_color = HexColor("#1F2937")
    resistance_color = HexColor("#B91C1C")
    support_color = HexColor("#15803D")

    styles = getSampleStyleSheet()
    body_style = styles['BodyText']
    body_style.textColor = text_color
    body_style.fontSize = 11
    body_style.leading = 14
    
    c.setFillColor(title_color)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(1 * inch, height - 1 * inch, "Nifty Technical Analysis")
    c.line(1*inch, height - 1.05*inch, width - 1*inch, height - 1.05*inch)

    if os.path.exists(CHART_FILE):
        chart_width = 6.5 * inch
        chart_height = (chart_width / 16) * 9
        c.drawImage(CHART_FILE, 1 * inch, height - 1.5 * inch - chart_height, width=chart_width, height=chart_height)
        y_position = height - 1.8 * inch - chart_height
    else:
        c.drawString(1*inch, height - 1.7*inch, "Chart could not be generated.")
        y_position = height - 2.2 * inch

    c.setFillColor(title_color)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(1 * inch, y_position, "◆ Overview")
    y_position -= 0.3 * inch

    daily_analysis_text = f"""
    <b>Daily Chart Analysis:</b> Nifty is currently trading around {data['latest_close']:.2f}, which is 
    {'above' if data['latest_close'] > data['latest_sma50'] else 'below'} the key 50-day SMA of 
    {data['latest_sma50']:.2f}. This indicates a {'bullish' if data['latest_close'] > data['latest_sma50'] else 'bearish'}
    medium-term trend. The index has shown {'strength' if data['latest_close'] > data['resistance_1']*0.98 else 'weakness'} 
    in recent sessions.
    """
    hourly_analysis_text = f"""
    <b>1-Hour Chart Analysis:</b> The short-term chart shows the price action consolidating near the recent 
    highs. A breakout above the immediate resistance at {data['resistance_1']:.2f} could trigger further 
    upside momentum, while a failure to hold the support at {data['support_1']:.2f} may lead to a pullback.
    """
    
    p = Paragraph(daily_analysis_text, body_style)
    p.wrapOn(c, width - 2*inch, height)
    p.drawOn(c, 1*inch, y_position - p.height)
    y_position -= (p.height + 0.2*inch)

    p = Paragraph(hourly_analysis_text, body_style)
    p.wrapOn(c, width - 2*inch, height)
    p.drawOn(c, 1*inch, y_position - p.height)
    y_position -= (p.height + 0.3*inch)

    c.setFillColor(title_color)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(1 * inch, y_position, "◆ Key Levels to Watch")
    y_position -= 0.3 * inch

    c.setFillColor(resistance_color)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(1.1 * inch, y_position, "RESISTANCE")
    c.setFont("Helvetica", 11)
    y_position -= 0.25 * inch
    c.drawString(1.2 * inch, y_position, f"● {data['resistance_1']:.2f} – Immediate resistance from recent highs.")
    y_position -= 0.25 * inch
    c.drawString(1.2 * inch, y_position, f"● {data['resistance_2']:.2f} – Next potential resistance zone.")
    y_position -= 0.4 * inch

    c.setFillColor(support_color)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(1.1 * inch, y_position, "SUPPORT")
    c.setFont("Helvetica", 11)
    y_position -= 0.25 * inch
    c.drawString(1.2 * inch, y_position, f"● {data['support_1']:.2f} – Immediate support from recent lows.")
    y_position -= 0.25 * inch
    c.drawString(1.2 * inch, y_position, f"● {data['support_2']:.2f} – Next strong support zone.")

    c.save()
    print("PDF report generated successfully.")
    
    if os.path.exists(CHART_FILE):
        os.remove(CHART_FILE)
        print(f"Cleaned up temporary file: {CHART_FILE}")

# --- Main Execution ---
if __name__ == "__main__":
    nifty_analysis = get_nifty_data()
    if nifty_analysis:
        chart_created = create_nifty_chart(nifty_analysis)
        if chart_created:
            generate_pdf_report(nifty_analysis)
            print(f"\nReport '{PDF_FILE}' is ready.")
        else:
            print("\nCould not not not generate the report because the chart creation failed.")