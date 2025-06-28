import yfinance as yf
import pandas as pd
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, black, grey, white, lightgrey

# --- Configuration ---
NIFTY_TICKER = "^NSEI"
PDF_FILE = "Market_Report_Dashboard_Nifty50.pdf"

# --- 1. Enhanced Data Fetching (No changes) ---
def get_nifty_dashboard_data():
    print("Fetching enhanced Nifty 50 data...")
    ticker = yf.Ticker(NIFTY_TICKER)
    hist_1y = ticker.history(period="1y")
    hist_2d = ticker.history(period="2d")

    if hist_1y.empty or hist_2d.empty:
        print("Could not download data. Check ticker or internet connection.")
        return None

    latest_day = hist_2d.iloc[-1]
    prev_day = hist_2d.iloc[-2]

    data = {
        "current_price": latest_day['Close'],
        "prev_close": prev_day['Close'],
        "open": latest_day['Open'],
        "intraday_high": latest_day['High'],
        "intraday_low": latest_day['Low'],
        "volume_lakhs": latest_day['Volume'] / 100000,
        "fifty_two_week_high": hist_1y['High'].max(),
        "fifty_two_week_low": hist_1y['Low'].min(),
    }
    data['change'] = data['current_price'] - data['prev_close']
    data['change_percent'] = (data['change'] / data['prev_close']) * 100
    print("Data fetching complete.")
    return data

# --- 2. PDF Generation Helper Functions ---
def draw_data_point(c, x, y, label, value, value_color):
    c.setFont("Helvetica", 9)
    c.setFillColor(grey)
    c.drawString(x, y, label)
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(value_color)
    c.drawString(x, y - 18, value)

def draw_slider_refined(c, x, y, width, height, label, low_val, high_val, current_val):
    # ***** THE FIX IS HERE *****
    # Save the current graphics state before we start clipping
    c.saveState()
    # ***************************

    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(black)
    c.drawString(x, y, label)

    y_bar = y - 35
    bar_radius = height / 2

    path = c.beginPath()
    path.roundRect(x, y_bar - bar_radius, width, height, bar_radius)
    c.clipPath(path, stroke=0, fill=0)
    c.linearGradient(x, y_bar, x + width, y_bar, (HexColor("#d90429"), HexColor("#f8b24f"), HexColor("#8ac926")), extend=False)

    # ***** AND HERE *****
    # Restore the graphics state to remove the clipping path
    c.restoreState()
    # ********************

    range_val = high_val - low_val
    position_ratio = (current_val - low_val) / range_val if range_val != 0 else 0.5
    position_ratio = max(0, min(1, position_ratio))
    marker_x = x + width * position_ratio

    label_text = f"{current_val:,.2f}"
    label_width = c.stringWidth(label_text, "Helvetica", 9) + 10
    label_height = 14
    label_x = marker_x - (label_width / 2)
    label_y = y_bar + 10

    c.setFillColor(white)
    c.setStrokeColor(lightgrey)
    c.roundRect(label_x, label_y, label_width, label_height, label_height/2, stroke=1, fill=1)
    
    c.setFillColor(black)
    c.setFont("Helvetica", 9)
    c.drawCentredString(marker_x, label_y + 4, label_text)

    c.setFillColor(white)
    c.setStrokeColor(grey)
    c.setLineWidth(1)
    c.circle(marker_x, y_bar, 6, stroke=1, fill=1)

    y_lowhigh = y_bar - 25
    c.setFont("Helvetica", 9)
    c.setFillColor(grey)
    c.drawString(x, y_lowhigh, "Low")
    c.drawRightString(x + width, y_lowhigh, "High")

    c.setFont("Helvetica", 12)
    c.setFillColor(black)
    c.drawString(x, y_lowhigh - 15, f"{low_val:,.2f}")
    c.drawRightString(x + width, y_lowhigh - 15, f"{high_val:,.2f}")


# --- 3. Main PDF Generation Function ---
def generate_stylish_pdf_report(data):
    if not data:
        print("No data to generate report.")
        return

    print(f"Creating PDF report: {PDF_FILE}...")
    c = canvas.Canvas(PDF_FILE, pagesize=(8.5*inch, 5*inch))
    width, height = (8.5*inch, 5*inch)
    
    # ***** FIX: Removed the redundant saveState() from here *****
    
    color_indigo = HexColor("#4f46e5")
    color_green = HexColor("#16a34a")
    color_red = HexColor("#dc2626")

    c.setFont("Helvetica-Bold", 18)
    c.setFillColor(black)
    c.drawString(0.5 * inch, height - 0.7 * inch, "NIFTY 50")

    c.setFont("Helvetica-Bold", 42)
    c.drawString(0.5 * inch, height - 1.3 * inch, f"{data['current_price']:,.2f}")

    x_change = 3.5 * inch
    y_change = height - 1.2 * inch
    
    if data['change'] >= 0:
        c.setFillColor(color_green)
        change_text = f"▲ {data['change']:.2f} ({data['change_percent']:.2f}%)"
    else:
        c.setFillColor(color_red)
        change_text = f"▼ {abs(data['change']):.2f} ({abs(data['change_percent']):.2f}%)"
    
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x_change, y_change, change_text)
    
    y_grid = height - 2.1 * inch
    draw_data_point(c, 0.5 * inch, y_grid, "Prev. Close", f"{data['prev_close']:,.2f}", color_indigo)
    draw_data_point(c, 2.2 * inch, y_grid, "Open", f"{data['open']:,.2f}", color_indigo)
    draw_data_point(c, 3.9 * inch, y_grid, "Volume (Lakhs)", f"{data['volume_lakhs']:,.2f}", color_indigo)

    slider_y = height - 3.2 * inch
    slider_width = 3.5 * inch
    slider_height = 10

    draw_slider_refined(c, 0.5 * inch, slider_y, slider_width, slider_height, "52 Week",
                data['fifty_two_week_low'], data['fifty_two_week_high'], data['current_price'])
    
    draw_slider_refined(c, 4.5 * inch, slider_y, slider_width, slider_height, "Intraday",
                data['intraday_low'], data['intraday_high'], data['current_price'])

    c.save()
    print("PDF report generated successfully.")


# --- Main Execution ---
if __name__ == "__main__":
    dashboard_data = get_nifty_dashboard_data()
    if dashboard_data:
        generate_stylish_pdf_report(dashboard_data)
        print(f"\nReport '{PDF_FILE}' is ready.")