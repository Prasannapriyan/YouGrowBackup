import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from temp_bulletin import create_filtered_market_bulletin_livemint
from gold import get_chennai_gold_rates
from silver import get_chennai_silver_rates
import re
import os

def sanitize_text(text):
    return (
        text
        .replace("–", "-")
        .replace("—", "-")
        .replace("“", '"')
        .replace("”", '"')
        .replace("’", "'")
        .encode('latin-1', 'replace')
        .decode('latin-1')
    )

def clean_price(val):
    return int(re.sub(r'[^\d]', '', str(val))) if val else 0

def create_chart(df, columns, title, filename):
    plt.figure(figsize=(8, 4))
    for col in columns:
        plt.plot(df["Date"], df[col], marker="o", label=col)
    plt.title(title)
    plt.xlabel("Date")
    plt.ylabel("Price")
    plt.xticks(rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.set_text_color(0, 102, 204)
        self.cell(0, 10, "YouGrow Daily Market Report", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

    def add_table(self, title, df):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, title, ln=True)
        self.set_font("Arial", "", 10)
        col_width = self.w / (len(df.columns) + 1)
        self.set_fill_color(240)
        for col in df.columns:
            self.cell(col_width, 8, str(col), border=1, fill=True)
        self.ln()
        for _, row in df.iterrows():
            for item in row:
                txt = str(item).replace("₹", "Rs.")
                self.cell(col_width, 8, txt, border=1)
            self.ln()
        self.ln(5)

    def add_image(self, path, title=""):
        if title:
            self.set_font("Arial", "B", 12)
            self.cell(0, 10, title, ln=True)
        self.image(path, w=self.w - 40)
        self.ln(5)

    def add_bulletin(self, bulletin_text):
        self.add_page()
        self.set_font("Arial", 'B', 14)
        self.set_text_color(255, 87, 34)  # Electric Orange
        self.cell(0, 10, "Market Bulletin", ln=True)
        self.set_font("Arial", '', 12)
        self.set_text_color(0, 0, 0)

        for line in bulletin_text.splitlines():
            if line.strip():
                clean_line = sanitize_text(line)
                self.multi_cell(0, 8, f"- {clean_line}")

def main():
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # 1. Gold
    gold_data = get_chennai_gold_rates()
    gold_df = pd.DataFrame(gold_data["last_10_days"])
    gold_df.rename(columns={
        "date": "Date",
        "price_24k": "24K Price",
        "price_22k": "22K Price"
    }, inplace=True)
    gold_df["24K Clean"] = gold_df["24K Price"].apply(clean_price)
    gold_df["22K Clean"] = gold_df["22K Price"].apply(clean_price)
    create_chart(gold_df, ["24K Clean", "22K Clean"], "Gold Price Trend", "gold_chart.png")

    pdf.add_page()
    pdf.add_table("Gold Rates - Last 10 Days", gold_df[["Date", "24K Price", "22K Price"]])
    pdf.add_image("gold_chart.png", "Gold Rate Chart")

    # 2. Silver
    silver_data = get_chennai_silver_rates()
    silver_df = pd.DataFrame(silver_data["last_10_days"])
    silver_df.rename(columns={
        "date": "Date",
        "price_10g": "10 gram",
        "price_100g": "100 gram",
        "price_1kg": "1 Kg"
    }, inplace=True)
    silver_df["10g Clean"] = silver_df["10 gram"].apply(clean_price)
    silver_df["100g Clean"] = silver_df["100 gram"].apply(clean_price)
    create_chart(silver_df, ["10g Clean", "100g Clean"], "Silver Price Trend", "silver_chart.png")

    pdf.add_page()
    pdf.add_table("Silver Rates - Last 10 Days", silver_df[["Date", "10 gram", "100 gram", "1 Kg"]])
    pdf.add_image("silver_chart.png", "Silver Rate Chart")

    # 3. Market Bulletin
    bulletin_text = create_filtered_market_bulletin_livemint()
    pdf.add_bulletin(bulletin_text)

    # Export PDF
    pdf.output("YouGrow_Report_Prototype.pdf")

    # Clean up
    os.remove("gold_chart.png")
    os.remove("silver_chart.png")

    print("Report generated successfully: YouGrow_Report_Prototype.pdf")

if __name__ == "__main__":
    main()
