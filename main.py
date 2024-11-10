from fpdf import FPDF
from pathlib import Path
import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, text=f"Invoice nr.{invoice_nr}", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, text=f"Date: {date}", new_x="LMARGIN", new_y="NEXT")

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = df.columns
    columns = [item.replace("_", "").title() for item in columns]
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, text=columns[0], border=1)
    pdf.cell(w=70, h=8, text=columns[1], border=1)
    pdf.cell(w=30, h=8, text=columns[2], border=1)
    pdf.cell(w=30, h=8, text=columns[3], border=1)
    pdf.cell(w=30, h=8, text=columns[4], border=1, new_x="LMARGIN", new_y="NEXT")

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, text=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, text=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, text=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, text=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, text=str(row["total_price"]), border=1, new_x="LMARGIN", new_y="NEXT")

    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(w=70, h=8, text="", border=1)
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(w=30, h=8, text="", border=1)
    pdf.cell(w=30, h=8, text=str(total_sum), border=1, new_x="LMARGIN", new_y="NEXT")

    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, text=f"The total price is {total_sum}", new_x="LMARGIN", new_y="NEXT")
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, text=f"Pythonhow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
