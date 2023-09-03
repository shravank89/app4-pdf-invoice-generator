import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    # Extracting date and filename from filepath using pathlib
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Writing to PDF
    pdf.set_font(family="Courier", size=15, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_nr}", ln=1)

    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Iterate over pandas df.columns object and making header
    columns = [x.replace("_", " ").title() for x in df.columns]
    pdf.set_font(family="Courier", size=10, style="B")
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=80, h=8, txt=columns[1], border=1)
    pdf.cell(w=50, h=8, txt=columns[2], border=1)
    pdf.cell(w=40, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Adding data from pandas
    for index, row in df.iterrows():
        pdf.set_font(family="Courier", size=10)
        pdf.set_text_color(100,100,100)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=80, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1,ln=1)

    # Adding total Sum cell
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Courier", size=10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=80, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Adding total payable line
    pdf.set_font(family="Courier", size=15,style="B")
    pdf.cell(w=75, h=8, txt=f"Total Payable is: {str(total_sum)}", ln=1)
    pdf.cell(w=75, h=8, txt=f"Cricket International")
    pdf.image("image.png", w=40)

    pdf.output(f"PDFs/{filename}.pdf")
