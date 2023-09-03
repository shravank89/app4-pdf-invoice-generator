import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extracting date and filename from filepath using pathlib
    filename = Path(filepath).stem
    path_info_list = filename.split("-")
    invoice_nr = path_info_list[0]
    date = path_info_list[1]

    # Writing to PDF
    pdf.set_font(family="Courier", size=15, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")

