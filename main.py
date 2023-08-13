import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]
    date = filename.split("-")[1]
    date = date.replace(".", "-")

    pdf = FPDF(orientation='portrait', unit='mm', format='A4')
    pdf.add_page()

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Invoice - {invoice_number}", ln=1)

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f"Date: {date}")

    pdf.output(f'PDF-Invoices/{filename}.pdf')
