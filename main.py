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
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Adding the Invoice table column names

    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family='Times', size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Adding columns data in the invoice

    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # Adding last row with total Invoice price
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(df['total_price'].sum()), border=1, ln=1)

    # Add total bill
    bill = f"Your total price is  {df['total_price'].sum()}$"
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=0, h=8, txt=bill, ln=1)

    # Add logo to look professional :)
    pdf.set_font(family='Times', size=14, style="B")
    pdf.cell(w=30, h=8, txt="Pug Society")
    pdf.image('Akica.jpg', w=10)

    pdf.output(f'PDF-Invoices/{filename}.pdf')
