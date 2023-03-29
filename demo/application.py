from flask import Flask, render_template, request, make_response
from img2table.document import PDF
from img2table.ocr import TesseractOCR
import pandas as pd
import os
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
import PyPDF2
import re

application = Flask(__name__)

@application.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file']
        filename = f.filename
        pdf = PDF(src=filename)

        # Instantiation of the OCR, Tesseract, which requires prior installation
        ocr = TesseractOCR(lang="eng")

        # Table identification and extraction
        pdf_tables = pdf.extract_tables(ocr=ocr)




        # We can also create an excel file with the tables
        pdf.to_xlsx('tables3.xlsx',
                    ocr=ocr)


        df = pd.read_excel('tables3.xlsx')

        pdf_file = open(filename, 'rb')

        # Create a PdfFileReader object
        pdf_reader = PyPDF2.PdfReader(pdf_file)

        page = pdf_reader.pages[0]
        # Extract the text
        text = page.extract_text()
        invoiceno = re.findall(r'((?<=Invoice No).{10})', text)
        ddate = re.findall(r'((?<=Document Date).{10})', text)
        cur = re.findall(r'((?<=Currency).{4})', text)
        df['Invoice No'] = invoiceno[0]
        df['Date'] = ddate[0]
        df['Currency'] = cur[0]
        # create two new columns by splitting the 'description' column
        df[['partno', 'description']] = df['Title & Description'].str.split(',', 1, expand=True)

        # strip any leading or trailing white space from the columns
        df['partno'] = df['partno'].str.strip()
        df['description'] = df['description'].str.strip()

        def convert_to_int(val):
            if val.isnumeric():
                return val
            else:
                return ''

        df['partno'] = df['partno'].apply(convert_to_int)

        df = df.drop('Title & Description', axis=1)
        # write the updated DataFrame to the Excel file
        df.to_excel('tables3.xlsx', index=False)



        print(f'invoice number is {invoiceno}')



        # Close the PDF file
        #pdf_file.close()

        # Extract the Quantity and Price columns
        quantity = df['Qty']
        price = df['Price']
        #description=df['Title & Description']
        #discount = df['Discount%']
        totalprice = df['Total']
        deliveryno = df['Delivery No']


        # Print the extracted data to the console
        print(f"Quantity: {quantity.tolist()}")
        print(f"Price: {price.tolist()}")
        qty = quantity.tolist()
        price1 = price.tolist()
        #description1=description.tolist()
        #discount1=discount.tolist()
        totalprice1 = totalprice.tolist()
        deliveryno2 = deliveryno.tolist()
        filename2 = 'tables3.xlsx'
        df.to_excel(filename2, index=False)

        # Return the Excel file as a download attachment
        with open(filename2, 'rb') as file:
            file_contents = file.read()
            response = make_response(file_contents)
            response.headers.set('Content-Type', 'application/vnd.ms-excel')
            response.headers.set('Content-Disposition', 'attachment', filename=filename2)
            return response


        return render_template('index.html', filename=qty,pp=price1,deliveryno3=deliveryno2,totalprice2=totalprice1,inv=invoiceno,dd=ddate,currency=cur)
    return render_template('index.html')

if __name__ == '__main__':
    application.run(debug=True)
