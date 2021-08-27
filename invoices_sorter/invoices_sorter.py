'''I have an excel file with two columns ("Billing Document" and " Customer Name"),
and folder full of PDF documents containing the billing document number in their filenames.
This script compares all these filenames with the excel file and finds the Customer Name
corresponding to the PDF. Then it moves the PDF to a new directory named with the Customer Name.

This code doesn't work without proper files. Sensitive data was changed.
 '''

import pandas as pd
import os, shutil

# Read excel file
file = "C:/Path/To/File/invoices_all.XLSX"
df = pd.read_excel(file)
# Check for NaN cells
df['Billing Document'] = df['Billing Document'].apply(lambda x: x if pd.isnull(x) else str(int(x)))

directory = r"C:/Path/To/Invoices/"

def sort_invoices():
    for item in os.listdir(directory):
        invoice = item[0:10] # PDF's first 10 digits are the Billing Document number from the excel file
        name = df.loc[df['Billing Document']==invoice, 'Customer Name'].values[0]
        print(name)
        dir_path = directory + name
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)
        else:
            file_path = directory + item
            shutil.move(file_path, dir_path)

sort_invoices()
