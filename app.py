import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

required_files = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    directory = single_input("Please enter the directory containing the ProductData.xlsx, MasterInvoice.pdf, and Order files.", width=600)
    directory = directory.replace('"', '')

    if directory.endswith('\\') == False:
        directory = directory + '\\'

    try:
        files_in_directory = os.listdir(directory)

        all_files_found = True
        for file in required_files:
            if file not in files_in_directory:
                all_files_found = False

        if all_files_found == True:
            break
        else:
            show_message("Error", "The directory you entered does not contain all of the required files. Please try again.")
    except:
        show_message("Error", "The directory you entered does not exist. Please try again.")

descriptions = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')

product_descriptions = {}
for index,row in descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')

product_prices = {}
for index,row in prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

def get_order_data(invoice_num, filename, directory):

    order_form = openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet = order_form['OrderForm']

    date = selected_sheet['C9'].value.strftime('%m/%d/%Y')
    email = str(selected_sheet['C13'].value).strip()
    name = str(selected_sheet['C14'].value).strip()
    company = str(selected_sheet['C15'].value).strip()
    address = str(selected_sheet['C16'].value).strip()
    city = str(selected_sheet['C17'].value).strip()
    state = str(selected_sheet['C18'].value).strip()
    zipcode = str(selected_sheet['C19'].value).strip()

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = company
    order_data['CustomerAddress1'] = address
    order_data['CustomerAddress2'] = f'{city}, {state} {zipcode}'

    total = 0
    row_num = 24
    item_num = 0

    while selected_sheet[f'B{row_num}'].value is not None and item_num < 5:
        unit_number = str(selected_sheet[f'B{row_num}'].value).strip()
        quantity = selected_sheet[f'C{row_num}'].value
        price = product_prices[unit_number]
        amount = price * quantity

        order_data[f'Description{item_num}'] = product_descriptions[unit_number]
        order_data[f'Price{item_num}'] = f'${price:,.2f}'
        order_data[f'Quantity{item_num}'] = str(quantity)
        order_data[f'Amount{item_num}'] = f'${amount:,.2f}'

        total = total + amount
        row_num = row_num + 1
        item_num = item_num + 1

    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf', directory+f'Invoice {invoice_num}.pdf', order_data)

def send_email(invoice_num, name, email, date, directory):

    gmail_user = "jonathanoverbaugh@gmail.com"

    recipients = [email]

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""Dear {name},

Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

filenames = os.listdir(directory)
filenames.sort()

invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename and '~$' not in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        print(f'Invoice {invoice_num} sent to {name}')
        invoice_num = invoice_num + 1