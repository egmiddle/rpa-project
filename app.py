# Insert your python code here

#IMPORT LIBRARIES
import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows


#FUNCTIONS

# Grab Order Data from File for PDF
def get_order_data(invoice_num, filename, directory):
    update_dict = {}

    #Open order form
    order = openpyxl.load_workbook(directory + filename, data_only= True)
    sheet = order['OrderForm']

    #grab name, email, and date(normalized)
    cust_name = sheet['C14'].value
    cust_email = sheet['C13'].value
    order_date = sheet['C9'].value.strftime("%m/%d/%Y")

    #grab addresses & company
    cust_street_address = sheet['C16'].value
    cust_city = sheet['C17'].value
    cust_state = sheet['C18'].value
    cust_zip = sheet['C19'].value
    cust_company = sheet['C15'].value
    
    #Loop to get 3 product info and total amount
    total = 0 #starting total
    cell_num = 24 #starting cell number for unit number in order form
    fill_pdf_index = 0 #starting index for pdf fields
   

    while True:
        #Grab unit number
        unit_cell = 'B' + str(cell_num)
        unit_num = sheet[unit_cell].value

        #Check that there is a unit number for entry
        if unit_num is None:
            break

        #Get Quantity, Description, and Price
        quant_cell = 'C' + str(cell_num)
        quantity = sheet[quant_cell].value

        description = product_descriptions[unit_num]
        price = product_prices[unit_num]
        
        #Calculate Amount per Product
        amount = price * quantity 

        #add to update dictionary
        description_key = 'Description' + str(fill_pdf_index)
        price_key = 'Price' + str(fill_pdf_index)
        quant_key = 'Quantity' + str(fill_pdf_index)
        amount_key = 'Amount' + str(fill_pdf_index)

        update_dict[description_key] = str(description)
        update_dict[price_key] = f'${price:,.2f}'
        update_dict[quant_key] = str(quantity)
        update_dict[amount_key] = f'${amount:,.2f}'

        #Calculate Order Total
        total = total + amount

        #Iterate cell_num and fill_pdf_index to use next cell number and pdf field
        cell_num = cell_num + 1
        fill_pdf_index = fill_pdf_index + 1



    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = str(order_date)
    update_dict['Total'] = f'${total:,.2f}'
    update_dict['CustomerName'] = cust_name
    update_dict['CustomerCompany'] = cust_company
    update_dict['CustomerAddress1'] = cust_street_address
    update_dict['CustomerAddress2'] = f'{cust_city}, {cust_state} {cust_zip}'


    return update_dict, cust_name, cust_email, order_date



# Create and Fill Invoice PDF using Template
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)



# Send Email with Invoice attached
def send_email(invoice_num, name, email, date, directory):
    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.
    
    Thank you for your business.
        
    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = f'{directory}Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


#MAIN BODY

#Grab Email password
gmail_user = 'ethansgonzalez@gmail.com'
gmail_password_path = select_file('Please select the gmail password file')

gmail_password = open(gmail_password_path).read()


#Select Directory
while True:
    directory = single_input('Input directory containing the appropriate files')
    filenames = os.listdir(directory)

    #Make sure directory has appropriate ending
    if directory[-1] != '\\':
        directory = directory + '\\'

    #Validate directory has necessary files
    if 'ProductData.xlsx'in filenames and 'MasterInvoice.pdf'in filenames and 'Order1.xlsx'in filenames and 'Order2.xlsx'in filenames and 'Order3.xlsx'in filenames and 'Order4.xlsx' in filenames:
        break

    #Error message if wrong directory
    else:
        show_message('Error','Directory does not contain necessary files.\nPlease input a valid directory.')



#Create Description and Price Dictionaries from Product Data
product_descriptions = {}
product_prices = {}

df1 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name= 'Descriptions')
df1.head()

for index,row in df1.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df2 = pd.read_excel(directory + 'ProductData.xlsx', sheet_name= 'Prices')
df2.head()

for index, row in df2.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']



#Call Functions

filenames = os.listdir(directory)
invoice_num = 8515

for file in filenames:
    if 'Order' in file and 'xlsx' in file:
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
