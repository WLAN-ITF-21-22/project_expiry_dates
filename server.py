##########################
### INSTALLING MODULES ###
##########################

import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# install('flask')

###############
### IMPORTS ###
###############

# Producten toevoegen, verwijderen
import openpyxl             # pip install openpyxl
from pathlib import Path
import mysql.connector      # pip install mysql-connector-python

# Rapport opstellen
from sqlalchemy import create_engine        # pip install sqlalchemy
import pandas as pd   
import pdfkit                               # pip install pdfkit
from datetime import datetime 

# Email versturen
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders

# Python scripts
import Producten_toevoegen, Producten_verwijderen, Rapport_opstellen, Email_direct

# Flask
from flask import Flask, render_template, redirect

#################
### VARIABLES ###
#################

# Excel-documents
path_add_products = '.'
name_add_products = 'Producten toevoegen.xlsx'
path_remove_products = '.'
name_remove_products = 'Producten verwijderen.xlsx'

# MySQL
mysql_host="127.0.0.1"
mysql_user="root"  # "root" or "unicenta"
mysql_password="abc123!"
mysql_database="unicentaopos"

# Pdfkit MySQL
username = 'root'
password = 'abc123!'
host = '127.0.0.1'
database = 'unicentaopos'

path_wkhtmltopdf = '.\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'

# e-mail

message = """
Geachte,

In de bijlage kan u het document vinden met de overschot van deze week.

Met vriendelijke groeten
"""

sender = 'supermarkt.test@gmail.com'
password = '2ccs02AH3'
receiver = 'vzwsupermarkt.test@gmail.com'

reports_folder = 'Reports'

######################
### INITIALIZATION ###
######################

# Excel
xlsx_add_file = Path(path_add_products, name_add_products)
wb_add_obj = openpyxl.load_workbook(xlsx_add_file)

sheet_add = wb_add_obj.active

xlsx_remove_file = Path(path_remove_products, name_remove_products)
wb_remove_obj = openpyxl.load_workbook(xlsx_remove_file)

sheet_remove = wb_remove_obj.active

# MySQL
mydb = mysql.connector.connect(
    host=mysql_host,
    user=mysql_user,
    password=mysql_password,
    database=mysql_database
)

db_cursor = mydb.cursor()

# Pdfkit MySQL
db_connection_str = 'mysql+pymysql://{}:{}@{}/{}'.format(username, password, host, database)
db_connection = create_engine(db_connection_str)
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# e-mail
msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = receiver
msg['Subject'] = 'Briefing'

current_date = datetime.now()

name = 'Rapport vervaldata - {}'.format(current_date.date())
pdfname = './{}/{}.pdf'.format(reports_folder, name)

# Flask
app = Flask(__name__)

#########################
### EXECUTION OF CODE ###
#########################

@app.route('/')
def index():
  return render_template('index.html')

@app.route('/add-product/')
def add_product():
  print ('Adding products to expired database')

  scanned_products = Producten_toevoegen.read_excel()
  for scanned_product in scanned_products:
      barcode_db_id = Producten_toevoegen.find_id_db(scanned_product[0])
      Producten_toevoegen.write_entry_expiration_db(barcode_db_id, scanned_product)
  
  return redirect("http://127.0.0.1:5000/")

@app.route('/remove-product/')
def remove_product():
  print('Removing products from expired database')

  sold_products = Producten_verwijderen.read_excel()
  Producten_verwijderen.loop_remove_products(sold_products)

  return redirect("http://127.0.0.1:5000/")

@app.route('/create-report/')
def create_report():
  print('Creating a pdf report in the Reports folder')
  current_date = datetime.now()

  db_info = Rapport_opstellen.pandas_read_db()
  Rapport_opstellen.build_pdf_html(db_info)

  return redirect("http://127.0.0.1:5000/")

@app.route('/send-email/')
def send_email():
  print("Sending email")

  Email_direct.send_mail()

  return redirect("http://127.0.0.1:5000/")

@app.route('/open-report/')
def open_redirect():
  print("Toon rapport")

  return render_template('Report.html')

if __name__ == '__main__':
  app.run(debug=True)
