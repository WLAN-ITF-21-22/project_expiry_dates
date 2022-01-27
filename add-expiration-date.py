###############
### IMPORTS ###
###############

# Reading Excel
import openpyxl
from pathlib import Path
# MySQL
import mysql.connector  # if not recognized, use "pip install mysql-connector-python"

#################
### VARIABLES ###
#################

# Excel-documents
path_add_products = 'C:\\Users\\Lander Wuyts\\IT\\Sys Eng Projectweek\\Python'
name_add_products = 'Barcodes.xlsx'
#path_remove_products = path_add_products
#name_remove_products = ''

# MySQL
mysql_host="127.0.0.1"
mysql_user="unicenta"
mysql_password="abc123!"
mysql_database="unicentaopos"

######################
### INITIALIZATION ###
######################

# Excel
xlsx_file = Path(path_add_products, name_add_products)
wb_obj = openpyxl.load_workbook(xlsx_file)

sheet = wb_obj.active

# MySQL
mydb = mysql.connector.connect(
    host=mysql_host,
    user=mysql_user,
    password=mysql_password,
    database=mysql_database
)

db_cursor = mydb.cursor()

#################
### FUNCTIONS ###
#################

def read_excel(excelsheet):
    # Start at 2nd place, 1st spot reserved for headers
    index = 2
    # search for latest entry
    while excelsheet["A{}".format(index)].value != None:
        index += 1
    index -= 1
    # get values
    barcode = excelsheet["A{}".format(index)].value
    amount = excelsheet["B{}".format(index)].value
    expiration_date = excelsheet["C{}".format(index)].value

    # do something with values        
    print("Barcode: {}\nAmount: {}\nExpiration date: {}\n".format(barcode, amount, expiration_date))


def read_db(cursor):
    # Read data
    cursor.execute("SELECT * FROM products")
    # Fetch data
    result = cursor.fetchall()

    # do something with values
    for line in result:
        print(line)


#########################
### EXECUTION OF CODE ###
#########################

read_excel(sheet)

read_db(db_cursor)