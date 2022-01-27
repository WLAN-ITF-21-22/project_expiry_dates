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

def print_list(list):
    """
    A function for printing items in a list. Useful for testing.
    Returns nothing.
    """
    for item in list:
        print(item)


def read_excel(excelsheet):
    """
    Reads a given excel sheet.
    Only the last entry (not accounting for gaps) is read.
    returns: a list containing the last entry's barcode, amount and expiration date
    """
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

    # return values
    return [barcode, amount, expiration_date]
    


def find_id_db(cursor, barcode):
    """
    Given a cursor (link to mysql database) and a barcode,
    this function returns the id if it exists in the database, 
    else it returns None and prints an error message
    """
    # Read data
    cursor.execute("SELECT id, code, name \
        FROM products \
        WHERE code = {}".format(barcode))
    # Fetch data
    result = cursor.fetchall()

    if result == []:
        # Return None and error message if the barcode is not in the database
        message = "Barcode not in database. Add product to stock in uniCenta"
        print(message)
        return None
    else:
        # Return id
        return result[0][0]


#########################
### EXECUTION OF CODE ###
#########################

scanned_product = read_excel(sheet)
barcode_db_id = find_id_db(db_cursor, '054102282742300')
print(barcode_db_id)
