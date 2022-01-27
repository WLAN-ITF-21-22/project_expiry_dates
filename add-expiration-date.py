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


def read_excel():
    """
    Reads the Excel sheet set earlier.
    Only the last entry (not accounting for gaps) is read.
    Returns: a list containing the last entry's barcode, amount and expiration date
    """
    # Global variables
    global sheet
    # Start at 2nd place, 1st spot reserved for headers
    index = 2
    # search for latest entry
    while sheet["A{}".format(index)].value != None:
        index += 1
    index -= 1
    # get values
    barcode = sheet["A{}".format(index)].value
    amount = sheet["B{}".format(index)].value
    expiration_date = sheet["C{}".format(index)].value

    # return values
    return [barcode, amount, expiration_date]
    

def find_id_db(barcode):
    """
    Given a barcode,
    this function returns the id if it exists in the database, 
    else it returns None and prints an error message
    """
    # Global variables
    global db_cursor
    # Read data
    db_cursor.execute("SELECT id \
        FROM products \
        WHERE code = {}".format(barcode))
    result = db_cursor.fetchall()

    # Check if the barcode is already in the database
    if result == []:
        # Return None and error message if the barcode is not in the database
        message = "Barcode not in database. Add product to stock in uniCenta"
        print(message)
        return None
    else:
        # Return id
        return result[0][0]


def check_expiry_database(id, expiration_date):
    """
    Checks whether the product already exists within the expiration date database
    returns: True or False
    """
    # Global variables
    global db_cursor
    # Check if the product with the expiration date already appears in the database
    db_cursor.execute("SELECT * \
        FROM expired\
        WHERE id = {} AND vervaldatum = {}".format(id, expiration_date))
    result = db_cursor.fetchall()
    # If the product is not yet in the expired database
    return result == []



def write_entry_expiration_db(id, scanned_product):
    """
    Updates "expires" database
    If the product with expiration date already appears, the product is updated
    If the product with expiration date does not appear, the product is inserted
    """
    # Global variables
    global db_cursor
    # Variables to be inserted
    amount = scanned_product[1]
    expiration_date = scanned_product[2]
    # Check if the product already exists in the database with this expiration date
    if check_expiry_database(id, expiration_date):
        # if the product is already in the database with the correct expiration date, update the amount
        db_cursor.execute("UPDATE expired \
            SET aantal += {}\
            WHERE id = {} AND vervaldatum = {}".format(amount, id, expiration_date))
    else:
        # if the product is not yet in the database, add it
        db_cursor.execute("INSERT INTO expired(id, aantal, vervaldatum)\
            VALUES({}, {}, {})".format(id, amount, expiration_date))


#########################
### EXECUTION OF CODE ###
#########################

scanned_product = read_excel()
# barcode_db_id = find_id_db(scanned_product)
barcode_db_id = find_id_db('05410228274230')
write_entry_expiration_db(barcode_db_id, scanned_product)