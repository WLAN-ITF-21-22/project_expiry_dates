###############
### IMPORTS ###
###############

# Reading Excel
import openpyxl             # pip install openpyxl
from pathlib import Path
# MySQL
import mysql.connector      # pip install mysql-connector-python

#################
### VARIABLES ###
#################

# Excel-documents
#path_add_products = 'C:\\Users\\Lander Wuyts\\IT\\Sys Eng Projectweek\\Python'
path_add_products = '.'
name_add_products = 'Producten toevoegen.xlsx'

# MySQL
mysql_host="localhost"
mysql_user="root"  # or "root"
mysql_password="1234"
mysql_database="unicentaopos"

######################
### INITIALIZATION ###
######################

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
    Returns: a list containing every cell's barcode, amount and expiration date (except the first cell)
    """
    # Global variables
    global path_add_products
    global name_add_products
    # Excel
    xlsx_add_file = Path(path_add_products, name_add_products)
    wb_add_obj = openpyxl.load_workbook(xlsx_add_file)

    sheet_add = wb_add_obj.active
    # Start at 2nd place, 1st spot reserved for headers
    index = 2
    products = []
    # search for latest entry
    while sheet_add["A{}".format(index)].value != None:
        # get values
        barcode = sheet_add["A{}".format(index)].value
        amount = sheet_add["B{}".format(index)].value
        expiration_date = sheet_add["C{}".format(index)].value
        # add to lists
        products.append([barcode, amount, expiration_date.date()])
        # increment
        index += 1

    # return values
    return products
    

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
        WHERE id = '{}' AND vervaldatum = '{}'".format(id, expiration_date))
    result = db_cursor.fetchall()
    # # If the product is not yet in the expired database
    return result != []



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
            SET aantal = (aantal + {})\
            WHERE id = '{}' AND vervaldatum = '{}'".format(amount, id, expiration_date))
        # print("updated")
    else:
        # if the product is not yet in the database, add it
        db_cursor.execute("INSERT INTO expired(id, aantal, vervaldatum)\
            VALUES('{}', {}, '{}')".format(id, amount, expiration_date))
        # print("inserted")
    mydb.commit()


#########################
### EXECUTION OF CODE ###
#########################

# scanned_products = read_excel()
# for scanned_product in scanned_products:
#     barcode_db_id = find_id_db(scanned_product[0])
#     write_entry_expiration_db(barcode_db_id, scanned_product)
