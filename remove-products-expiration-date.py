###############
### IMPORTS ###
###############

# Reading Excel
from tabnanny import check
import openpyxl             # pip install openpyxl
from pathlib import Path
# MySQL
import mysql.connector      # pip install mysql-connector-python

#################
### VARIABLES ###
#################

# Excel-documents
# path_remove_products = 'C:\\Users\\Lander Wuyts\\IT\\Sys Eng Projectweek\\Python'
path_remove_products = '.'
name_remove_products = 'Barcodes sold.xlsx'

# MySQL
mysql_host="127.0.0.1"
mysql_user="unicenta"
mysql_password="abc123!"
mysql_database="unicentaopos"

######################
### INITIALIZATION ###
######################

# Excel
xlsx_file = Path(path_remove_products, name_remove_products)
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
    Returns all barcodes of sold products and the amount sold
    """
    # Global variables
    global sheet
    # Start at 2nd place, 1st spot reserved for headers
    index = 2
    products = []
    # search for latest entry
    while sheet["A{}".format(index)].value != None:
        # get values
        barcode = sheet["A{}".format(index)].value
        amount = sheet["B{}".format(index)].value
        # add to list
        products.append([barcode, amount])
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


def check_expiry_database(id):
    """
    Checks whether the product already exists within the expiration date database
    returns: True or False
    """
    # Global variables
    global db_cursor
    # Check if the product date already appears in the database
    db_cursor.execute("SELECT * \
        FROM expired\
        WHERE id = '{}'".format(id))
    result = db_cursor.fetchall()
    # If the product is not yet in the expired database
    return result != []


def get_oldest_product(id):
    """
    Searches the "expired" database for the product with the oldest expiration date
    Returns: the expireID (Primary key for database "expired") of the product with the oldest expiration date
    """
    # Global variables
    global db_cursor
    # if the product exists in the "expired" database, continue
    if check_expiry_database(id):
        db_cursor.execute("SELECT expireID\
            FROM expired\
            WHERE id = '{}'\
            ORDER BY vervaldatum".format(id)) 
        result = db_cursor.fetchall()
        return result[0][0]


def get_amount(expireID):
    """
    Returns: the amount of products with the given ID
    """
    # Global variables
    global db_cursor
    db_cursor.execute("SELECT aantal\
        FROM expired\
        WHERE expireID = '{}'".format(expireID))
    result = db_cursor.fetchall()
    return result[0][0]


def remove_amount_db(id, amount):
    """
    If the product is in the "expired" database,
    the amount of products will be reducted by the amount specified in the Excel file
    If the amount of products becomes zero (or lower),
    the product is deleted from the "expired" database
    Returns: nothing
    """
    # Global variables
    global db_cursor
    # if the product exists in the "expired" database, reduce the amount value
    if check_expiry_database(id):
        expireID = get_oldest_product(id)
        db_cursor.execute("UPDATE expired\
            SET aantal = (aantal - {})\
            WHERE expireID = '{}'".format(amount, expireID))
        # if the amount is zero, delete the entry
        if get_amount(expireID) <= 0:
            db_cursor.execute("DELETE FROM expired\
                WHERE expireID = '{}'".format(expireID))
        mydb.commit()
        


def loop_remove_products():
    """
    Loops over a list of (Excel) entries,
    executing "remove_amount_db" for every product
    """
    # Global variables
    global sold_products
    # loop over sold_products and remove amount for every product
    for product in sold_products:
        barcode = product[0]
        amount = product[1]
        db_id = find_id_db(barcode)
        remove_amount_db(db_id, amount)


#########################
### EXECUTION OF CODE ###
#########################

sold_products = read_excel()
loop_remove_products()
