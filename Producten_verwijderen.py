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
name_remove_products = 'Producten verwijderen.xlsx'

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


def read_excel(path, name):
    """
    Reads the Excel sheet set earlier.
    Returns all barcodes of sold products and the amount sold
    """
    # Excel
    xlsx_remove_file = Path(path, name)
    wb_remove_obj = openpyxl.load_workbook(xlsx_remove_file)

    sheet_remove = wb_remove_obj.active
    # Start at 2nd place, 1st spot reserved for headers
    index = 2
    products = []
    # search for latest entry
    while sheet_remove["A{}".format(index)].value != None:
        # get values
        barcode = sheet_remove["A{}".format(index)].value
        amount = sheet_remove["B{}".format(index)].value
        # add to list
        products.append([barcode, amount])
        # increment
        index += 1

    # return values
    return products


def find_id_db(cursor, barcode):
    """
    Given a barcode,
    this function returns the id if it exists in the database, 
    else it returns None and prints an error message
    """
    # Read data
    cursor.execute("SELECT id \
        FROM products \
        WHERE code = {}".format(barcode))
    result = cursor.fetchall()

    # Check if the barcode is already in the database
    if result == []:
        # Return None and error message if the barcode is not in the database
        message = "Barcode not in database. Add product to stock in uniCenta"
        print(message)
        return None
    else:
        # Return id
        return result[0][0]


def check_expiry_database(cursor, id):
    """
    Checks whether the product already exists within the expiration date database
    returns: True or False
    """
    # Check if the product date already appears in the database
    cursor.execute("SELECT * \
        FROM expired\
        WHERE id = '{}'".format(id))
    result = cursor.fetchall()
    # If the product is not yet in the expired database
    return result != []


def get_oldest_product(cursor, id):
    """
    Searches the "expired" database for the product with the oldest expiration date
    Returns: the expireID (Primary key for database "expired") of the product with the oldest expiration date
    """
    # if the product exists in the "expired" database, continue
    if check_expiry_database(cursor, id):
        cursor.execute("SELECT expireID\
            FROM expired\
            WHERE id = '{}'\
            ORDER BY vervaldatum".format(id)) 
        result = cursor.fetchall()
        return result[0][0]


def get_amount(cursor, expireID):
    """
    Returns: the amount of products with the given ID
    """
    cursor.execute("SELECT aantal\
        FROM expired\
        WHERE expireID = '{}'".format(expireID))
    result = cursor.fetchall()
    return result[0][0]


def remove_amount_db(db, cursor, id, amount):
    """
    If the product is in the "expired" database,
    the amount of products will be reducted by the amount specified in the Excel file
    If the amount of products becomes zero (or lower),
    the product is deleted from the "expired" database
    Returns: nothing
    """
    # Reduce the amount, one by one
    for _ in range(amount):
        # if the product exists in the "expired" database, reduce the amount value by 1
        if check_expiry_database(cursor, id):
            expireID = get_oldest_product(cursor, id)
            cursor.execute("UPDATE expired\
                SET aantal = (aantal - 1)\
                WHERE expireID = '{}'".format(expireID))
            # if the amount is zero, delete the entry
            if get_amount(cursor, expireID) <= 0:
                cursor.execute("DELETE FROM expired\
                    WHERE expireID = '{}'".format(expireID))
            db.commit()
        


def loop_remove_products(db, cursor, sold_products):
    """
    Loops over a list of (Excel) entries,
    executing "remove_amount_db" for every product
    """
    # loop over sold_products and remove amount for every product
    for product in sold_products:
        barcode = product[0]
        amount = product[1]
        db_id = find_id_db(cursor, barcode)
        remove_amount_db(db, cursor, db_id, amount)


#########################
### EXECUTION OF CODE ###
#########################

# sold_products = read_excel()
# loop_remove_products()
