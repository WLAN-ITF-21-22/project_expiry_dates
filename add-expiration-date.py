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
path_remove_products = path_add_products
#name_remove_products = ''

######################
### INITIALIZATION ###
######################

xlsx_file = Path(path_add_products, name_add_products)
wb_obj = openpyxl.load_workbook(xlsx_file)

sheet = wb_obj.active

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


#########################
### EXECUTION OF CODE ###
#########################

read_excel(sheet)