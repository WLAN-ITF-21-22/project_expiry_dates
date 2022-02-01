##########################
### INSTALLING MODULES ###
##########################

import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])

# install('mysql-connector-python')

import mysql.connector  # if not recognized, use "pip install mysql-connector-python"

def read_db():
    mydb = mysql.connector.connect(
        host="localhost",
        user="root",
        password="1234",
        database="unicentaopos"
    )

    mycursor = mydb.cursor()

    mycursor.execute("SELECT * FROM products")

    myresult = mycursor.fetchall()

    text = ""
    for x in myresult:
        print(x)
    return text

print(read_db())