import mysql.connector  # if not recognized, use "pip install mysql-connector-python"

mydb = mysql.connector.connect(
    host="127.0.0.1",
    user="unicenta",
    password="abc123!",
    database="unicentaopos"
)

mycursor = mydb.cursor()

mycursor.execute("SELECT * FROM products")

myresult = mycursor.fetchall()

for x in myresult:
    print(x)