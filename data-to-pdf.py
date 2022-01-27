###############
### IMPORTS ###
###############

from sqlalchemy import create_engine  # pip install sqlalchemy
import pandas as pd  # sql install pandas
import pymysql as pymysql  # pip install pymysql
import matplotlib.pyplot as plt  # pip install matplotlib
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime 

#################
### VARIABLES ###
#################


username = 'root'
password = '1234'
host = '127.0.0.1'
database = 'unicentaopos'
db_connection_str = 'mysql+pymysql://{}:{}@{}/{}'.format(username, password, host, database)

##################
### INITIATION ###
##################

db_connection = create_engine(db_connection_str)

#################
### FUNCTIONS ###
#################

def pandas_read_db():
    # Global variables
    global db_connection
    # SQL string
    sql_string = 'SELECT c.name as \'Categorie\', p.name as \'Naam\', aantal as \'Hoeveelheid\', vervaldatum as \'Vervaldatum\' \
        FROM unicentaopos.products p \
        JOIN unicentaopos.expired e ON p.id = e.id\
        JOIN unicentaopos.categories c ON p.category = c.id'

    # Read
    expired_dataframe = pd.read_sql(sql_string, con=db_connection)
    return expired_dataframe

def build_pdf(dataframe):
    # Document layout
    fig, ax =plt.subplots(figsize=(12,8))
    ax.axis('tight')
    ax.axis('off')
    plt.title('Vervaldata')

    # Add data
    text_alignment = 'left'
    expiration_table = ax.table(cellText=dataframe.values,\
                colLabels=dataframe.columns,\
                loc='upper center',\
                cellLoc=text_alignment,\
                colLoc=text_alignment)
    # Name of document
    name = 'Rapport vervaldata - {}'.format(datetime.now().date())
    folder = "Reports"
    pp = PdfPages("{}/{}.pdf".format(folder, name))
    #pp = PdfPages("Reports/foo.pdf")
    pp.savefig(fig, bbox_inches='tight')
    pp.close()


#################
### EXECUTION ###
#################

build_pdf(pandas_read_db())