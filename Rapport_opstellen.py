###############
### IMPORTS ###
###############

# Retrieve data from database
from sqlalchemy import create_engine        # pip install sqlalchemy
import pandas as pd                         # sql install pandas
import pymysql as pymysql                   # pip install pymysql
# Plot data with pdfkit
import pdfkit                               # pip install pdfkit
# Get date for document naming
from datetime import datetime 

#################
### VARIABLES ###
#################

username = 'root'
password = '1234'
host = '127.0.0.1'
database = 'unicentaopos'
db_connection_str = 'mysql+pymysql://{}:{}@{}/{}'.format(username, password, host, database)

"""
'wkhtmlopdf' needs to be installed for this to work
If not yet installed, go to folder 'wkhtmlopdf' 
and execute 'wkhtmltox-0.12.6-1.msvc2015-win64.exe' to install it
"""
path_wkhtmltopdf = '.\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'

#####################
### CONFIGURATION ###
#####################

db_connection = create_engine(db_connection_str)
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

#################
### FUNCTIONS ###
#################

def pandas_read_db():
    """
    Reads the data from a MySQL table
    Returns: data in a dataframe format
    """
    # Global variables
    global db_connection
    # SQL string
    sql_string = '\
        SELECT c.name as \'Categorie\', aantal as \'#\', p.name as \'Naam\', vervaldatum as \'Vervaldatum\' \
        FROM unicentaopos.products p \
        JOIN unicentaopos.expired e ON p.id = e.id\
        JOIN unicentaopos.categories c ON p.category = c.id\
        WHERE vervaldatum = curdate() OR vervaldatum = curdate() + INTERVAL 1 DAY\
        ORDER BY c.name, vervaldatum'
    # Read
    expired_dataframe = pd.read_sql(sql_string, con=db_connection)
    return expired_dataframe


def build_pdf_html(dataframe):
    """
    Uses dataframe formatted data to build pdf document
    Creates pdf documents in folder "Reports"
    Returns: nothing
    """
    # Global variables
    global config
    current_date = datetime.now()

    # Build document
    file = open('templates/Report.html', 'w')
    contents = dataframe.to_html()
    # Write document
    file.write(write_html_heading())
    file.write(contents)
    file.write(write_html_footer())
    file.close()
    # Name of document
    name = 'Rapport vervaldata - {}'.format(current_date.date())
    folder = "Reports"
    # Save document
    pdfkit.from_file('templates/Report.html', \
        "{}/{}.pdf".format(folder,name), \
        configuration=config)


def write_html_heading():
    """
    Builds a text variable, which can be used for writing html pages
    Creates an html header, imports css code in that header
    Returns: string with html header (including css)
    """
    # Global variables
    current_date = datetime.now()
    # html text
    text = '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Verslag vervaldata</title>
    '''
    text += write_html_ccs()
    text += '''
    </head>
    <body>
    <h1>Overzicht vervaldata</h1>
    <h2>{}</h2>
    '''.format(current_date.strftime("%d/%m/%Y - %H:%M"))
    return text


def write_html_footer():
    """
    Builds a text variable, which can be used for writing html pages
    Creates closing tags for html page
    Returns: string with closing tags for html page
    """
    return '''
    </body>
    </html>
    '''


def write_html_ccs():
    """
    Builds a text variable, which can be used for writing html pages
    Creates css code usable in the html header (internal)
    Returns: string with html header (including css)
    """
    text = '''
    <style>
        body {
            margin: 3rem;
            font-family: "Times New Roman";
        }
        h1, h2 {
            text-align: center;
            font-size: 150%;
        }

        h1 {
            margin-bottom: 0.25rem;
        }

        h2 {
            color: #696969;
            font-size: 125%;
            margin-top: 0.25rem;
        }

        table {
            width: 100%;
        }

        table, tr, th, td {
            border-collapse: collapse;
            border: none;
            padding: 1rem;
            text-align: left;
        }

        tbody > tr {
            border-top: 1px solid lightgrey;
        }

        tr > td:nth-child(5), thead > tr > th:nth-child(5) {
            text-align: right;
        }

        tr > td:nth-child(3), thead > tr > th:nth-child(3) {
            text-align: right;
            padding-right: 0;
            font-style: italic;
        }

        tr > th:first-child {
            display: none;
        }
    </style>
    '''
    return text


#################
### EXECUTION ###
#################

# db_info = pandas_read_db()
# build_pdf_html(db_info)

