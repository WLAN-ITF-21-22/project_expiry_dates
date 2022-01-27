import openpyxl
from pathlib import Path

xlsx_file = Path('C:\\Users\\Lander Wuyts\\IT\\Sys Eng Projectweek\\Python', 'Barcodes.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file)

sheet = wb_obj.active

index = 1

while sheet["A{}".format(index)].value != None:
    index += 1
                
print(sheet["A{}".format(index - 1)].value)
