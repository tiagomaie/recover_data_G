from sys import path

import xlrd
import os
import shutil

my_dict = {}


# clean the output dir before we start
for outf in os.listdir("./output"):
    outf_path = os.path.join("./output", outf)
    try:
        if os.path.isfile(outf_path):
            os.unlink(outf_path)
    except Exception as e:
        print(e)

# TODO: iterate through dictioraries
# read dictionary
with open('./dictionaries/date_dictionary.txt', 'r') as f_dict:
    for line in f_dict:
        if line.strip():
            my_dict[str.split(line.strip(), ":")[0]] = str.split(line.strip(), ":")[1].split(",")

print(my_dict)
# print(my_dict['Echantillonnage_eulemurs_'][0])
# print(type(my_dict['Echantillonnage_eulemurs_']))
# print(my_dict.values())
# print(["Sec", "SymbolColor"] in my_dict.values())


for filename in os.listdir("./data/"):
    print(filename)
    tentative_key = set()
    wb = xlrd.open_workbook(filename="./data/"+filename)
    for ws_idx in range(0, wb.nsheets):  # for each worksheet in the workbook
        ws = wb.sheet_by_index(ws_idx)
        for row_idx in range(0,3):  # for the x to y rows in the worksheet
            ws_row = ws.row(row_idx)
            for cell in ws_row:  # for each cell in the rows
                #print(cell.value)
                for value in my_dict.values():  # for each list of values in my dictionary
                    if cell.value in value:  # if the value of a cell corresponds to a value on my dictionary
                        #print(cell.value)
                        tentative_key.add(cell.value)  # put this value in a set
    if list(tentative_key) in my_dict.values():  # if tentative_key is in the dictionary
        print(list(tentative_key))
        print(my_dict.keys()[my_dict.values().index(list(tentative_key))])
        # copy data file and rename it with desired tag
        shutil.copyfile(src="./data/"+filename,
                        dst="./output/test_" +
                            my_dict.keys()[my_dict.values().index(list(tentative_key))] +
                            ".xls")

