import pandas as pd
from os import path, listdir
from rich.progress import track
import json

# Version 1.0

filepath = './input.xlsx'

def init():
    print("Loading input.xlsx file to memory")
    tessian = load_report_defender(filepath)

    print("Writing to dough.json")
    f = open("dough.json", "w")
    f.write(json.dumps(tessian))
    f.close()
    print("All done 🎉")
    print(" ")
    print("dough.json is now ready to be uploaded to tessian.dev/bakery 🍩")


def load_report_defender(data_source_defender):
    xls = pd.ExcelFile(data_source_defender)
    print("Loaded ✅")
    defender_flags = []

    sheets_to_load = []
    for sheetName in xls.sheet_names:
        if sheetName not in ["Coversheet", "Most Targeted", "Most Impersonated"]:
            sheets_to_load.append(sheetName)

    for filter in sheets_to_load:
        print("Parsing sheet '" + filter + "'")
        current_row = {}
        df = xls.parse(filter, skiprows=8)
        print("Done ✅")

        is_first_row = True
        row_map = {}

        output_data = []

        for row in track(df.iterrows(), description="Analysing each row...", total=df.shape[0]):

            if is_first_row:

                #First we read in the column header rows to use for our own mappings
                #This avoids hardcoding row positions and makes us a little more resilient to change
                index = 0
                for column in row[1]:
                    row_map[index] = str(column).lower().replace(" [internal]", "").replace(" ","_")
                    index += 1

                is_first_row = False

            else:
                
                row_obj = {}
                for col in range(len(row_map)):
                    if (row_map[col] != "header"):
                        row_obj[row_map[col]] = str(row[1][col]) 

                output_data.append(row_obj)

        return output_data

init()
