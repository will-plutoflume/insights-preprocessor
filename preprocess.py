import pandas as pd
from os import path, listdir
from rich.progress import track
import json
import time

# Version 1.0

filepath = './input.xlsx'


def init():
    print("Loading input.xlsx file to memory")

    xls = pd.ExcelFile(filepath)

    if "Coversheet" in xls.sheet_names:
        tessian = load_report_defender(xls)
    elif "breaches" in xls.sheet_names:
        tessian = load_report_enforcer_historical(xls)
    else:
        print("Report is an unsupported type. Supported reports are: Defender Live, Enforcer Historical")
        quit()

    print("Writing to dough.json")
    f = open("dough.json", "w")
    f.write(json.dumps(tessian))
    f.close()
    print("All done 🎉")
    print(" ")
    print("dough.json is now ready to be uploaded to tessian.dev/bakery 🍩")

def load_report_defender(xls):
    print("Loaded ✅")
    print("Report is Defender 🟥")
    defender_flags = []

    sheets_to_load = []
    for sheetName in xls.sheet_names:
        if sheetName not in ["Coversheet", "Most Targeted", "Most Impersonated"]:
            if not "Custom Pro" in sheetName:
                sheets_to_load.append(sheetName)

    output_data = []
    for filter in sheets_to_load:
        print("Parsing sheet '" + filter + "'")
        current_row = {}
        df = xls.parse(filter, skiprows=8)
        print("Done ✅")

        is_first_row = True
        row_map = {}

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
                    else:
                        headers = parse_and_select_headers(str(row[1][col]))
                        row_obj["headers_parsed"] = headers

                output_data.append(row_obj)

    return output_data

def load_report_enforcer_historical(xls):
    print("Loaded ✅")
    print("Report is Enforcer Historical 🟨")

    sheets_to_load = ["breaches"]
               
    output_data = []
    for filter in sheets_to_load:
        print("Parsing sheet '" + filter + "'")
        df = xls.parse(filter, header=None)
        print("Done ✅")

        is_first_row = True
        row_map = {}

        df = df.fillna("")

        for row in track(df.iterrows(), description="Analysing each row...", total=df.shape[0]):

            if is_first_row:

                #First we read in the column header rows to use for our own mappings
                #This avoids hardcoding row positions and makes us a little more resilient to change
                index = 0
                for column in row[1]:
                    row_map[index] = str(column).lower().replace(" [internal]", "").replace(" ","_")
                    if (row_map[index] == "nan"):
                        row_map[index] = "check_log_id"
                    index += 1

                is_first_row = False

            else:

                col_blacklist = ["sensitivity_features","priority_dump", "regex"]
                row_obj = {}
                for col in range(len(row_map)):
                    if row_map[col] not in col_blacklist:
                        if (row_map[col] == "recipient_ads"):
                            row_obj[row_map[col]] = json.loads(str(row[1][col]).replace('"','').replace("'",'"'))
                        elif (row_map[col] == "attachment_names" or row_map[col] == "attachments_extensions"):
                            # This dirty hack is required because of the super weird 'json' that the xlsx gets
                            # Array elements are surrounded with ' UNLESS there's a ' in the element text, in which case they're surrounded by "
                            # This is bad

                            temprow = str(row[1][col])
                            temprow = temprow.replace('"',"'").replace(", '", ', "').replace("',", '",').replace("']",'"]').replace("['",'["')
                            row_obj[row_map[col]] = json.loads(temprow)
                        else:
                            row_obj[row_map[col]] = row[1][col]

                output_data.append(row_obj)

    return output_data

def parse_and_select_headers(blob):
    blacklist = [
        "received",
        "arc-authentication-results",
        "received-spf",
        "dkim-signature",
        "x-antiabuse",
        "x-get-message-sender-via",
        "x-authenticated-sender",
        'x-authenticated-sender',
        'x-source',
        'x-source-args',
        'x-source-dir',
        'x-helodomain',
        'x-mailcontrol-inbound',
        'x-mailcontrol-reportspam',
        'x-scanned-by',
        'return-path',
        'x-organizationheaderspreserved',
        'x-ms-exchange-organization-expirationstarttime',
        'x-ms-exchange-organization-expirationstarttimereason',
        'x-ms-exchange-organization-expirationinterval',
        'x-ms-exchange-organization-expirationintervalreason',
        'x-ms-exchange-organization-network-message-id',
        'x-eopattributedmessage',
        'x-ms-exchange-organization-messagedirectionality',
        'x-ms-exchange-skiplistedinternetsender',
        'x-crosspremisesheaderspromoted',
        'x-crosspremisesheadersfiltered',
        'x-ms-publictraffictype',
        'x-ms-traffictypediagnostic',
        'x-ms-exchange-organization-authsource',
        'x-ms-exchange-organization-authas',
        'x-originatororg',
        'x-ms-office365-filtering-correlation-id',
        'x-ms-exchange-crosstenant-originalarrivaltime',
        'x-ms-exchange-crosstenant-network-message-id',
        'x-ms-exchange-crosstenant-originalattributedtenantconnectingip',
        'x-ms-exchange-crosstenant-fromentityheader',
        'x-ms-exchange-transport-crosstenantheadersstamped',
        'x-ms-exchange-transport-endtoendlatency',
        'x-ms-exchange-processed-by-bccfoldering',
        'authentication-results',
        "x-microsoft-antispam-message-info",
        "mime-version",
        "x-ld-processed",
        "x-ms-exchange-crosstenant-id"
    ]

    lines = blob.split("\n")
    current_header = ""
    current_value = ""
    out = []

    for line in lines:
        words = line.split(" ")
        if (len(words[0]) > 0 and words[0][-1] == ":"):
            if (current_header.lower() not in blacklist):
                out.append([current_header,current_value])
            current_header = words.pop(0).replace(":","")
            current_value = " ".join(words)
        else:
            current_value += " " + line
    if (current_header.lower() not in blacklist):
        out.append([current_header,current_value])
    out.pop(0)
    return out

init()
