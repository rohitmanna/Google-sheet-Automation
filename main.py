# Loading the Environment
import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
# Opening a New Worksheet
 spreadsheet = client.open("Why")
 # To add a worksheet into Gsheets
ws = spreadsheet.add_worksheet(title="This is Me", rows=100, cols=20)
# Accessing a Specific Spreadsheet
gm = spreadsheet.get_worksheet(1)
# Updating Cells of a Worksheet
gm.update_cell(1,1,"Why Would You Hire Me?")
gm.format('A1', {'textFormat': {'bold': True}})
# Bulk Updation
cell_list = gm.range('A2:A11')
cell_values = ["Once Convinced a P.E Investor to invest in my employment","Built faith in One of the Finest Product Managers in the country through my ethics and skills",
               "I punch above my weight - Immense self understanding and dedication which allows me to do so. ", "I eat MySQL for Dinner","I Hate my comfort zone - I work to achieve what seems imposssible  ", "I am Lazy - So I upskill to automate monotonous tasks","I ask a ton of Questions - The path to Innovation Begins With Curiosity",
               "The 'H' in my name stands for Hustle", "I may miss out on the exact solution you are looking for, but i turn up with alternatives which are equally befitting","I want to Start something of my own - Thus Hiring me is also Sowing seeds to a future Angel-Founder relationship"]
# For Loop to Update all values at once
for i, val in enumerate(cell_values):  #gives us a tuple of an index and value
    cell_list[i].value = val    #use the index on cell_list and the val from cell_values
# Updates Values of A Cell
gm.update_cells(cell_list)





## Clearing the worksheet
gm.clear()
### Clearing specific cells of a worksheet
worksheet.batch_clear(["A1:B1", "C2:E2", "my_named_range"])


# Exporting pandas dataframe directly onto gsheet

import pandas as pd
import numpy

d1 = pd.read_csv(r'C:\Users\manna\Downloads\d2.csv')
# Load the library
from df2gspread import df2gspread as d2g

## Code to upload pandas dataframe - Sample below

#d2g.upload(dataframe,"spreadsheet_key","worksheet-NAME",credentials=creds, row_names=True)
d2g.upload(d1, "1pbMDwgGkHWzD7HPqqgaZnl2ibDFU2jGmEaTQfeUpH6s", "Sup", credentials=creds, row_names=True)
print(d1)
