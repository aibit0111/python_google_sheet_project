import gspread
from fpdf import FPDF
from google.oauth2 import service_account
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import requests
import PyPDF2
import os
import time

scope = ['https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('/Users/karlabachiega/Desktop/zz. Python Scripts/InstructionsPageGeneration/myproject1-383720-dfd9b521027f.json', scope)
service = build('sheets', 'v4', credentials=creds)

client = gspread.authorize(creds)

sheet_a = client.open('InstructionsPageGeneration').worksheet('SheetA')
sheet_b = client.open('InstructionsPageGeneration').worksheet('SheetB')

pdf = FPDF() 

desktop_folder = r'C:\Users\mohit\OneDrive\Desktop\remove_background\python_antonio\output' 

sheet_b_list = sheet_b.get_all_values()

for row in range(len(sheet_b_list)):
    if row == 0:
        continue
    folder_name = sheet_b_list[row][0]
    file_name = sheet_b_list[row][1]
    hyperlink = sheet_b_list[row][2]
    sheet_a.update_acell('B7', f'=HYPERLINK("{hyperlink}","CLICK HERE TO DOWNLOAD")')

    # create folder if it doesn't exist
    full_path = os.path.join(desktop_folder, folder_name)
    if not os.path.exists(full_path):
        os.makedirs(full_path)

    url = 'https://docs.google.com/spreadsheets/export?format=pdf&id=' + sheet_a.spreadsheet.id
    headers = {'Authorization': 'Bearer ' + creds.create_delegated("").get_access_token().access_token}
    res = requests.get(url, headers=headers)

    while res.status_code == 429:
        print('Rate limit exceeded. Sleeping for 60 seconds.')
        time.sleep(60)
        res = requests.get(url, headers=headers)

    if res.status_code == 200:
        # save file in the created folder
        temp_file_path = os.path.join(full_path, "temp_" + file_name + ".pdf")
        with open(temp_file_path, 'wb') as f:
            for chunk in res.iter_content(1024):
                f.write(chunk)
    else:
        print('Failed to download file. Status Code: ', res.status_code)
        print('Response: ', res.text)
        continue

    # read the downloaded pdf
    reader = PyPDF2.PdfReader(temp_file_path)

    # create a new pdf writer object
    writer = PyPDF2.PdfWriter()

    # add first page of the pdf to the writer object
    writer.add_page(reader.pages[0])

    # write the page to a new pdf
    first_page_file_path = os.path.join(full_path, file_name + ".pdf")
    with open(first_page_file_path, 'wb') as f:
        writer.write(f)

    # delete the temporary file
    os.remove(temp_file_path)
