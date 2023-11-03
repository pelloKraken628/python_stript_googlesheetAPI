#import all the packages
import imaplib, email, requests, quopri, re
import gspread
import time
import json
import smtplib
from google.oauth2 import service_account
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Scratch email configuration
user = 'user name or user email goes here'
password = 'user email app 2 step verification passed password goes here'
sender = 'sender email goes here'

#  IMAP server conf
imap_ssl_host = 'imap.gmail.com'
imap_ssl_port = 993
# SMTP server configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# Function to send an email with the sheet link URL
def send_email(to_email, subject, link):

    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # Convert the list of URLs to a string
    link_str = '\n'.join(link)

    # Attach the link string to the email
    msg.attach(MIMEText(link_str, 'plain'))
    
    # Connect SMTP server
    server = smtplib.SMTP(smtp_server, smtp_port) 
    server.starttls()
    server.login(user, password)

    # Send the email
    server.send_message(msg)
    server.quit()

def edit_gSheets(sheetList):

    # a file that is being opened in the code. 
    # It is likely a JSON file that contains some configuration or authentication information needed for accessing Google Sheets API. 
    # The code is using this file to load the service account information required for authentication with the Google Sheets API.
    with open('json file name permission allowed already goes here') as file:
        service_account_info = json.load(file)

    credentials = service_account.Credentials.from_service_account_info(service_account_info)
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds_with_scope = credentials.with_scopes(scope)

    client = gspread.authorize(creds_with_scope)

    for sheetLink in sheetList:
        spreadsheet = client.open_by_url(sheetLink)
        worksheet = spreadsheet.get_worksheet(0)
        column_value = 'Total Needed'
        cells = worksheet.findall(column_value)
        for cell in cells:
            column_index = cell.col
            last_row = len(worksheet.col_values(column_index))

            # Edit cells
            cell_range = f'{chr(64 + column_index)}{last_row + 1}'
            worksheet.update(range_name=cell_range, values=[['B-360246']])
            cell_format = {
                "horizontalAlignment": "LEFT"
            }
            worksheet.format(cell_range, cell_format)
            next_cell_range = f'{chr(64 + column_index + 1)}{last_row + 1}'
            worksheet.update(range_name=next_cell_range, values=[[6]])
            next_cell_format = {
                "horizontalAlignment": "RIGHT"
            }
            worksheet.format(next_cell_range, next_cell_format)
            
            cell_range = f'{chr(64 + column_index)}{last_row + 2}'
            worksheet.update(range_name=cell_range, values=[['B-614566']])
            cell_format = {
                "horizontalAlignment": "LEFT"
            }
            worksheet.format(cell_range, cell_format)
            next_cell_range = f'{chr(64 + column_index + 1)}{last_row + 2}'
            worksheet.update(range_name=next_cell_range, values=[[6]])
            next_cell_format = {
                "horizontalAlignment": "RIGHT"
            }
            worksheet.format(next_cell_range, next_cell_format)

last_email = 0

while True:
    # Have to login/logout each time because that's the only way to get fresh results.
    server = imaplib.IMAP4_SSL(imap_ssl_host, imap_ssl_port)
    server.login(user, password)
    server.select('INBOX')

    result, msgs = server.search(None, 'FROM', '"{}"'.format(sender))
    email_ids = msgs[0].split()
    if email_ids[-1].decode() == last_email:
        time.sleep(10)
        continue
    last_email = email_ids[-1].decode()
    url_list = []
    msg = server.fetch(email_ids[-1], '(RFC822)')
    for sent in msg:
        if type(sent[0]) is tuple:
            # encoding set as utf-8
            content = str(sent[0][1], 'utf-8')
            data = str(content)

            # Handling errors related to unicodenecode
            try:
                indexstart = data.find("ltr")
                data2 = data[indexstart + 5: len(data)]
                indexend = data2.find("</div>")
                decoded_data = quopri.decodestring(data2[0: indexend]).decode('utf-8')
                # Extracting URLs from decoded_data
                urls = re.findall(r'(https?://docs.google.com/spreadsheets\S+)', decoded_data)
                # Appending the URLs to the email_content_html variable
                for url in urls:
                    if "</a>" in url:
                        continue
                    temp_url = url.split('">https')[0]
                    url_list.append(temp_url)

            except UnicodeEncodeError as e:
                pass
    if len(url_list) > 0:
        edit_gSheets(url_list)
        send_email(sender, 'Yo hallo thire lol', url_list)
    server.logout()
    time.sleep(10)