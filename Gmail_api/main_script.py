from gmail_api import init_gmail_service, search_emails, get_email_message_details
from Google import create_service


client_secret_file = 'client_secret.json'

# Here I am setting up the API fro extraction of data from the mail 
gmail_service = init_gmail_service(client_secret_file)

# Here I am connecting python to the spreadsheet
sheet_service = create_service(
    client_secret_file,
    'sheets',
    'v4',
    ['https://www.googleapis.com/auth/spreadsheets']
)

# Here is the sheet name and the cell from which i am goin to show the data ........it can be changed to whatever space 
spreadsheet_id = '13cDKN4WOwH-EJ0b-9gWmbQ6RzbMvDoSJa9H3iFfB_Qo'  # Here you need to put the sheet id correctly, else the code would not run effectively
worksheet_name = 'gmail'
starting_cell = 'B2'
range_name =  f"'{worksheet_name}'!{starting_cell}"

#this query will give me data from the mail that have been sent by me 
query = 'from:me'
messages = search_emails(gmail_service, query, max_results=10)#this will allow to search the messages 


rows = [('Subject', 'Body', 'Date')]  #This is the header row or coloumn 

for msg in messages:
    details = get_email_message_details(gmail_service, msg['id'])
    if details:
        subject = details['subject']
        body = details['body'].replace('\n', ' ').replace('\r', '')[:200]  #here I have subjected the body to 200 chars only
        date = details['date']
        rows.append((subject, body, date))

# this snippet will write to the gmail spreadsheet
value_range_body = {
    'majorDimension': 'ROWS',
    'values': rows
}

response = sheet_service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id,
    valueInputOption='USER_ENTERED',
    range=range_name,
    body=value_range_body
).execute()

print(f"{response.get('updatedCells')} cells updated successfully!")
