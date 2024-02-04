import os
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SPREADSHEET_ID = "15sjS_1BAuVhQuNc4ntSuT18u2A_1FkC1E2xGnI2siXk"


# Main function

def main():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)  
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and  credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())

    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()

        for row in range(4, 28):

            skip = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!C{row}").execute().get("values")[0][0])

            p1 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!D{row}").execute().get("values")[0][0])

            p2 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!E{row}").execute().get("values")[0][0])

            p3 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!F{row}").execute().get("values")[0][0])

            m = (p1 + p2 + p3) / 30

            if skip > 15:

              sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!G{row}",
                                     valueInputOption="USER_ENTERED", body={"values": [[f"{"Reprovado por Falta"}"]]}).execute()
              
              sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!H{row}",
                                     valueInputOption="USER_ENTERED", body={"values": [[f"{0}"]]}).execute()

            else:

                if m < 5:  
      
                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!G{row}",
                                    valueInputOption="USER_ENTERED", body={"values": [[f"{"Reprovado por Nota"}"]]}).execute()
              
                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!H{row}",
                                    valueInputOption="USER_ENTERED", body={"values": [[f"{0}"]]}).execute()

                elif 5 <= m < 7:

                    naf = max(2 * 5 - m, 0)

                    naf = round(naf,1)  
                    
                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!G{row}",
                                    valueInputOption="USER_ENTERED", body={"values": [[f"{"Exame Final"}"]]}).execute()
              
                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!H{row}",
                                    valueInputOption="USER_ENTERED", body={"values": [[f"{naf}"]]}).execute()

                else:
       
                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!G{row}",
                                    valueInputOption="USER_ENTERED", body={"values": [[f"{"Aprovado"}"]]}).execute()
              
                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"sheetdesafio!H{row}",
                                    valueInputOption="USER_ENTERED", body={"values": [[f"{0}"]]}).execute()
                    
            
            # The porpuse of "time.sleep" is to delay the function, because of Google sheet's API, it would request more than it could.
            # "time.sleep" prevents errors, making the function work smoothly.
                    
            time.sleep(2.5)
    

    except HttpError as error:
        print(error)

if __name__ == "__main__":
    main()