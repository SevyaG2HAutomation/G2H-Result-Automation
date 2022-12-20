import tkinter as tk
import datetime
from tkcalendar import DateEntry

w = tk.Tk()
w.title("G2h Result Automation")
w.geometry('330x200+450+100')
img = tk.PhotoImage(file="Sevya.png")
label = tk.Label(w, image=img)
label.place(x=10, y=10)

d1 = tk.StringVar()
d2 = tk.StringVar()
label = tk.Label(w, text=" ")
label.grid(row=0, column=0, padx=5, pady=10)
l1 = tk.Label(w, text=" ")
l1.grid(row=0, column=1, padx=5, pady=10)
label1 = tk.Label(w, text="Choose Start Date", bg='white', fg="black")
label1.grid(row=0, column=2, padx=5, pady=10)
cal = DateEntry(w, date_pattern='yyyy-mm-dd', textvariable=d1)
cal.grid(row=0, column=3, padx=15)

label2 = tk.Label(w, text='Choose End Dte', bg='white', fg='black')
label2.grid(row=2, column=2, padx=5, pady=10)
cal1 = DateEntry(w, date_pattern='yyyy-mm-dd', textvariable=d2)
cal1.grid(row=2, column=3, padx=15)


def filter():
    import webbrowser
    webbrowser.open_new("https://docs.google.com/spreadsheets/d/1H7fmq7ygBmIZMCHNLWWaDeRYhyMx_QCplcBCCPvljLA/edit#gid=0")
    import gspread
    import pandas as pd
    Cred_FILE = r'g2h-sevya-75101dd35267.json'
    gc = gspread.service_account(Cred_FILE)
    gc
    database = gc.open("G2H Result Automation")
    database
    wks = database.worksheet("Test1")
    wks
    import pandas as pd
    import pymysql
    con = pymysql.connect(host='127.0.0.1', user='root', password='root', database='g2hresultautomation')
    cur = con.cursor()
    qry = f'''select " ",u.createdDate,"","",firstName," "," "," ",dateOfBirth," "," "," "," "," "," "," "," ",mobile,email,user.experience," "," "," "," "," ",
        score,cutoffscore,j.name,c.name from user inner join userprogresssummary as u on user.id=u.userid  
        inner join jobs as j on j.id=u.jobId
        inner join userprogressdetails as up on u.id=up.userProgressSummaryId 
        inner join questchallengemapper as q on up.questChallengeMapperId=q.id 
        inner join challenge as c on c.id=q.challengeId where j.name='Sevya Mock Test 3'
        and date(user.createdDate) between '{cal.get()}' and '{cal1.get()}';
        '''
    #find the cell number
    cur.execute(qry)
    data = cur.fetchall()
    con.close()
    df = pd.DataFrame(data)

    from googleapiclient.discovery import build

    from google.oauth2 import service_account

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SERVICE_ACCOUNT_FILE = r'g2h-sevya-75101dd35267.json'

    creds = None
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    SAMPLE_SPREADSHEET_ID = '1H7fmq7ygBmIZMCHNLWWaDeRYhyMx_QCplcBCCPvljLA'

    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    word = "G2H TBD"
    cell = wks.find(word)
    c = "D" + str(cell.row + 1)

    word = "Not Qualified"
    cell = wks.find(word)
    e = "AF" + str(cell.row + 1)
    word = "G2H TBD"
    cell = wks.find(word)
    c1 = "A" + str(cell.row + 1)

    df[0] = "inprogress"
    # Average marks in all subjects
    data_case_1 = ['English Language Skills', 'Aptitude challenge', 'Logical reasoning challenge', 'Tecnical Apptitude']
    df_data_case_1 = df[28].isin(data_case_1)

    # cuttoff marks 7
    df_cutoff_case_1 = df[25] >= 7

    df_overall_case_1 = df[df_data_case_1 & df_cutoff_case_1]
    df_over_1 = df_overall_case_1[df_overall_case_1[28] == "Tecnical Apptitude"]

    # taking 'Tecnical Apptitude'
    data_case_3 = ['Tecnical Apptitude']
    df_data_case_3 = df[28].isin(data_case_3)

    # cuttoff marks 10
    df_cutoff_case_3 = df[25] >= 10

    df_overall_case_3 = df[df_data_case_3 & df_cutoff_case_3]

    # merge two dataframe
    raw_data = pd.concat([df_over_1, df_overall_case_3], ignore_index=True)

    # Remove duplicates
    final_data = raw_data.drop_duplicates()

    df1 = final_data
    df1[1]=df[1].dt.strftime('%Y-%m-%d')

    lst = df1.values.tolist()
    word = "Not Qualified"
    cell = wks.find(word)
    e1 = 'AF' + str(cell.row - 1)

    request = sheet.values().clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + (c) + ":" + (e1)).execute()

    word = "G2H TBD"
    cell = wks.find(word)
    c1 = "A" + str(cell.row + 1)

    word = "L1"
    cell = wks.find(word)
    d = "A" + str(cell.row + 1)

    request1 = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + d,
                                     valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS",
                                     body={"values": lst}).execute()
    word = "G2H TBD"
    cell = wks.find(word)
    c_end = "AF" + str(cell.row - 1)

    # To read Data
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + d + ":" + c_end).execute()
    values = result.get('values', [])
    data = pd.DataFrame(values)

    df = data[data[0] == "Selected"]
    df1 = data[data[0] != "Selected"]
    lst = df.values.tolist()
    lst1 = df1.values.tolist()

    while (True):
        while (True):
            x = request1["updates"]["updatedRange"]
            try:
                word = "L2"
                cell1 = wks.find(word)
                L2 = "A" + str(cell1.row + 1)

                word = "L1"
                cell = wks.find(word)
                L1_end = "AA" + str(cell.row - 1)
                result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=x).execute()
                values = result.get('values', [])
                data = pd.DataFrame(values)

                df = data[data[0] == "Selected"]
                df[0] = "L2 in process"
                lst = df.values.tolist()
                df1 = data[data[0] == "Notselected"]
                df1[0] = "Notselected for L2/failed in L1"
                lst1 = df1.values.tolist()
            except:
                pass
            l = lst + lst1
            if (l != []):
                for i in l:
                    try:
                        if (i[0] == "L2 in process"):
                            word = "Selected"
                            x3 = wks.find(word)
                            x4 = x3.row
                            wks.delete_rows(x4)

                            word = "L2"
                            cell = wks.find(word)
                            L2 = "A" + str(cell.row + 1)
                            request2 = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + L2,
                                                             valueInputOption="USER_ENTERED",
                                                             insertDataOption="INSERT_ROWS",
                                                             body={"values": lst}).execute()


                    except:
                        pass
                    else:
                        try:
                            word = "Notselected"
                            x3 = wks.find(word)
                            x4 = x3.row
                            wks.delete_rows(x4)

                            word = "Not Qualified"
                            cell = wks.find(word)
                            e2 = 'A' + str(cell.row + 1)
                            request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + e2,
                                                            valueInputOption="USER_ENTERED",
                                                            insertDataOption="INSERT_ROWS",
                                                            body={"values": lst1}).execute()
                            break
                        except:
                            pass
            else:
                break

        try:
            word = "L1"
            cell = wks.find(word)
            d = "A" + str(cell.row + 1)
            word = "G2H TBD"
            cell = wks.find(word)
            c_end = "AA" + str(cell.row - 1)

            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + d + ':' + c_end).execute()
            values = result.get('values', [])
            data = pd.DataFrame(values)
        except:
            pass
        if (len(data) == 0):
            break
    while (True):
        while (True):
            try:
                result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                            range="Test1!" + L2 + ":" + c_end).execute()
                values = result.get('values', [])
                data2 = pd.DataFrame(values)
                df1 = data2[data2[0] == "Selected in L2"]
                df1[0] = "L3 is in progress"
                df2 = data2[data2[0] == "Notselected in L2"]
                df2[0] = "Notselected for L3/failed in L2"
            except:
                pass

            lst2 = df1.values.tolist()
            lst3 = df2.values.tolist()
            l1 = lst2 + lst3

            if (l1 != []):
                for i in l1:
                    if (i[0] == "L3 is in progress"):
                        try:
                            word = "L3"
                            cell = wks.find(word)
                            L3 = "A" + str(cell.row + 1)
                            request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + L3,
                                                            valueInputOption="USER_ENTERED",
                                                            insertDataOption="INSERT_ROWS",
                                                            body={"values": lst2}).execute()
                            word = "Selected in L2"
                            x3 = wks.find(word)
                            x4 = x3.row
                            wks.delete_rows(x4)
                            count1 = count1 - 1
                            continue
                        except:
                            pass
                    else:
                        try:
                            word = "Notselected in L2"
                            x3 = wks.find(word)
                            x4 = x3.row
                            wks.delete_rows(x4)

                            word = "Not Qualified"
                            cell = wks.find(word)
                            e2 = 'A' + str(cell.row + 1)
                            request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Test1!" + e2,
                                                            valueInputOption="USER_ENTERED",
                                                            insertDataOption="INSERT_ROWS",
                                                            body={"values": lst3}).execute()
                            count1 = count1 - 1
                            break
                        except:
                            pass




            else:
                break
        try:
            word = "L2"
            cell = wks.find(word)
            L2 = "A" + str(cell.row + 1)
            word = "L1"
            cell = wks.find(word)
            L1_end = "AA" + str(cell.row - 1)
            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                        range="Test1!" + L2 + ":" + L1_end).execute()
            values = result.get('values', [])
            data2 = pd.DataFrame(values)
        except:
            pass
        if (len(data2) == 0):
            break


Button1 = tk.Button(w, text="Display", command=filter)
Button1.grid(row=3, column=3, padx=5)

w.mainloop()
