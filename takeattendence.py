import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import time
print("start time:",datetime.datetime.now().time())


scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("googleattendencebotkey.json", scope)

client = gspread.authorize(creds)

main_sheet= client.open("BUEEC Attendance 2022").sheet1

#Input first half attendance Spreadsheet Name and worksheet name
first_sheet= client.open("Automatize Spring 2022 (Responses)").worksheet('First Half')

#Input second half attendance Spreadsheet Name and worksheet name
second_sheet=client.open("Automatize Spring 2022 (Responses)").worksheet('Second Half')

#write the cell Column Number where attendance will be taken; F=6
event_cell=4

main_data = main_sheet.get_all_values()
first_attendance=first_sheet.get_all_values()
second_attendance=second_sheet.get_all_values()

f1=0
f2=0
both=0
missing=[]

for serial,students in enumerate(first_attendance):
    time.sleep(1)
    if serial==0:
        pass
    else:
        f1+=1
        id=students[3]
        missing_id = True
        for index,info in enumerate(main_data):
            if info[1]==str(id):
                main_sheet.update_cell(index+1, event_cell, '1st Half')
                missing_id=False
        if missing_id:
            values_list = first_sheet.row_values(serial + 1)
            values_list.append("First Half")
            missing.append(values_list)

for serial,students in enumerate(second_attendance):
    time.sleep(1)
    if serial==0:
        pass
    else:
        f2+=1
        id=students[3]
        missing_id = True
        for index,info in enumerate(main_data):
            if info[1]==str(id):
                val = main_sheet.cell(index+1, event_cell).value
                if val=='1st Half':
                    both+=1
                    main_sheet.update_cell(index+1, event_cell, 'Present')
                else:
                    main_sheet.update_cell(index + 1, event_cell, '2nd Half')
                missing_id=False
        if missing_id:
            values_list = second_sheet.row_values(serial+1)
            values_list.append("Second Half")
            missing.append(values_list)

overall=f1+f2-both

s=str(f1)+" Present in First Half\n"
s+=str(f2)+" Present in Second Half\n"
s+=str(both)+" Present Full Time\n"
s+=str(overall)+" Present Overall\n"
s+=str(len(missing))+" ID were missing from the list\n"
s+="The missing id are:\n"
if missing==[]:
    s+="Everyone is successfully added\n"
else:
    for people in missing:
        s+=str(people)
        s+='\n'

with open(f'Attendance/{str(datetime.date.today())}.txt','w') as file:
    file.write(s)
    file.close()
print("done")
print("end time:",datetime.datetime.now().time())