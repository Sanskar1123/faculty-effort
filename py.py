import os.path
import googleapiclient
from googleapiclient.discovery import build
from google.oauth2 import service_account
import numpy as np

# SCOPES = ['https://www.googleapis.com/auth/sqlservice.admin']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'credentials.json'

creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# If modifying these scopes, delete the file token.json.


# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1wCZdDE9r9xxWkGfQfIkSJubNReHU-VJ8'
READ_RANGE_NAME = 'spring 2021!B2:N500'
WRITE_RANGE_NAME = 'spring 2021!Q2'
WRITE_TOTAL_LOAD_RANGE = 'Total Load Spring 2021!B1'

def calc_office_hours(no_of_students):
    if(no_of_students > 70):
        print("1\n")
        hours_per_student = 0.25
    elif(no_of_students > 40):
        print("2\n")
        hours_per_student = 0.5
    elif(no_of_students > 20):
        print("3\n")
        hours_per_student = 0.75
    else:
        print("4\n")
        hours_per_student = 1.0
    office_hours = hours_per_student * no_of_students/14
    print("Returning office hours for ", no_of_students, " as ", office_hours, "  ", hours_per_student, "\n\n")
    if(office_hours > 5):
        office_hours = 5
    elif(office_hours < 3/4):
        office_hours = 3/4    
    return office_hours

def calc_preparation(component):
    if(component == 'PRA' or component == 'TUT'):
        print("Went to if in calc_preparation")
        preparation = 2
    elif(component == 'LEC'):
        print("Went to elif in calc_preparation")
        preparation = 2
    else :
        print("Went to else in calc_preparation")
        preparation = -1
    return preparation

def calc_evaulation_time(no_of_students, component):
    return no_of_students*(1/4.0)*component/14.0

def calc_teaching_hours(component, credits):
    if(component == 'LEC'):
        print("Went to if in calc_teaching_hours")
        teaching_hours = credits
    elif(component == 'TUT'):
        print("Went to elif TUT in calc_teaching_hours")
        teaching_hours = credits
    elif(component == 'PRA'):
        print("Went to elif PRAC in calc_teaching_hours")
        teaching_hours = 2
    else:
        print("Went to else in calc_teaching_hours")
        teaching_hours = -1
    return teaching_hours

def calc_grading_component(component, credits):
    if(component == 'LEC' and credits >= 3):
        grading_component = 5
    elif(component == 'LEC' and credits <= 1.5):
        grading_component = 3
    else:
        grading_component = 0
    return grading_component

def calc_load(data):
    p = data.pop(0)
    faculty = {}
    prep = []
    load_on_faculty = []
    total_load = {}
    i=0
    # p = [['Teaching Hours', 'Preparation Hours:Teaching', 'Share Factor', 'Office Hours', 'Grading Components','Preparion Time:Grading','Evaluation Time', 'Faculty Load']]
    for row in data:                                                #For each entry in the excel
        print(i)
        share_factor = ((float)(row[9]))/100
        no_of_students = (float)(row[11])
        credits = (float)(row[10])
        teaching_hours = (float)(calc_teaching_hours(row[5], credits))
        office_hours = (float)(calc_office_hours(no_of_students))
        preparation = (float)(calc_preparation(row[5]))
        grading_component = (float)(calc_grading_component(row[5], credits))
        
        evaluation_time = (float)(calc_evaulation_time(no_of_students, grading_component))
        if row[0] in faculty:                                       #If the instructor entry is there in faculty dictionary
            if row[3] in faculty[row[0]]:                           #If the course entry respective to that instructor is in the dictionary
                for class_component in faculty[row[0]][row[3]]:     #For each class section taught by the faculty in that course
                    if(class_component == row[5]):                  #If the same faculty is already teaching the same component of the course,
                        preparation = 0                             #the preparation is already calculated once
                        break
                faculty[row[0]][row[3]][row[5]] = row[9]
                if(preparation==0):
                    grading_component = 0
                # load_on_faculty.append([(teaching_hours + teaching_hours*preparation)*share_factor + office_hours + (grading_component*3.0)/14.0])
            else:
                faculty[row[0]][row[3]] = {row[5]:row[9]}
                # load_on_faculty.append([(teaching_hours + teaching_hours*preparation)*share_factor + office_hours + (grading_component*3.0)/14.0])
                # p.append([teaching_hours], [preparation], [share_factor], [office_hours], [grading_component])
        else:
            total_load[row[0]] = 0
            faculty[row[0]] = {row[3]:{row[5]:row[9]}}
            # load_on_faculty.append([(teaching_hours + teaching_hours*preparation)*share_factor + office_hours + (grading_component*3.0)/14.0])
            # p.append([teaching_hours], [preparation], [share_factor], [office_hours], [grading_component])
        
        faculty_load = (teaching_hours + teaching_hours*preparation)*share_factor + office_hours + (grading_component*3.0*share_factor)/14.0 + evaluation_time
        total_load[row[0]] += faculty_load
        p.append([teaching_hours*share_factor, preparation*teaching_hours*share_factor, share_factor, office_hours, grading_component, (grading_component*3.0*share_factor)/14.0, evaluation_time*share_factor, faculty_load])
        i+=1    
    return p, total_load

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                            range=READ_RANGE_NAME).execute()
values = result.get('values', [])
if not values:
    print('No data found.')
# print(values[0][0],values[0][4])
load_list =[[1000,20], [2000], [12]]

prep, total_load = calc_load(values)
# prep = np.reshape(prep,(records,1))
# np.reshape(load,(-1,1))
print(prep,"\n\n\n\n",total_load)
request = sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                                range=WRITE_RANGE_NAME,valueInputOption='USER_ENTERED', body={"values":prep}).execute()
indiv = [['Faculty Name', 'Total Teaching Load']]
fac=[]
for faculty in total_load:
    indiv.append([faculty,total_load[faculty]])
sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                    range=WRITE_TOTAL_LOAD_RANGE,valueInputOption='USER_ENTERED', body={"values":indiv}).execute()
#sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
  #                          range='spring 2021!R3',valueInputOption='USER_ENTERED', body={"values":load}).execute()


    # print('Name, Major:')
    # for row in values:
    #Print columns A and E, which correspond to indices 0 and 4.
    #     print('%s, %s' % (row[0], row[4]))