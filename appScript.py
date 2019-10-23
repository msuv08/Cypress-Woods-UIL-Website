'''
Created on Sep 15, 2019

@authors: mihir s. + alex h.
'''
import gspread, json, smtplib
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

with open("UIL-login.json", 'r') as f:
    datastore = json.load(f)
    
email = datastore["email"]
pas = datastore["pas"]
name=''
smtp = "smtp.gmail.com" 
port = 587
server = smtplib.SMTP(smtp,port)
server.starttls()
server.login(email,pas)
msg = MIMEMultipart()

listOfClasses = ''
eventTypeDict = {'Accounting':'rsmzzj', 'Calculator Applications':'15cs2j0', 'Computer Applications':'7namizm', 
                 'Computer Science':'04vbwvd', 'Current Issues and Events':'r2uaw6', 'Journalism':'c7v3xv7', 
                 'Literary Criticism':'l78iud', 'Math':'sjeun0', 'Number Sense':'lhkju6', 'Ready Writing':'tywm0p', 
                 'Science':'4rwknz3', 'Social Studies':'s6bad19','Speech and Debate':'24109v', 
                 'Spelling and Vocabulary':'i781v0'}

def parse_grade(n):
    if n == 'Freshman':
        return 9
    elif n == 'Sophomore':
        return 10
    elif n == 'Junior':
        return 11
    elif n == 'Senior':
        return 12
    else:
        return -1

def classRoom(c):
    retFinal = ''
    for event in c:
        for eventType in eventTypeDict:
            if event==eventType:
                retFinal += '\n' + event + ':  ' + eventTypeDict[event]+ '\n' 
    return retFinal

def sendMMS(number,getEvents):
    sms_gateway = number
    msg['From'] = email
    msg['To'] = sms_gateway
    msg['Subject'] = '\n'
    body = 'Welcome to the Cypress-Woods 2019-2020 UIL Club, ' + name + '!\n\nHere are your join codes:\n'
    body+=classRoom(getEvents)
    body+= '\n\nThanks for joining! 212!'
    msg.attach(MIMEText(body, 'plain'))
    sms = msg.as_string()
    server.sendmail(email,sms_gateway,sms)
    
    
def sendText(n,events):
    mmsGateway=open('numberlist.txt', 'r')
    for i in mmsGateway:
        phoneNumber = i.replace('number',n)
        sendMMS(phoneNumber,events)
    print('Text was sent.')
    
    
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('UIL Academics-e3225fa16671.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Contact Information (Responses)').sheet1

for i in range(sheet.row_count-1):
    if sheet.cell(i+1, 8).value == 'bot handled' or sheet.cell(i+1,8).value=='done':
        pass
    elif sheet.cell(i+1, 1).value == '':
        print('end of input')
        break
    else:
        name = sheet.cell(i+1, 3).value
        grade = parse_grade(sheet.cell(i+1, 5).value)
        email = sheet.cell(i+1, 2).value
        events = sheet.cell(i+1,6).value +''
        eventTypeList= events.split(', ')
        if '-' in str(sheet.cell(i+1,4).value):
            sheet.update_cell(i+1,4,str(sheet.cell(i+1,4).value).replace('-',''))
        phone = sheet.cell(i+1, 4).value  
        sendText(phone, eventTypeList)
        sheet.update_cell(i+1,8,'done')
        print(name + ' ' + str(grade) + 'th grade ' + email) 
server.quit()
        