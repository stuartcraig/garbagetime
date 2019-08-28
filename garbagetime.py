#---------------------------------------------------------------------------------------garbagetime.py
#
# Stuart Craig
# Last updated	20190827
#-----------------------------------------------------------------------------------------------------


import smtplib, ssl
import datetime
import xlrd

#----------------------------------------------------
# What's going on this week
#----------------------------------------------------

workbook = xlrd.open_workbook('D:/Dropbox/financial_admin/apartment/trash/20190827_trashschedule.xls')
sheet = workbook.sheet_by_index(0)

now = datetime.datetime.now()
xl_today = xlrd.xldate.xldate_from_date_tuple((now.year, now.month, now.day),0)

# Pull all columns
weeks = sheet.col(0)
weeks2= sheet.col(1)
holidays = sheet.col(2)
apt = sheet.col(3)

# Find today's stuff:
index=-9
max = len(weeks)
for i in range(0,max):
	dist = xlrd.xldate_as_datetime(sheet.cell_value(i,0),0) - now
	days = dist.days
	if days>-1 and days<7:
		index=i

#xlrd.xldate_as_datetime(trashday,0)
if index==-9:
	STOP

# Choose any week index for testing
#index=28

# Pull this week's values
apt=sheet.cell_value(index,3)
holiday=sheet.cell_value(index,2)
trashday = sheet.cell_value(index,1)

print("Weekly values:")
print(apt)
print(trashday)
print(holiday)


# Display date
displaydate = xlrd.xldate_as_datetime(trashday,0)
displaydate = (displaydate.strftime("%m/%d/%Y"))
print(displaydate)


#----------------------------------------------------
# Send the email:
#----------------------------------------------------

# Receiver depends on the week!
# Whose week is it?
if apt==1: 
	receiver_email = [emails of people in apartment 1]
	receivers = [names of people in apartment 1]
if apt==2: 
	receiver_email = [emails of people in apartment 2]
	receivers = [names of people in apartment 2]
if apt==3: 
	receiver_email = [emails of people in apartment 3]
	receivers = [names of people in apartment 3]


# Overwrite for testing
#receiver_email = "stuart.v.craig@gmail.com; stucraig@upenn.edu" 

# Email environment vars
port = 465
smtp_server = "smtp.gmail.com"
context = ssl.create_default_context()
sender_email = [account set up for this purpose]
password = [the password]

# Compose the message
if holiday:
	holiday_message = "Because of the holiday, trash pickup is on Thursday, " + displaydate
else:
	holiday_message = "Trash pickup is on Wednesday, " + displaydate

message1 = """\
Subject: Garbage reminder

"""
message2 = "Hi " + receivers + "!"
message3 = """ 
Just a quick reminder that this week is one of your trash duty weeks. Thanks!!
"""

message = message1 + message2 + message3 + holiday_message

# Send
with smtplib.SMTP_SSL(smtp_server, port) as server:
	server.login(sender_email, password)
	server.sendmail(sender_email, receiver_email, message)

