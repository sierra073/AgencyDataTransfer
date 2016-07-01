
from win32com.client import Dispatch
import os,pyodbc,csv,zipfile,time, io

#outlook connection and downloads the file from the specific folder
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6").Folders("Dashboards")
time.sleep(120)
all_inbox = inbox.Items
all_inbox.Sort("CreationTime")
msg = all_inbox.GetLast()


#loop searches for the sub and downloads the latest file
while msg:
	if 'Datorama | Report "Daily Spend by Placement"' in msg.Subject:
		break
	print 'hi'
	msg = all_inbox.GetPrevious()

#downloading all the attachments
for att in msg.Attachments:
	att.SaveAsFile("C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Costfile.zip")

loca="C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Costfile"
#delete all previous files in folder
for files in os.listdir(loca):
	os.remove(loca+'\\'+files)

#unzip file, put contents into folder
with zipfile.ZipFile("C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Costfile.zip",'r') as z:
    z.extractall("C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Costfile")

#read in file starting on the 7th line
x = os.listdir(loca)[0]
lines = open(loca+'\\'+x).readlines()
filex = lines[7:len(lines)]
file1 = [w.replace('"','').replace('\n','') for w in filex]

out_file=open("C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Costfile\\Clean_Cost.csv",'wb')
out_file.write('\n'.join(file1))
out_file.close()

os.system('fastload <C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\fastload.sql> C:\\Users\\scostanza\\Desktop\\testt.txt')

#connecting to teradata
conn=pyodbc.connect('DRIVER={Teradata};DBCNAME=;UID=;PWD=;QUIETMODE=YES;');
cursor=conn.cursor();

#Running the daily query
sqlQuery=''
sqlFile=io.open('C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Display dashboard schedule4.sql','r',encoding='utf-8')
for line in sqlFile:
    sqlQuery=sqlQuery + ' ' + line
    sqlQuery=(c for c in sqlQuery if 0 < ord(c) < 127)
    sqlQuery = ''.join(sqlQuery).replace('\n','')
sqlQuery_list = sqlQuery.split(';')
sqlFile.close()

sqlQuery_list=sqlQuery_list[:-1]

for i in range(0,len(sqlQuery_list)):
    print i
    if 'drop' in sqlQuery_list[i]:
        try:
            cursor.execute(sqlQuery_list[i]+';')
            conn.commit();
        except:
            print sqlQuery_list[i]
            pass
    else:
        cursor.execute(sqlQuery_list[i]+';')
        conn.commit();

# Closing connection
cursor.close()
conn.close();

from datetime import date, timedelta

conn=pyodbc.connect('DRIVER={Teradata};DBCNAME=;UID=;PWD=;QUIETMODE=YES;', ANSI=False, autocommit=True)
curs=conn.cursor()
sqlQuery=''
sqlFile=io.open('C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Daily Consumer Health Query.sql','r',encoding='utf-8-sig')
for line in sqlFile:
    sqlQuery=sqlQuery + ' ' + line
sqlFile.close()

curs.execute(sqlQuery)

header=''
for c in curs.description:
    header+=c[0] + ","


# Output
out_file="Y:\\Kepler\\Kepler_Display_DB_"+str(date.today() - timedelta(days=2)).replace('-','')+".csv"
with open(out_file, 'wb') as f:
    f.write(header.strip(',')+ '\n')
    csv.writer(f, quoting=csv.QUOTE_NONE).writerows(curs)

# Closing connection
curs.close()
conn.close();

#######################Merchant Export

conn=pyodbc.connect('DRIVER={Teradata};DBCNAME=;UID=;PWD=;QUIETMODE=YES;', ANSI=False, autocommit=True)
curs=conn.cursor()
sqlQuery=''
sqlFile=open('C:\\Users\\scostanza\\PycharmProjects\\Daily_Dashboard\\Daily Merchant Health Query.sql','r')
for line in sqlFile:
    sqlQuery=sqlQuery + ' ' + line
    sqlQuery=(c for c in sqlQuery if 0 < ord(c) < 127)
    sqlQuery = ''.join(sqlQuery).replace('\n','')
sqlFile.close()

curs.execute(sqlQuery)

header=''
for c in curs.description:
    header+=c[0] + ","


# Output
out_file="Y:\\Kepler\\Kepler_SMB_Display_DB_"+str(date.today() - timedelta(days=2)).replace('-','')+".csv"
with open(out_file, 'wb') as f:
    f.write(header.strip(',')+ '\n')
    csv.writer(f, quoting=csv.QUOTE_NONE).writerows(curs)

# Closing connection
curs.close()
conn.close();






