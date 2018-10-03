# -*- coding: utf-8 -*-
import xlsxwriter
from collections import Counter
import collections
import os
from jira import JIRA
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime as d
from calculation import calculate
import time
from pathlib import Path


# Get execution time of the program
start_time = time.time()

# Statistics Calculator for Issue Objects
calculator = calculate()

# reload(sys)

# sys.setdefaultencoding('utf8')
now = d.datetime.now()
print(f'Rapor Alınmaya Başladı. . .({now})')
#Turkish Months
months = ['', u'Ocak', u'Şubat', u'Mart', u'Nisan', u'Mayıs', u'Haziran',
          u'Temmuz', u'Ağustos', u'Eylül', u'Ekim', u'Kasım', u'Aralık']

yy_ank_project_list = []
yy_ank_assg_list = []
yy_ank_priority_list = []
yy_ank_state_list = []
ig_ank_project_list = []
ig_ank_assg_list = []
ig_ank_priority_list = []
ig_ank_state_list = []
yy_ist_project_list = []
yy_ist_assg_list = []
yy_ist_priority_list = []
yy_ist_state_list = []
state_list = []
assignee_list = []
unassigned_list = []
satis_cat = []
total_count = 1
int_count = 1
user_name = 'JIRA_USER_NAME'
password = 'JIRA_PASSWORD'
mail_password = 'MAIL_PASSWORD'
# Jira Server Connection
options = {
    'server': 'YOUR_DOMAIN'}
# Auth
try:
    jira = JIRA(options, basic_auth=(f'{user_name}', f'{password}'))
except BaseException as Be:
    print(Be)
props = jira.application_properties()

########################### EXCEL FILES ###########################
# Output Name of The Excel File
outputxls = f"rapor_{d.datetime.now().strftime('%Y-%m-%d')}"
# Excel Header
head = str(months[now.month-1])+u' Ayı Raporu'
workbook = xlsxwriter.Workbook(os.path.dirname(
    os.path.realpath(__file__))+'\output\\'+outputxls+'.xlsx')
# ChartView
worksheet1 = workbook.add_worksheet(u'GenelDurum')
# Priorities Count Sheet
worksheet2 = workbook.add_worksheet(u'AnkaraYY')
# Projects Count Sheet
worksheet3 = workbook.add_worksheet(u'AnkaraIG')
# States Count Sheet
worksheet4 = workbook.add_worksheet(u'İstanbulYY')

########################### JIRA QUERIES ###########################
# All issues in a month. If written as startOfMonth(-1) issues of the last month will be queried.
all_issues = jira.search_issues(
    'created > startOfMonth(-1) AND created < endOfMonth(-1)', maxResults=False
)

# All closed issues of the last month
all_closed_issues = jira.search_issues(
    'resolution = Çözüldü or resolution = "Daha Sonra Çözülecek" or resolution = "İptal Edildi"  or resolution = Mükerrer  or resolution = "Yeniden Tekrarlanamadı" or resolution = "İptal Edildi"  and created > startOfMonth(-1) AND created < endOfMonth(-1) order by createdDate  asc', maxResults=False
)

# YY-Ankara Issues
yy_ankara_issues = jira.search_issues(
    'created > startOfMonth(-1) AND created < endOfMonth(-1) and category = YY-Ankara', maxResults=False
)
# IG-Ankara Issues
ig_ankara_issues = jira.search_issues(
    'created > startOfMonth(-1) AND created < endOfMonth(-1) and category = IG-Ankara', maxResults=False
)
# YY-İstanbul Issues
yy_istanbul_issues = jira.search_issues(
    'created > startOfMonth(-1) AND created < endOfMonth(-1) and category = YY-İstanbul', maxResults=False)

print("---Query Time:  %s seconds ---" % (time.time() - start_time))


def jiraQueryHandler(jiraQuery, priorityList, stateList, projectList, assgList):
    for i in range(0, len(jiraQuery)):
        priorityList.append(jiraQuery[i].raw[u'fields'][u'priority'][u'name'])
        stateList.append(jiraQuery[i].raw[u'fields'][u'status'][u'name'])
        projectList.append(jiraQuery[i].fields.project.name)
        try:
            assgList.append(jiraQuery[i].raw[u'fields'][u'assignee'][u'name'])
        except BaseException as Be:
            assgList.append('N/A')


# YY Ankara List
jiraQueryHandler(yy_ankara_issues, yy_ank_priority_list,
                 yy_ank_state_list, yy_ank_project_list, yy_ank_assg_list)
# IG-Ankara Lists
jiraQueryHandler(ig_ankara_issues, ig_ank_priority_list,
                 ig_ank_state_list, ig_ank_project_list, ig_ank_assg_list)
# YY-İstanbul Lists
jiraQueryHandler(yy_istanbul_issues, yy_ist_priority_list,
                 yy_ist_state_list, yy_ist_project_list, yy_ist_assg_list)


########################### SHEETS ###########################

# Sheet1 General View
# Widen the first column to make the text clearer.
worksheet1.set_column('A:A', 50)
# Add formats to highlight cells.
bold = workbook.add_format({'bold': True})
italic = workbook.add_format({'italic': True})
worksheet1.write(('A'+str(int_count)),
                 str(months[now.month-1])+u' Ayı Raporu', bold)
int_count += 1

worksheet1.write(('A'+str(int_count)), u'Toplam YY Ankara')
worksheet1.write(('B'+str(int_count)), len(yy_ankara_issues))
int_count += 1

worksheet1.write(('A'+str(int_count)), u'Toplam IG Ankara')
worksheet1.write(('B'+str(int_count)), len(ig_ankara_issues))
int_count += 1

worksheet1.write(('A'+str(int_count)), u'Toplam YY-İstanbul')
worksheet1.write(('B'+str(int_count)), len(yy_istanbul_issues))
int_count += 1

worksheet1.write(('A'+str(int_count)), u'Toplam Talep Sayısı', bold)
worksheet1.write(('B'+str(int_count)), len(all_issues))
int_count += 2

worksheet1.write(('A'+str(int_count)),
                 u'Ortalama Talep Kapanma Süresi (Gün) : ')
worksheet1.write(('B'+str(int_count)),
                 f'{calculator.meantime(all_closed_issues):1.2f}')
int_count += 2

worksheet1.write(('A'+str(int_count)),
                 u'Median Talep Kapanma Süresi (Gün) : ')
worksheet1.write(('B'+str(int_count)),
                 calculator.mediantime(all_closed_issues))

chart1 = workbook.add_chart({'type': 'pie'})

chart1.add_series({
    'name': 'Service Desk '+str(months[now.month-1]) + ' Ayı Talepleri',
    'categories': '=GenelDurum!$A$2:$A$4',
    'values':     '=GenelDurum!$B$2:$B$4',
    'points': [
        {'fill': {'color': '#5ABA10'}},
        {'fill': {'color': '#FE110E'}},
        {'fill': {'color': '#CA5C05'}},
    ],
})

try:
    worksheet1.insert_chart('D1', chart1)
except BaseException as Be:
    print(Be)

########################################## Sheet2 YY-Ankara Report ##########################################
# Priority Report
worksheet2.set_column('A:A', 50)
int_count = 1
worksheet2.write(('A'+str(int_count)),
                 u'Talep Önceliğine Göre Talep Sayıları', bold)
int_count += 1
yy_ank_priorty_count = Counter(yy_ank_priority_list).most_common()

chart5 = workbook.add_chart({'type': 'column'})
chart5.set_legend({'none': True})
chart5.add_series({
    'name': 'Talep Önceliğine Göre Talep Sayılar',
    'categories': '=AnkaraYY!$A$'+str(int_count)+':$A$'+str(int_count+len(yy_ank_priorty_count)-1),
    'values':     '=AnkaraYY!$B$'+str(int_count)+':$B$'+str(int_count+len(yy_ank_priorty_count)-1),
})
try:
    worksheet2.insert_chart('H1', chart5)
except BaseException as Be:
    print(Be)

for x in range(0, len(yy_ank_priorty_count)):
    worksheet2.write(('B'+str(int_count)), yy_ank_priorty_count[x][1])
    worksheet2.write(('A'+str(int_count)), yy_ank_priorty_count[x][0])
    int_count += 1

# State Report
int_count += 1
worksheet2.write(('A'+str(int_count)),
                 u'Talep Durumuna Göre Talep Sayıları', bold)
yy_ank_state_count = Counter(yy_ank_state_list).most_common()
int_count += 1

chart4 = workbook.add_chart({'type': 'column'})
chart4.set_legend({'none': True})
chart4.add_series({
    'name': 'Talep Durumuna Göre Talepler',
    'categories': '=AnkaraYY!$A$'+str(int_count)+':$A$'+str(int_count+len(yy_ank_state_count)-1),
    'values':     '=AnkaraYY!$B$'+str(int_count)+':$B$'+str(int_count+len(yy_ank_state_count)-1),
    'points': [
        {'fill': {'color': '#009900'}},
    ],
})
try:
    worksheet2.insert_chart('P1', chart4)
except BaseException as Be:
    print(Be)

for x in range(0, len(yy_ank_state_count)):
    if u'Resolved' == yy_ank_state_count[x][0]:
        yy_ank_state_count[x] = tuple([u'Çözüldü', yy_ank_state_count[x][1]])
    elif u'Closed' == yy_ank_state_count[x][0]:
        yy_ank_state_count[x] = tuple([u'Kapandı', yy_ank_state_count[x][1]])
    elif u'Waiting for support' == yy_ank_state_count[x][0]:
        yy_ank_state_count[x] = tuple(
            [u'Destek Bekleniyor', yy_ank_state_count[x][1]])
    elif u'Waiting for customer' == yy_ank_state_count[x][0]:
        yy_ank_state_count[x] = tuple(
            [u'Müşteriden Yanıt Bekleniyor', yy_ank_state_count[x][1]])
    elif u'On Hold' == yy_ank_state_count[x][0]:
        yy_ank_state_count[x] = tuple([u'Beklemede', yy_ank_state_count[x][1]])

for x in range(0, len(yy_ank_state_count)):
    worksheet2.write(('B'+str(int_count)), yy_ank_state_count[x][1])
    worksheet2.write(('A'+str(int_count)), yy_ank_state_count[x][0])
    int_count += 1

# Project Report
int_count += 1
worksheet2.write(('A'+str(int_count)), u'Projelere Göre Talep Sayıları', bold)
int_count += 1
chart2 = workbook.add_chart({'type': 'pie'})
chart2.add_series({
    'name': 'En Yoğun 5 Proje',
    # When all projects wanted to be listed instead of top 5 len(yy_ank_project_count) variable can be used.
    'categories': '=AnkaraYY!$A$'+str(int_count)+':$A$'+str(int_count+4),
    'values':     '=AnkaraYY!$B$'+str(int_count)+':$B$'+str(int_count+4),
    'points': [
        {'fill': {'color': '#3210ba'}},
        {'fill': {'color': '#ba106a'}},
        {'fill': {'color': '#10ba70'}},
        {'fill': {'color': '#d6ff5b'}},
        {'fill': {'color': '#f7891b'}}
    ],
})

try:
    worksheet2.insert_chart('H17', chart2)
except BaseException as Be:
    print(Be)
yy_ank_project_count = Counter(yy_ank_project_list).most_common()
for x in range(0, len(yy_ank_project_count)):
    worksheet2.write(('B'+str(int_count)), yy_ank_project_count[x][1])
    worksheet2.write(('A'+str(int_count)),
                     str(yy_ank_project_count[x][0]))
    int_count += 1

# Assignee Report
worksheet2.set_column('D:D', 30)
int_count = 1
worksheet2.write(('D'+str(int_count)),
                 u'Atanan Kişiye Göre Talep Sayıları', bold)
worksheet2.set_column('D:D', 15)
worksheet2.write(('F'+str(int_count)), u'Ortalama Geri Dönüş', bold)


int_count += 1

yy_ank_assignee_count = Counter(yy_ank_assg_list).most_common()

chart3 = workbook.add_chart({'type': 'column'})
chart3.set_legend({'none': True})
chart3.add_series({
    'name': 'Atanan Kişilere Göre Talep',
    'categories': '=AnkaraYY!$D$'+str(int_count)+':$D$'+str(len(yy_ank_assignee_count)),
    'values':     '=AnkaraYY!$E$'+str(int_count)+':$E$'+str(len(yy_ank_assignee_count)),
    'points': [
        {'fill': {'color': '#3210ba'}},
    ],
})
try:
    worksheet2.insert_chart('P17', chart3)
except BaseException as Be:
    print(Be)

for x in range(0, len(yy_ank_assignee_count)):
    worksheet2.write(('E'+str(int_count)), yy_ank_assignee_count[x][1])
    worksheet2.write(('D'+str(int_count)),
                     str(yy_ank_assignee_count[x][0]))
    int_count += 1

########################################## Sheet3 IG-Ankara Report ##########################################
# Priority Report
worksheet3.set_column('A:A', 50)
int_count = 1
worksheet3.write(('A'+str(int_count)),
                 u'Talep Önceliğine Göre Talep Sayıları', bold)
int_count += 1
ig_ank_priorty_count = Counter(ig_ank_priority_list).most_common()

chart9 = workbook.add_chart({'type': 'column'})
chart9.set_legend({'none': True})
chart9.add_series({
    'name': 'Talep Önceliğine Göre Talep Sayıları',
    'categories': '=AnkaraIG!$A$'+str(int_count)+':$A$'+str(int_count+len(ig_ank_priorty_count)-1),
    'values':     '=AnkaraIG!$B$'+str(int_count)+':$B$'+str(int_count+len(ig_ank_priorty_count)-1),
})
try:
    worksheet3.insert_chart('G1', chart9)
except BaseException as Be:
    print(Be)

for x in range(0, len(ig_ank_priorty_count)):
    worksheet3.write(('B'+str(int_count)), ig_ank_priorty_count[x][1])
    worksheet3.write(('A'+str(int_count)), ig_ank_priorty_count[x][0])
    int_count += 1

# State Report
int_count += 1
worksheet3.write(('A'+str(int_count)),
                 u'Talep Durumuna Göre Talep Sayıları', bold)
ig_ank_state_count = Counter(ig_ank_state_list).most_common()

int_count += 1

chart8 = workbook.add_chart({'type': 'column'})
chart8.set_legend({'none': True})
chart8.add_series({
    'name': 'Talep Durumuna Göre Talep',
    'categories': '=AnkaraIG!$A$'+str(int_count)+':$A$'+str(int_count+len(ig_ank_state_count)-1),
    'values':     '=AnkaraIG!$B$'+str(int_count)+':$B$'+str(int_count+len(ig_ank_state_count)-1),
    'points': [
        {'fill': {'color': '#009900'}},
    ],
})
try:
    worksheet3.insert_chart('O1', chart8)
except BaseException as Be:
    print(Be)

for x in range(0, len(ig_ank_state_count)):
    if u'Resolved' == ig_ank_state_count[x][0]:
        ig_ank_state_count[x] = tuple([u'Çözüldü', ig_ank_state_count[x][1]])
    elif u'Closed' == ig_ank_state_count[x][0]:
        ig_ank_state_count[x] = tuple([u'Kapandı', ig_ank_state_count[x][1]])
    elif u'Waiting for support' == ig_ank_state_count[x][0]:
        ig_ank_state_count[x] = tuple(
            [u'Destek Bekleniyor', ig_ank_state_count[x][1]])
    elif u'Waiting for customer' == ig_ank_state_count[x][0]:
        ig_ank_state_count[x] = tuple(
            [u'Müşteriden Yanıt Bekleniyor', ig_ank_state_count[x][1]])
    elif u'On Hold' == ig_ank_state_count[x][0]:
        ig_ank_state_count[x] = tuple([u'Beklemede', ig_ank_state_count[x][1]])


for x in range(0, len(ig_ank_state_count)):
    worksheet3.write(('B'+str(int_count)), ig_ank_state_count[x][1])
    worksheet3.write(('A'+str(int_count)), ig_ank_state_count[x][0])
    int_count += 1

# Project Report
int_count += 1
worksheet3.write(('A'+str(int_count)), u'Projelere Göre Talep Sayıları', bold)
int_count += 1
ig_ank_project_count = Counter(ig_ank_project_list).most_common()

chart6 = workbook.add_chart({'type': 'pie'})
chart6.add_series({
    'name': 'En Yoğun 5 Proje',
    'categories': '=AnkaraIG!$A$'+str(int_count)+':$A$'+str(int_count+5),
    'values':     '=AnkaraIG!$B$'+str(int_count)+':$B$'+str(int_count+5),
    'points': [
        {'fill': {'color': '#3210ba'}},
        {'fill': {'color': '#ba106a'}},
        {'fill': {'color': '#10ba70'}},
        {'fill': {'color': '#d6ff5b'}},
        {'fill': {'color': '#f7891b'}}
    ],
})

try:
    worksheet3.insert_chart('G17', chart6)
except BaseException as Be:
    print(Be)

for x in range(0, len(ig_ank_project_count)):
    worksheet3.write(('B'+str(int_count)), ig_ank_project_count[x][1])
    worksheet3.write(('A'+str(int_count)),
                     str(ig_ank_project_count[x][0]))
    int_count += 1

# Assignee Report
worksheet3.set_column('D:D', 30)
int_count = 1
worksheet3.write(('D'+str(int_count)),
                 u'Atanan Kişiye Göre Talep Sayıları', bold)
int_count += 1
ig_ank_assignee_count = Counter(ig_ank_assg_list).most_common()

chart7 = workbook.add_chart({'type': 'column'})
chart7.set_legend({'none': True})
chart7.add_series({
    'name': 'Atanan Kişilere Göre Talep',
    'categories': '=AnkaraIG!$D$'+str(int_count)+':$D$'+str(len(ig_ank_assignee_count)),
    'values':     '=AnkaraIG!$E$'+str(int_count)+':$E$'+str(len(ig_ank_assignee_count)),
    'points': [
        {'fill': {'color': '#3210ba'}},
    ],
})
try:
    worksheet3.insert_chart('O17', chart7)
except BaseException as Be:
    print(Be)


for x in range(0, len(ig_ank_assignee_count)):
    worksheet3.write(('E'+str(int_count)), ig_ank_assignee_count[x][1])
    worksheet3.write(('D'+str(int_count)),
                     str(ig_ank_assignee_count[x][0]))
    int_count += 1

########################################## Sheet3 YY-İstanbul Report ##########################################
# Priority Report
worksheet4.set_column('A:A', 30)
int_count = 1
worksheet4.write(('A'+str(int_count)),
                 u'Talep Önceliğine Göre Talep Sayıları', bold)
int_count += 1
yy_ist_priorty_count = Counter(yy_ist_priority_list).most_common()

chart10 = workbook.add_chart({'type': 'column'})
chart10.set_legend({'none': True})
chart10.add_series({
    'name': 'Talep Önceliğine Göre Talep Sayılar',
    'categories': '=İstanbulYY!$A$'+str(int_count)+':$A$'+str(int_count+len(yy_ist_priorty_count)-1),
    'values':     '=İstanbulYY!$B$'+str(int_count)+':$B$'+str(int_count+len(yy_ist_priorty_count)-1),
})
try:
    worksheet4.insert_chart('G1', chart10)
except BaseException as Be:
    print(Be)

for x in range(0, len(yy_ist_priorty_count)):
    worksheet4.write(('B'+str(int_count)), yy_ist_priorty_count[x][1])
    worksheet4.write(('A'+str(int_count)), yy_ist_priorty_count[x][0])
    int_count += 1

# State Report
int_count += 1
worksheet4.write(('A'+str(int_count)),
                 u'Talep Durumuna Göre Talep Sayıları', bold)
yy_ist_state_count = Counter(yy_ist_state_list).most_common()
int_count += 1

chart11 = workbook.add_chart({'type': 'column'})
chart11.set_legend({'none': True})
chart11.add_series({
    'name': 'Talep Durumuna Göre Talep',
    'categories': '=İstanbulYY!$A$'+str(int_count)+':$A$'+str(int_count+len(yy_ist_state_count)-1),
    'values':     '=İstanbulYY!$B$'+str(int_count)+':$B$'+str(int_count+len(yy_ist_state_count)-1),
    'points': [
        {'fill': {'color': '#009900'}},
    ],
})
try:
    worksheet4.insert_chart('O1', chart11)
except BaseException as Be:
    print(Be)


for x in range(0, len(yy_ist_state_count)):
    if u'Resolved' == yy_ist_state_count[x][0]:
        yy_ist_state_count[x] = tuple([u'Çözüldü', yy_ist_state_count[x][1]])
    elif u'Closed' == yy_ist_state_count[x][0]:
        yy_ist_state_count[x] = tuple([u'Kapandı', yy_ist_state_count[x][1]])
    elif u'Waiting for support' == yy_ist_state_count[x][0]:
        yy_ist_state_count[x] = tuple(
            [u'Destek Bekleniyor', yy_ist_state_count[x][1]])
    elif u'Waiting for customer' == yy_ist_state_count[x][0]:
        yy_ist_state_count[x] = tuple(
            [u'Müşteriden Yanıt Bekleniyor', yy_ist_state_count[x][1]])
    elif u'On Hold' == yy_ist_state_count[x][0]:
        yy_ist_state_count[x] = tuple([u'Beklemede', yy_ist_state_count[x][1]])

for x in range(0, len(yy_ist_state_count)):
    worksheet4.write(('B'+str(int_count)), yy_ist_state_count[x][1])
    worksheet4.write(('A'+str(int_count)), yy_ist_state_count[x][0])
    int_count += 1

# Project Report
int_count += 1
worksheet4.write(('A'+str(int_count)), u'Projelere Göre Talep Sayıları', bold)
int_count += 1
yy_ist_project_count = Counter(yy_ist_project_list).most_common()

chart12 = workbook.add_chart({'type': 'pie'})
chart12.add_series({
    'name': 'En Yoğun 5 Proje',
    'categories': '=İstanbulYY!$A$'+str(int_count)+':$A$'+str(int_count+5),
    'values':     '=İstanbulYY!$B$'+str(int_count)+':$B$'+str(int_count+5),
    'points': [
        {'fill': {'color': '#3210ba'}},
        {'fill': {'color': '#ba106a'}},
        {'fill': {'color': '#10ba70'}},
        {'fill': {'color': '#d6ff5b'}},
        {'fill': {'color': '#f7891b'}}
    ],
})

try:
    worksheet4.insert_chart('G17', chart12)
except BaseException as Be:
    print(Be)

for x in range(0, len(yy_ist_project_count)):
    worksheet4.write(('B'+str(int_count)), yy_ist_project_count[x][1])
    worksheet4.write(('A'+str(int_count)),
                     str(yy_ist_project_count[x][0]))
    int_count += 1

# Assignee Report
worksheet4.set_column('D:D', 30)
int_count = 1
worksheet4.write(('D'+str(int_count)),
                 u'Atanan Kişiye Göre Talep Sayıları', bold)
int_count += 1
yy_ist_assignee_count = Counter(yy_ist_assg_list).most_common()

chart13 = workbook.add_chart({'type': 'column'})
chart13.set_legend({'none': True})
chart13.add_series({
    'name': 'Atanan Kişilere Göre Talep',
    'categories': '=İstanbulYY!$D$'+str(int_count)+':$D$'+str(len(yy_ist_assignee_count)),
    'values':     '=İstanbulYY!$E$'+str(int_count)+':$E$'+str(len(yy_ist_assignee_count)),
    'points': [
        {'fill': {'color': '#3210ba'}},
    ],
})
try:
    worksheet4.insert_chart('O17', chart13)
except BaseException as Be:
    print(Be)

for x in range(0, len(yy_ist_assignee_count)):
    worksheet4.write(('E'+str(int_count)), yy_ist_assignee_count[x][1])
    worksheet4.write(('D'+str(int_count)),
                     str(yy_ist_assignee_count[x][0]))
    int_count += 1

workbook.close()
print(f'Excel File Saved. . . ({now})')


def sendemail(from_addr,
              to_addr_list,
              cc_addr_list,
              subject, body,
              login, password,
              smtpserver='smtp-mail.outlook.com:587'):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = str(months[now.month-1]) + \
        " Ayı Satış Kategorisine Göre Jira Raporu "
    msg['From'] = from_addr
    msg['To'] = ", ".join(to_addr_list)
    html = """\
     <html>
      <head></head>
      <body>
           <p>Merhabalar,<br/>
            """+str(months[now.month-1])+""" Ayı Jira raporu ektedir. <br/>
           İyi Çalışmalar <br/>
           </p>
     </body>
     </html>
     """

    # part1 = MIMEText(text, 'plain')
    part2 = MIMEText(html, 'html')
    # msg.attach(part1)
    msg.attach(part2)

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(os.path.dirname(os.path.realpath(
        __file__))+'\\output\\'+outputxls+".xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="rapor.xlsx"')
    msg.attach(part)
    server = smtplib.SMTP(smtpserver)
    server.starttls()
    server.login(login, password)
    server.sendmail(from_addr, to_addr_list, msg.as_string())
    server.quit()


# Mailing Created Excel File
try:
    sendemail(from_addr='huseyin.capan@netcad.com.tr',
              to_addr_list=['capanh@gmail.com'],
              # 'lokman.cetin@netcad.com.tr'],
              cc_addr_list='',
              subject=str(months[now.month])+" Ayı Jira Raporu",
              body="Last writing was successfull!",
              login="huseyin.capan@netcad.com.tr",
              password=f'{mail_password}')
except BaseException as Be:
    print(Be)

total_time_lapse = time.time() - start_time
cur_date = d.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

print(f"Mail Sent --- Total Elapsed Time:  {total_time_lapse} seconds ---")

my_file = Path(os.path.join(os.path.dirname(__file__)), 'output\\log.txt')

if my_file.exists():
    file = open(my_file, 'a')
    file.write('\n')
    file.write(f'{total_time_lapse} Seconds at {cur_date} ')
    file.close()
else:
    file = open(my_file, 'w')
    file.write('\n')
    file.write(f'{total_time_lapse} Seconds at {cur_date} ')
    file.close()
