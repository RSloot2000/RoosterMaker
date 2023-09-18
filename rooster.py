from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from collections import Counter
from easygui import *
from tqdm import tqdm
import time
import datetime
import re
import xlsxwriter as excl
import itertools

options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-features=EnableEphemeralFlashPermission")
options.add_argument('--disable-logging') 
driver = webdriver.Chrome(options=options)

driver.get("https://persoonlijkrooster.ru.nl/schedule")
time.sleep(1)

select = driver.find_element(By.XPATH, "//a[@title='Inloggen']")
select.click()

select = driver.find_element(By.ID, 'idp__titleremaining1')
select.click()

F = True
while F == True:

    text = "Voer hier je gebruikersnaam en wachtwoord in"
    title = "Gegevens"
    fields = ["Gebruikersnaam", "Wachtwoord"]
    userww = multpasswordbox(text, title, fields)

    username = driver.find_element(By.NAME, "username")
    password = driver.find_element(By.NAME, "password")
    username.clear()
    password.clear()
    username.send_keys(userww[0])
    password.send_keys(userww[1])
    
    select = driver.find_element(By.ID, 'submit')
    select.click()
    time.sleep(0.5)
    
    try:
        driver.find_element(By.XPATH, "//div//p[contains(text(), 'De opgegeven')]")
        F = True
        msgbox("Verkeerde gebruikersnaam of wachtwoord. Probeer het opnieuw.", "Error")
    except:
        F = False
        
pbar = tqdm(total=100)

time.sleep(5)

now = datetime.datetime.now()
week = now.strftime("%U")
text = "Van welke week wil je het rooster hebben? (0 = huidige week, 1 = volgende week enz.)"
title = "Weekselectie"
output = enterbox(text, title)
week = int(week) + int(output)
if int(output) >= 1:
    select = driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[6]/div/div[3]/div/div[3]/div/div[2]/div/div[2]/div/div[2]/div[2]/div/button[3]')
    select.click()
    time.sleep(5)

pbar.update(10)
pbar.refresh()

days = []
monday = driver.find_element(By.XPATH, "//table/tbody/tr/td[contains(text(), 'ma')]").get_attribute("textContent")
days.append(monday)
tuesday = driver.find_element(By.XPATH, "//table/tbody/tr/td[contains(text(), 'di')]").get_attribute("textContent")
days.append(tuesday)
wednesday = driver.find_element(By.XPATH, "//table/tbody/tr/td[contains(text(), 'wo')]").get_attribute("textContent")
days.append(wednesday)
thursday = driver.find_element(By.XPATH, "//table/tbody/tr/td[contains(text(), 'do')]").get_attribute("textContent")
days.append(thursday)
friday = driver.find_element(By.XPATH, "//table/tbody/tr/td[contains(text(), 'vr')]").get_attribute("textContent")
days.append(friday)

course_codes = ['MGZ', 'PFS', 'PPG', 'CSI', 'CPR']
tests = ['VT', 'RADAR']
types = ['HC', 'LC', 'RC', 'WG', 'PR']

#2 = mo
#3 = tu
#4 = we
#5 = th
#6 = fr

pbar.update(10)
pbar.refresh()

def claslst(clas_m, clas_c, clas_t, g):
    i = 0
    ccc = 0
    for i in range(len(clas_m)):
        clas_list.append(clas_m[i].text)
        clas_list.append(clas_c[i].text)
        clas_list.append(clas_t[i].text)
        found = False
        for code in course_codes:
            if code in clas_m[i].text:
                found = True
                if g == 2:
                    clist[code].append(monday)
                elif g == 3:
                    clist[code].append(tuesday)
                elif g == 4:
                    clist[code].append(wednesday)
                elif g == 5:
                    clist[code].append(thursday)
                else:
                    clist[code].append(friday)
                clist[code].append(clas_c[i].text)
                clist[code].append(clas_t[i].text)
                break
        for code in tests:
            if code in clas_m[i].text:
                found = True
                if g == 2:
                    clist['Toetsen'].append(monday)
                elif g == 3:
                    clist['Toetsen'].append(tuesday)
                elif g == 4:
                    clist['Toetsen'].append(wednesday)
                elif g == 5:
                    clist['Toetsen'].append(thursday)
                else:
                    clist['Toetsen'].append(friday)
                clist['Toetsen'].append(clas_c[i].text)
                clist['Toetsen'].append(clas_t[i].text)
                break
            if ccc == 4:
                if found == False:
                    found = True
                    if g == 2:
                        clist['Overige'].append(monday)
                    elif g == 3:
                        clist['Overige'].append(tuesday)
                    elif g == 4:
                        clist['Overige'].append(wednesday)
                    elif g == 5:
                        clist['Overige'].append(thursday)
                    else:
                        clist['Overige'].append(friday)
                    clist['Overige'].append(clas_c[i].text)
                    clist['Overige'].append(clas_t[i].text)
            ccc = ccc + 1
    return clas_list, clist

pbar.update(10)
pbar.refresh()

total_dict = {day: [] for day in range(2,7)}
clas_list = []
crs = []
clist = {code: [] for code in course_codes}
clist.update({"Overige": []})
clist.update({"Toetsen": []})
clas_m_t = "//td[contains(@class, 'wc-day-column day-2')]//div[@class='wc-module-code']"
clas_c_t = "//td[contains(@class, 'wc-day-column day-2')]//div[@class='wc-module-name']"
clas_t_t = "//td[contains(@class, 'wc-day-column day-2')]//div[@class='wc-time']"

g = 2
m = 3
for r in range(5):
    clas_m = driver.find_elements(By.XPATH, clas_m_t)
    clas_c = driver.find_elements(By.XPATH, clas_c_t)
    clas_t = driver.find_elements(By.XPATH, clas_t_t)
    claslst(clas_m, clas_c, clas_t, g)
    for f in range(len(clas_list)):
        total_dict[g].append(clas_list[f])
    clas_list = []
    
    for t in range(len(clas_m)):
        for i in clas_m_t:
            if i.isdigit():
                clas_m_t = clas_m_t.replace(i, str(m))
    for t in range(len(clas_c)):
        for i in clas_c_t:
            if i.isdigit():
                clas_c_t = clas_c_t.replace(i, str(m))
    for t in range(len(clas_t)):
        for i in clas_t_t:
            if i.isdigit():
                clas_t_t = clas_t_t.replace(i, str(m))
    m = m + 1
    g = g + 1

pbar.update(10)
pbar.refresh()

MGZ_co = clist['MGZ']
PFS_co = clist['PFS']
PPG_co = clist['PPG']
CSI_co = clist['CSI']
CPR_co = clist['CPR']
Overige_co = clist['Overige']
Toetsen_co = clist['Toetsen']

workbook = excl.Workbook('Rooster Week ' + str(week) + '.xlsx')
bold = workbook.add_format({'bold': True})
bita = workbook.add_format({'italic': True, 'bold': True})
#WG Green
WGf = workbook.add_format()
WGf.set_bg_color('#CFFE96')
#HC Blue
HCf = workbook.add_format()
HCf.set_bg_color('#A3CDFF')
#RC Blue
RCf = workbook.add_format()
RCf.set_bg_color('#A3B7E2')
#PR Orange
PRf = workbook.add_format()
PRf.set_bg_color('#F9C081')
#Test Red
TSf = workbook.add_format()
TSf.set_bg_color('#E7A2A2')
#LC Blue
LCf = workbook.add_format()
LCf.set_bg_color('#A3B7E2')

worksheet0 = workbook.add_worksheet('Alle Lessen')
worksheet1 = workbook.add_worksheet('MGZ')
worksheet2 = workbook.add_worksheet('PFS')
worksheet3 = workbook.add_worksheet('PPG')
worksheet4 = workbook.add_worksheet('CSI')
worksheet5 = workbook.add_worksheet('CPR')
worksheet6 = workbook.add_worksheet('Overige')
worksheet7 = workbook.add_worksheet('Toetsen')

pbar.update(10)
pbar.refresh()

#Alle Lessen
row = 0
col = 0
headers = ['Module Code', 'Course Code', 'Tijd']
for i, day in enumerate(days):
    worksheet0.write(row, col, day, bita)
    for j, header in enumerate(headers):
        worksheet0.write(1, col + j, header, bold)
    col += len(headers)

col = 0
n = 3
for p, lst in enumerate(total_dict.values()):
    row = 2
    for i in range(0, len(lst), n):
        crs = lst[i:i+n]
        for l in range(len(crs)):
            worksheet0.write(row, col + l, crs[l])
        row += 1
    col += 3

pbar.update(10)
pbar.refresh()

headers = ['Dag', 'Course Code', 'Tijd', 'Type', 'ZSO?', 'ZSO af?']
n = 3  
color_dict = {'HC': HCf, 'RC': RCf, 'LC': LCf, 'WG': WGf, 'PR': PRf}

#MGZ
for col, header in enumerate(headers):
    worksheet1.write(0, col, header, bita)
for row, i in enumerate(range(0, len(MGZ_co), n), start=1):
    crs = MGZ_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet1.write(row, col, value)
        if col == 1:
            for ctype in types:
                if ctype in value:
                    format = color_dict.get(ctype)
                    worksheet1.write(row, col + 2, ctype, format)

#PFS
for col, header in enumerate(headers):
    worksheet2.write(0, col, header, bita)
for row, i in enumerate(range(0, len(PFS_co), n), start=1):
    crs = PFS_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet2.write(row, col, value)
        if col == 1:
            for ctype in types:
                if ctype in value:
                    format = color_dict.get(ctype)
                    worksheet2.write(row, col + 2, ctype, format)       

pbar.update(10)
pbar.refresh()

#PPG
for col, header in enumerate(headers):
    worksheet3.write(0, col, header, bita)
for row, i in enumerate(range(0, len(PPG_co), n), start=1):
    crs = PPG_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet3.write(row, col, value)
        if col == 1:
            for ctype in types:
                if ctype in value:
                    format = color_dict.get(ctype)
                    worksheet3.write(row, col + 2, ctype, format)

#CSI
for col, header in enumerate(headers):
    worksheet4.write(0, col, header, bita)
for row, i in enumerate(range(0, len(CSI_co), n), start=1):
    crs = CSI_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet4.write(row, col, value)
        if col == 1:
            for ctype in types:
                if ctype in value:
                    format = color_dict.get(ctype)
                    worksheet4.write(row, col + 2, ctype, format)

pbar.update(10)
pbar.refresh()

#CPR
for col, header in enumerate(headers):
    worksheet5.write(0, col, header, bita)
for row, i in enumerate(range(0, len(CPR_co), n), start=1):
    crs = CPR_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet5.write(row, col, value)
        if col == 1:
            for ctype in types:
                if ctype in value:
                    format = color_dict.get(ctype)
                    worksheet5.write(row, col + 2, ctype, format) 
pbar.update(10)
pbar.refresh()

#Overige
for col, header in enumerate(headers):
    worksheet6.write(0, col, header, bita)
for row, i in enumerate(range(0, len(Overige_co), n), start=1):
    crs = Overige_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet6.write(row, col, value)
        if col == 1:
            for ctype in types:
                if ctype in value:
                    format = color_dict.get(ctype)
                    worksheet6.write(row, col + 2, ctype, format)

headers = ['Dag', 'Toets', 'Tijd']
#Toetsen
for col, header in enumerate(headers):
    worksheet7.write(0, col, header, bita)
for row, i in enumerate(range(0, len(Toetsen_co), n), start=1):
    crs = Toetsen_co[i:i+n]
    for col, value in enumerate(crs):
        worksheet7.write(row, col, value, TSf)

worksheets = [worksheet0, worksheet1, worksheet2, worksheet3, worksheet4, worksheet5, worksheet6, worksheet7]
for worksheet in worksheets:
    worksheet.autofit()

pbar.update(10)
pbar.refresh()
workbook.close()
driver.quit()
