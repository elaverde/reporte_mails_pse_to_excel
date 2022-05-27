# -*- coding: utf-8 -*-
from imap_tools import MailBox, AND
from bs4 import BeautifulSoup
import datetime
import re
from Excel import Excel
from dotenv import load_dotenv
import os
load_dotenv()
IMAP = os.getenv('IMAP')
EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
xls = Excel
xls.created_file(xls)
xls.write_cell(xls,"A1","Código")#2
xls.write_cell(xls,"B1","Estado")#1
xls.write_cell(xls,"C1","Empresa")#3
xls.write_cell(xls,"D1","Descripción")#4
xls.write_cell(xls,"E1","Servicios")#0
xls.write_cell(xls,"F1","Valor de la Transacción")#5
xls.write_cell(xls,"G1","Fecha")#6

def strip_tags(value):
    value = re.sub(r'<[^>]*?>', '', value)
    value=re.sub('\n+', '', value)
    return " ".join(value.split())

def get_mail_pse(email):
    bs = BeautifulSoup(msg.html, "html.parser")
    table = bs.table.table
    ignore=0
    for row in table.find_all('tr'):
        cells = row.find_all('td') 
        if(ignore==1): #permite selecionar la segunda tabla la primera no nos interesa
            if len(cells)>0:
                items = cells[0].prettify().split("<br/>")
                itemsData=[]
                for item in items:
                    item=strip_tags(item)
                    if(len(item)>0):
                        itemsData.append(item)
        ignore += 1
    ignore = 0
    #convertir a diccionario
    data=[]
    for item in itemsData:
        if(item.find("Gracias")>=0):
            data.append({
                "key":"Servicios",
                "value":item.split(":")[0].strip()
            })
        else:
            data.append({
                "key":item.split(":")[0].strip(),
                "value":item.split(":")[1].strip()
            })

    return data
  
with MailBox(IMAP).login(EMAIL, PASSWORD) as mailbox:
    emails=[]
    for msg in mailbox.fetch(AND(subject='PSE', date_gte=datetime.date(2020, 1, 15))):
        print(msg.subject)
        if msg.subject.find("Confirmación Transacción PSE")>=0:
            print(msg.date, msg.subject, len(msg.text or msg.html))
            emails.append({"email":get_mail_pse(msg.html)})

    #grabamos los email en el excel 
    countRow=2
    for email in emails:
        data = email["email"]
        xls.write_cell(xls,"A"+str(countRow),data[2]["value"])
        xls.write_cell(xls,"B"+str(countRow),data[1]["value"])
        xls.write_cell(xls,"C"+str(countRow),data[3]["value"])
        xls.write_cell(xls,"D"+str(countRow),data[4]["value"])
        xls.write_cell(xls,"E"+str(countRow),data[0]["value"])
        xls.write_cell(xls,"F"+str(countRow),data[5]["value"].replace("$ ","").replace(".",""))
        xls.write_cell(xls,"G"+str(countRow),data[6]["value"])
        countRow+=1
    
    xls.save_file(xls,"./reporte.xlsx")