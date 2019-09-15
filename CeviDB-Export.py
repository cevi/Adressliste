#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
import xlsxwriter
import datetime

# user token and get-params predefined
token = "INSERT TOKEN HERE"
email = "INSERT EMAIL HERE"
ende = "?user_email="+email+"&user_token="+token

ortsgruppe = "INSERT ORTSGRUPPE HERE"

# Gruppen und Farben einstellungen:
# groups: eine liste mit ID's (aus der CeviDB), welche in dieser Reihenfolge in der Adressliste auftauchen
groups = ["1758","1771","1756","2083","1760"]
# Colors: die hintergrundfarben für die jeweiligen gruppen (gleiche reihenfolge), wobei die erste farbe für die leiter, die zweite für die TN verwendet wird
colors = [["#f4b183","#f8cbad"],["#c184e0","#d6adea"],["#a9d18e","#c5e0b4"],["#8faadc","#b4c7e7"],["#ffd966","#ffe699"]]

# generate excel sheet
datum = datetime.datetime.today().strftime("%Y_%m_%d")
workbook = xlsxwriter.Workbook("Adressliste_Cevi_"+ortsgruppe+"_"+datum+".xlsx")
worksheet = workbook.add_worksheet()

# write titles
data = ["Funktion","Vorname","Nachname","Ceviname","Haupt-Email","Adresse","PLZ","Ort","Geschlecht","Geburtstag","Klasse/Beruf","Name Eltern","Stufe","Weitere Email","Eigene Handynummer","Telefonnummern"]

# format for title row
title_format = workbook.add_format()
title_format.set_bold()
title_format.set_border()
worksheet.write_row("A1",data,title_format)

#other inits
cols = "ABCDEFGHIJKLMNOP"
widths = [20,15,15,15,30,30,5,15,10,12,15,30,12,60,15,60]
count = 0

# set widths
for i in cols:
    worksheet.set_column(i+":"+i,widths[count])
    count+=1
total = 0

# cell format for yellow cells
yellow_cell = workbook.add_format()
yellow_cell.set_bg_color("yellow")


group_index = 0
for GRUPPE in groups:
    # read url for group, order by role asc
    res = requests.get("https://db.cevi.ch/groups/"+GRUPPE+"/people.json"+ende+"&sort=roles&sort_dir=asc")

    # print what it's doing
    for i in res.json()["linked"]["groups"]:
        if(i["id"]==GRUPPE):
            print("Doing group "+i["name"])
            break

    # get json
    gr = res.json()["people"]
    full_gr = res.json()
    for i in range(len(gr)):
        print("Fetching person "+gr[i]['first_name']+" "+gr[i]['last_name']+" "+str(i+1)+"/"+str(len(gr)))
        total +=1
        row = str(total+1)
        # get role and group by id
        for j in full_gr["linked"]["roles"]:
            if(j["id"]==gr[i]["links"]["roles"][0]):
                tn_role = j["role_type"]
                tn_group = j["links"]["group"]
                break
        phones = []
        email = []
        eigene_phone = None
        # get phone numbers, emails
        try:
            for j in full_gr["linked"]["phone_numbers"]:
                for k in gr[i]["links"]["phone_numbers"]:
                    if(k==j["id"] and j["label"]!="Mobil"):
                        phones.append(j["number"])
                    if(k==j["id"] and j["label"]=="Mobil"):
                        eigene_phone = j["number"]
            for j in full_gr["linked"]["additional_emails"]:
                for k in gr[i]["links"]["additional_emails"]:
                    if(k==j["id"]):
                        email.append(j["email"])
        except Exception:
            pass
        # get tn_group by name
        for j in full_gr["linked"]["groups"]:
            if(tn_group == j["id"]):
                tn_group = j["name"]
        # get some more data
        user_data = requests.get(gr[i]["href"]+ende).json()["people"][0]
        # get gender
        geschlecht = user_data["gender"].replace("m","männlich").replace("w","weiblich")
        # strip /-in if m, else replace only /-
        if(user_data["gender"] == "m"):
            if("Material" in tn_role): tn_role = "Materialverantwortlicher"
            else: tn_role = tn_role.replace("/-in","")
        elif(user_data["gender"] == "w"):
            if("Material" in tn_role): tn_role = "Materialverantwortliche"
            elif("Frei" in tn_role): tn_role = "Freie Mitarbeiterin"
            else: tn_role = tn_role.replace("/-","")
        else:
            pass

        try:
            birthday = datetime.datetime.strptime(user_data["birthday"],"%Y-%m-%d").strftime("%d.%m.%Y")
        except Exception:
            birthday = None
        # set active cell color
        col_ = workbook.add_format()
        col_.set_border()
        if("Teilnehmer" in tn_role):
            col_.set_bg_color(colors[group_index][1])
        else:
            col_.set_bg_color(colors[group_index][0])
        # set yellow cell format
        yellow = workbook.add_format()
        yellow.set_bg_color("yellow")
        yellow.set_border()
        col = col_
        # start writing to excel file
        worksheet.write("A"+row,tn_role,col)
        worksheet.write("B"+row,gr[i]["first_name"],col)
        worksheet.write("C"+row,gr[i]["last_name"],col)
        worksheet.write("D"+row,gr[i]["nickname"],col)
        worksheet.write("E"+row,gr[i]["email"],col)
        # fill yellow if empty, coninue writing to excel file
        if(gr[i]["address"] == None):
            col = yellow
        worksheet.write("F"+row,gr[i]["address"],col)
        col = col_
        if(gr[i]["zip_code"] == None):
            col = yellow
        worksheet.write("G"+row,gr[i]["zip_code"],col)
        col = col_
        if(gr[i]["town"] == None):
            col = yellow
        worksheet.write("H"+row,gr[i]["town"],col)
        col = col_
        if(geschlecht==None):
            col = yellow
        worksheet.write("I"+row,geschlecht,col)
        col = col_
        if(birthday==None):
            col = yellow
        worksheet.write("J"+row,birthday,col)
        col = col_
        if(user_data["profession"]=="" and "Teilnehmer" in tn_role):
            col = yellow
        worksheet.write("K"+row,user_data["profession"],col)
        col = col_
        # write even more data
        worksheet.write("L"+row,gr[i]["name_parents"],col)
        worksheet.write("M"+row,tn_group,col)
        worksheet.write("N"+row,str(email).strip('[]').strip("'").replace("', '",", "),col)
        worksheet.write("O"+row,eigene_phone,col)
        worksheet.write("P"+row,str(phones).strip('[]').strip("'").replace("', '",", "),col)
    group_index += 1
# write filter
worksheet.autofilter("A1:P"+str(total+1))
# close and write excel file
workbook.close()
