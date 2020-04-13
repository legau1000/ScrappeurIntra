#!/usr/bin/env python3 

import requests
from bs4 import BeautifulSoup
import time
import json
import xlrd
import xlsxwriter
from datetime import datetime
import smtplib
import unicodedata

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'}

blackList = []

class MyPresence:
    def __init__(self, name, present):
        self._name = name
        self._present = present

    def GetName(self):
        return (self._name)

    def GetPresent(self):
        return (self._present)

class MyScore:
    def __init__(self, name, xp):
        self._name = name
        self._xp = xp
        self._activity = []
        self._TypeActivity = [0, 0, 0, 0]

    def GetName(self):
        return (self._name)

    def AddXP(self, xp):
        self._xp = self._xp + xp

    def GetXP(self):
        return (self._xp)

    def GetActivity(self):
        return (self._activity)

    def AddActivity(self, activity):
        self._activity.append(activity)

    def GetTypeActivity(self):
        return (self._TypeActivity)

    def addViewer(self, tab):
        if (tab[0] == 1):
            self._TypeActivity[0] = self._TypeActivity[0] + 1
        if (tab[1] == 1):
            self._TypeActivity[2] = self._TypeActivity[2] + 1

    def addTeacher(self, tab):
        if (tab[0] == 1):
            self._TypeActivity[1] = self._TypeActivity[1] + 1
        if (tab[1] == 1):
            self._TypeActivity[3] = self._TypeActivity[3] + 1

    def verifLimite(self, data):
        self._TypeActivity[0] = self._TypeActivity[0] + data[0]
        self._TypeActivity[1] = self._TypeActivity[1] + data[1]
        self._TypeActivity[2] = self._TypeActivity[2] + data[2]
        self._TypeActivity[3] = self._TypeActivity[3] + data[3]
        if (data[0] == 1 and self._TypeActivity[0] > 15):
            return (False)
        if (data[1] == 1 and self._TypeActivity[1] > 6):
            return (False)
        if (data[2] == 1 and self._TypeActivity[2] > 10):
            return (False)
        if (data[3] == 1 and self._TypeActivity[3] > 3):
            return (False)
        return (True)



class IActivite:
    def __init__(self, link, tabInfos):
        self._link = link
        self._teacher = []
        self._student = []
        self.pres = tabInfos[0]
        self.abs = tabInfos[1]
        self.teacherpres = tabInfos[2]
        self.teacherabs = tabInfos[3]

    def SetTeacher(self, line):
        tabTeacher = line.find("div", {"class": "item teachers"})
        tmp = tabTeacher.findAll("a", {"class": "picture"})
        for teacher in tmp:
            idx = teacher['href'].find('user')
            idx = idx + 5
            self._teacher.append(teacher['href'][idx:])

    def scrapPresence(self, link):
        requete = requests.get(self._link + link, timeout=30, headers=headers)
        page = requete.content
        page = str(page)
        index = page.find("items:[")
        index = index + 9
        name = None
        present = None
        idxtwo = page[index:].find("login") + index
        while idxtwo - index != -1:
            index = idxtwo + 8
            idxtwo = page[index:].find("\"") + index
            name = page[index:idxtwo]
            idxtwo = page[index:].find("present") + index
            index = idxtwo + 10
            idxtwo = page[index:].find("\"") + index
            present = page[index:idxtwo]
            self._student.append(MyPresence(name, present))
            idxtwo = page[index:].find("login") + index

    def GetName(self):
        return (self._name)

    def SetScore(self):
        index = 0
        tabScore = []
        while index < len(self._student):
            pres = self._student[index].GetPresent()
            if (pres == "present"):
                xp = self.pres
            elif (pres == "absent"):
                xp = self.abs
            else:
                xp = 0
            if (xp != 0):
                Viewer = MyScore(self._student[index].GetName(), xp)
                Viewer.addViewer(self._TypeActivity)
                tabScore.append(Viewer)
            index = index + 1
        for line in self._teacher:
            teacher = MyScore(line, self.teacherpres)
            teacher.addTeacher(self._TypeActivity)
            tabScore.append(teacher)
        return (tabScore)

class Meetup(IActivite):
    def __init__(self, link, name):
        self._TypeActivity = [1, 0]
        self._name = name
        IActivite.__init__(self, link, [1, -1, 4, -4])

class Talk(IActivite):
    def __init__(self, link, name):
        self._TypeActivity = [1, 0]
        self._name = name
        IActivite.__init__(self, link, [1, -1, 4, -4])

class Workshop(IActivite):
    def __init__(self, link, name):
        self._TypeActivity = [0, 1]
        self._name = name
        IActivite.__init__(self, link, [3, -3, 10, -15])

class Hackathon(IActivite):
    def __init__(self, link, name):
        self._TypeActivity = [0, 0]
        self._name = name
        if name.find("Climathon") != -1:
            IActivite.__init__(self, link, [12, -6, 0, 0])
        else:
            IActivite.__init__(self, link, [6, -6, 15, -20])

class StockAll:
    def __init__(self, link):
        self._link = link
        self._people = []

    def listAll(self):
        for people in self._people:
            print(people.GetName(), people.GetXP())

    def writexls(self):
        # open the file for reading
        wbRD = xlrd.open_workbook('Hub.xlsx')
        sheets = wbRD.sheets()

        workbook = xlsxwriter.Workbook('Hub.xlsx')
        mtn = datetime.now()
        name = str(mtn.year) + "-" + str(mtn.month) + "-" + str(mtn.day)
        # run through the sheets and store sheets in workbook
        # this still doesn't write to the file yet
        for sheet in sheets: # write data from old file
            worksheet = workbook.add_worksheet(sheet.name)
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    worksheet.write(row, col, sheet.cell(row, col).value)
        newDay = workbook.add_worksheet(name)
        row = 0
        for people in self._people:
            newDay.write(row, 0, people.GetName())
            newDay.write(row, 1, people.GetXP())
            row = row + 1
        workbook.close()

    def AddPeople(self, tab, name):
        done = False
        for data in tab:
            for people in self._people:
                if (people.GetName() == data.GetName()):
                    if (people.verifLimite(data.GetTypeActivity())):
                        people.AddXP(data.GetXP())
                        people.AddActivity(name + " => " + str(data.GetXP()) + "xp")
                    else:
                        people.AddActivity(name + " => 0xp")
                    done = True
            if (done == False):
                acti = MyScore(data.GetName(), data.GetXP())
                acti.AddActivity(name + " => " + str(data.GetXP()) + "xp")
                self._people.append(acti)
            done = False

    def bubletri(self):
        haveChange = True
        index = 0
        tmp = []
        while (haveChange == True):
            if self._people[index].GetXP() < self._people[index + 1].GetXP():
                tmp = self._people[index]
                self._people[index] = self._people[index + 1]
                self._people[index + 1] = tmp
                index = 0
            else:
                index = index + 1
            if index > len(self._people) - 2:
                haveChange = False

    def TestWithoutMails(self):
        for people in self._people:
            if people.GetName() == "":
                print(str(people.GetXP()))
                print(str(people.GetActivity()))

    def SendMails(self):
        mtn = datetime.now()
        dayDate = str(mtn.day) + "/" + str(mtn.month) + "/" + str(mtn.year)
        s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
        s.starttls()
        s.login("<Your mail>", "<Your password>")
        fromaddr = '<Your name <Your mail>'
        sujet = "Recap XP Hub au " + dayDate
        blocked = False
        #sujet = "Erreur XP intra 12/11/2019"
        for people in self._people:
            blocked = False
            for blackListed in blackList:
                if (people.GetName() == blackListed):
                    blocked = True
            if blocked == False:
                toaddrs = [people.GetName()]
                xp = str(people.GetXP())
                message = """\
Bonjour,

La semaine du Hub est finie."

Je t'envoie donc un petit mail avec le récap des XP. Pour rappel:
Talk/Meetup: Participation: +1XP / -1XP
             Organisation: +4XP / -6XP

Workshop: Participation: +3XP / -3XP
          Organisation: +10XP / -15XP

Hackathon: Participation: +6XP / -6XP
Climathon: Participation: +12XP / -6XP

Pour les projets Hub, l'xp est à voir, mais on ne t'oublie pas!

Pour le moment tu as {}xp.

Le Hackathon Exoflow n'est pas pris en compte dans le calcul total. Pense à te le rajouter.

Hésite pas à venir nous voir pour nous proposer des Workshop / Hubtalk ou autre!
Voila un résumé de vos xp:\n
""".format(xp)
                message = message + self.lineActivity(people)
                message = message + "\nBien cordialement <3"

            #message = "Y'a une erreur, l'intra marche pas bien. Je relance le script de scrap quand l'intra ira mieux. Pas d'inquietude!"
                msg = """\
From: %s\r\n\
To: %s\r\n\
Subject: %s\r\n\
\r\n\
%s
""" % (fromaddr, ", ".join(toaddrs), sujet, message)
                s.sendmail(fromaddr, toaddrs, msg.encode('utf-8'))
        s.quit()



    def lineActivity(self, people):
        data = people.GetActivity()
        msg = ""
        for line in data:
            msg = msg + line + "\n"
        return (msg)