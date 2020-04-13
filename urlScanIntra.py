#!/usr/bin/env python

import re
import os
import csv
import requests
from bs4 import BeautifulSoup
from Class import Meetup, Workshop, Hackathon, Talk, StockAll

#lien auto login sans le #!/all de fin
autologin = "<AutoLogin>" + "/module/2019/B-INN-000/BDX-0-1/"

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'}

if __name__ == '__main__':
    allActivity = []
    requete = requests.get(autologin + "#!/all", timeout=30)
    page = requete.content
    soup = BeautifulSoup(page, "html.parser")
    stock = soup.find("ul", {"class": "past"})
    tmp = stock.findAll("li", {"data-nb_group": "1"})
    for activite in tmp:
        tmpActivity = None
        line = activite.find('div').find('h2').find('span').find('a')
        if line.text.find('Talk') != -1 or line.text.find('Google Developer Group') != -1:
            tmpActivity = Talk(autologin, line.text)
        elif line.text.find('Meetup') != -1 or line.text.find('sentation projet Urg') != -1 or line.text.find('Pycon') != -1:
            tmpActivity = Meetup(autologin, line.text)
        elif line.text.find('Workshop') != -1:
            tmpActivity = Workshop(autologin, line.text)
        elif line.text.find('Hackathon') != -1 or line.text.find('Semaine de l\'innovation') != -1:
            tmpActivity = Hackathon(autologin, line.text)
        if (tmpActivity != None):
            print(line)
            tmpActivity.SetTeacher(activite)
            link = activite.find("a", {"class": "registered"})
            try:
                tmpActivity.scrapPresence(link['href'])
                allActivity.append(tmpActivity)
            except:
                print("error")
        else:
            print("ERROR:" + line.text)
    TabAllPeople = StockAll(autologin)
    for scraped in allActivity:
        TabAllPeople.AddPeople(scraped.SetScore(), scraped.GetName())
    TabAllPeople.bubletri()
    #TabAllPeople.writexls()
    TabAllPeople.TestWithoutMails()
    #TabAllPeople.SendMails()