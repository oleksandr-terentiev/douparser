# -*- coding: utf-8 -*-
from lxml import etree
from lxml import html
from openpyxl import Workbook
from datetime import datetime
import csv

class dou_parse:
    
    def __init__(self, city = {'en': 'harkov', 'ru': 'Харьков'}):
        self.city = city
        
    def geturls(self, xmlfile):
        companies = etree.parse(xmlfile)
        NSMAP = companies.getroot().nsmap.copy()
        NSMAP['xmlns'] = NSMAP.pop(None)
        return companies.xpath("//xmlns:loc/text()", namespaces=NSMAP)
        
    def grabinfo(self, url):                
        try:
            company = html.parse(url)
        except Exception:
            print 'Bad URL: ', url
            return ['Bad URL', '---', '---', '---', '---', url]
        offices = company.xpath("//div[@class='offices']/text()")
        if len(offices) > 0 and self.city['ru'].decode('utf-8') in offices[0]:
            offices = offices[0].strip()
        else:
            return None
        companyname = company.xpath("//h1[@class='g-h2']/text()")[0].strip()
        staff = company.xpath("//div[@class='company-info']/text()")
        if len(staff) > 0:
            staff = max([el.strip() for el in staff])
        else:
            staff = ''
        site = company.xpath("//div[@class='site']/a/@href")
        if len(site) > 0:
            site = site[0]
        else:
            site = ''
        companyoffice = html.parse(url+'offices/')
        adress = companyoffice.xpath("//a[@name='"+self.city['en']+"']/../div/div[2]/div[1]/div/div[1]/text()")
        if len(adress) > 0:
            adress = adress[0].strip()
        else:
            adress = ''
        return [companyname, staff, offices, adress, site, url]
        
if __name__ == '__main__':
    
    city = {}
    cities = list(csv.reader(open('cities.csv'), delimiter=";"))
    print "Hi!"
    while True:
        print "Choose city, please. (\"Харьков\" is default, just press \"Enter\")"
        for i, rec in enumerate(cities):
            print i+1, "-", rec[1]
        print "Enter 0 for exit"        
        citynumber = raw_input ("\nEnter the number of city: ")
        if len(citynumber) == 0: 
            city['en'] = 'harkov'
            city['ru'] = 'Харьков'
            break
        try:
            citynumber = int(citynumber)
        except:
            print "Warning:", '"'+str(citynumber)+'"', "is not a correct number!!!\n"
            continue
        if citynumber == 0:
            exit()
        elif citynumber > 0 and citynumber <= len(cities):
            city['en'] = cities[citynumber-1][0]
            city['ru'] = cities[citynumber-1][1]
            break
        else:
            print "Warning:", '"'+str(citynumber)+'"', "is not a correct number of city!!!\n"
    
    u = dou_parse(city)
    wb = Workbook()
    ws = wb.active
    ws.append(['Company', 'Staff', 'Offices', 'Adress', 'Site', 'Douurl'])
    urllist = list(set(u.geturls('sitemap-companies-112.xml')))  # Change filename to yours jobs.dou.ua/sitemap-companies.xml
    urllistlen = len(urllist)
    starttime = datetime.now()
    print "Go!"
    print "Loking for offices in", city['ru']
    for i, url in enumerate(urllist):
        result = u.grabinfo(url)
        if result is not None:
            ws.append(result)
            wb.save('results.xlsx')
        print "Progres: ", i + 1," out of ", urllistlen,"| url: ",  url
    print "\nDone!"
    print "Parser strted at: ", starttime.strftime("%Y-%m-%d %H:%M:%S")
    print "Finished at:      ", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print "Elapsed time:     ", datetime.now() - starttime
