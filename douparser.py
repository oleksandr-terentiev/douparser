#!/usr/bin/python3
# -*- coding: utf-8 -*-

from lxml import etree
from lxml import html
from openpyxl import Workbook
from datetime import datetime
import csv
import sys
import argparse

# Configuration
version = "2.1.0"
citiescsv = 'cities.csv'
defaultcity = {'en':'harkov', 'ru': 'Харьков'}
defsitemap = 'sitemap-companies_test.xml'  # Change filename to your jobs.dou.ua/sitemap-companies.xml
defresfile = 'results.xlsx'

def main():
    # Preparing
    city = defaultcity
    try:
        citylist = list(csv.reader(open(citiescsv), delimiter=";"))
    except:
        citylist = [[city['en'], city['ru']]]
        print("Erorr (2) with: \""+citiescsv+"\". Used default city: "+city['en'])
    argparser = createParser([i[0] for i in citylist], defaultcity)
    params = argparser.parse_args(sys.argv[1:])
    if params.city is None:
        city = askcity(citylist, defaultcity)
    else:
        try:
            city['en'] = params.city
            city['ru'] = [i[1] for i in citylist if i[0] == city['en']][0]
        except:
            print("\nError! (3)")
            print("Somthing wrong with \""+citiescsv+"\"")
            print("Check please.")
            exit(3)
    dou = dou_parser(city)
    wb = Workbook()
    ws = wb.active
    ws.append(['Company', 'Staff', 'Offices', 'Adress', 'Site', 'Douurl'])
    
    # Getting the list of companies URLs
    try:
        urllist = list(set(dou.geturls(params.sitemap)))
    except:
        print("\nError! (4)")
        print("Somthing wrong with sitemap file: \""+params.sitemap+"\"")
        print("Check please.")
        exit(4)
        
    # Grabbing begins
    urllistlen = len(urllist)
    starttime = datetime.now()
    resfile = params.resfile
    print("Go!")
    print("Loking for offices in", city['ru'])
    for i, url in enumerate(urllist):
        result = dou.grabinfo(url)
        if result is not None:
            ws.append(result)
            wb.save(resfile)
        print("Progres: ", i + 1," out of ", urllistlen,"| url: ",  url)
    print("\nDone!")
    print("Parser strted at: ", starttime.strftime("%Y-%m-%d %H:%M:%S"))
    print("Finished at:      ", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    print("Elapsed time:     ", datetime.now() - starttime)
    print("\nResults was written into:", resfile)

def createParser (cities, defcity):
    parser = argparse.ArgumentParser(
            description = '''A Python script for grabbing information about companies which are situated in a particular city.
            You can run the scrypt without any parametres, choose a city manualy, and use default settings.''',
            epilog = '(c) Alex Terentiev 2016.'
            )
            
    parser.add_argument ('-c', '--city', choices=cities, #default=defcity['en'],
            help = 'City where we will search offices. Default \''+defcity['en']+'\'',
            metavar = 'CITY')
            
    parser.add_argument ('-sm', '--sitemap', default=defsitemap,
            help = 'Sitemap filename. Default \''+defsitemap+'\'',
            metavar = 'SITEMAP')
    
    parser.add_argument ('-r', '--resfile', default=defresfile,
            help = 'Filename for results. Default \''+defresfile+'\'',
            metavar = 'RESFILE')
            
    parser.add_argument ('-v', '--version',
            action='version',
            help = 'Print script version.',
            version='%(prog)s {}'.format (version))
    return parser
    
def askcity(cities, defcity):
    # User menu
    city = defcity
    print("Hi!")
    while True:
        print("Choose city, please. (\""+defcity['ru']+"\" is default, just press \"Enter\")")
        try:
            for i, rec in enumerate(cities):
                print(i+1, "-", rec[1])
        except:
            print("\nError! (5)")
            print("Somthing wrong with \""+citiescsv+"\"")
            print("Check please.")
            exit(5)
        print("Enter 0 for exit")
        citynumber = input("\nEnter the number of city: ")  # raw_input() in Python 2.7
        if len(citynumber) == 0:
            break
        try:
            citynumber = int(citynumber)
        except:
            print("Warning:", '"'+str(citynumber)+'"', "is not a correct number!!!\n")
            continue
        if citynumber == 0:
            exit(0)
        elif citynumber > 0 and citynumber <= len(cities):
            city['en'] = cities[citynumber-1][0]
            city['ru'] = cities[citynumber-1][1]
            break
        else:
            print("Warning:", '"'+str(citynumber)+'"', "is not a correct number of city!!!\n")
    return city
            
class dou_parser:
    
    def __init__(self, city=defaultcity):
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
            print('Bad URL: ', url)
            return ['Bad URL', '---', '---', '---', '---', url]
        offices = company.xpath("//div[@class='offices']/text()")
        if len(offices) > 0 and self.city['ru'] in offices[0]:  # self.city['ru'].decode('utf-8') in Python 2.7
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
    main()
