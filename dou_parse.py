# -*- coding: utf-8 -*-
from lxml import etree
from lxml import html
from openpyxl import Workbook
from datetime import datetime


class dou_parse:
    
    city = ''
    citylist = {'harkov': 'Харьков',
                  'kiev': 'Киев'}
                  
    def __init__(self, city = 'harkov'):
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
        if len(offices) > 0 and self.citylist[self.city].decode('utf-8') in offices[0]:
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
        adress = companyoffice.xpath("//a[@name='"+self.city+"']/../div/div[2]/div[1]/div/div[1]/text()")
        if len(adress) > 0:
            adress = adress[0].strip()
        else:
            adress = ''
        return [companyname, staff, offices, adress, site, url]
        
if __name__ == '__main__':
    
    u = dou_parse('harkov')
    wb = Workbook()
    ws = wb.active
    ws.append(['Company', 'Staff', 'Offices', 'Adress', 'Site', 'Douurl'])
    urllist = u.geturls('sitemap-companies.xml')
    urllistlen = len(urllist)
    starttime = datetime.now()
    print 'Go!'
    for i, url in enumerate(urllist):
        result = u.grabinfo(url)
        if result is not None:
            ws.append(result)
            wb.save('results.xlsx')
        print 'Progres: ', i + 1,' out of ', urllistlen,'| url: ',  url
    print '\nDone!'
    print "Parser strted at: ", starttime.strftime("%Y-%m-%d %H:%M:%S")
    print "Finished at:      ", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print "Elapsed time:     ", datetime.now() - starttime
    
