# -*- coding: utf-8 -*-
from lxml import etree
from lxml import html
from openpyxl import Workbook
from datetime import datetime

class dou_parse:

    def geturls(self, xmlfile):
        companies = etree.parse(xmlfile)
        NSMAP = companies.getroot().nsmap.copy()
        NSMAP['xmlns'] = NSMAP.pop(None)
        return companies.xpath("//xmlns:loc/text()", namespaces=NSMAP)
        
    def grabinfo(self, urllist, city='Харьков'):
        cities = {'Харьков': 'harkov',
                  'Киев':'kiev'}
        result = {}
        l = len(urllist)
        for i, url in enumerate(urllist):
            print 'Progres: ', i + 1,' out of ', l,'| url: ',  url
            try:
                company = html.parse(url)
            except Exception:
                print 'Bad URL: ', url
                result.update({'Bad URL':{'staff': '---', 'offices': '---', 'adress': '---', 'site': '---', 'douurl': url}})
                continue
            offices = company.xpath("//div[@class='offices']/text()")
            if len(offices) > 0 and city.decode('utf-8') in offices[0]:
                offices = offices[0].strip()
            else:
                continue
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
            adress = companyoffice.xpath("//a[@name='"+cities[city]+"']/../div/div[2]/div[1]/div/div[1]/text()")
            if len(adress) > 0:
                adress = adress[0].strip()
            else:
                adress = ''
            result.update({companyname:{'staff': staff, 'offices': offices, 'adress': adress, 'site': site, 'douurl': url}})
        return result
        
    def writeinfo(self, data, filename='results.xlsx'):
        wb = Workbook()
        ws = wb.active
        ws.title = 'Results'
        ws.append(['Company', 'Staff', 'Offices', 'Adress', 'Site', 'Douurl'])
        for row in data:
            ws.append([row, data[row]['staff'], data[row]['offices'], data[row]['adress'], data[row]['site'], data[row]['douurl']])
        wb.save(filename)
        
        
if __name__ == '__main__':
    
    d = dou_parse()
    print 'Go!'
    starttime = datetime.now()
    urllist = d.geturls('sitemap-companies-112.xml')
    print 'URLs were taken successfully'
    res = d.grabinfo(urllist)
    print 'Info was grabbed successfully'
    d.writeinfo(res)
    print '\nDone!'
    print "Parser strted at: ", starttime.strftime("%Y-%m-%d %H:%M:%S")
    print "Finished at:      ", datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print "Elapsed time:     ", datetime.now() - starttime
