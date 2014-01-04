'''
Created on Jan 4, 2014

@author: blaz1988
'''
from xlwt import *
import glob
import os

#download xlwt from https://pypi.python.org/pypi/xlwt




def excel_fun(data):
    wbook = Workbook()
    style1 = easyxf(
    'font: name Arial, bold yes, colour black, italic yes, height 180;'
    'alignment: vertical center, horizontal center, wrap yes;'
    'borders: left medium, right medium, top medium, bottom medium;'
    'pattern: pattern solid, fore_colour light_orange;')
    style2 = easyxf(
    'font: name Arial,  colour black,  height 180;'
    'alignment: vertical center, horizontal center, wrap yes;')
    sheetName="My Sheet"
    #Add name of sheeet
    wsheet = wbook.add_sheet(sheetName)
    

    wsheet.col(0).width = 10000
    wsheet.col(1).width = 5000
    
    
    wsheet.write(0, 0, "Website", style1)
    wsheet.write(0, 1, "Visits", style1)
    
    i=1
    for d in data:
        
        j=0
        wsheet.write(i, j, d["web"], style2)
        j+=1
        wsheet.write(i, j, d["visits"], style2)
        i+=1
        
        
    
        
 
            
            
    wbook.save("book.xls")
    print "Excel file book.xls is stored in " + os.getcwd()
websites=[{"web":"http://najponude.com/","visits":10000},{"web":"http://kako-napraviti.geek.hr/","visits":5000},{"web":"http://hackspc.com","visits":8000}]    
excel_fun(websites)
