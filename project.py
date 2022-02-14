from xlrd import open_workbook
import urllib.request
import xlwt
import xlsxwriter
import requests
from bs4 import BeautifulSoup
from xlutils.copy import copy
import sqlite3

## sample.xls is the input file with url and the words to search for
rb = open_workbook('sample.xls')
book = xlsxwriter.Workbook('chart_pie1.xlsx')

##format specifications for the output file
headings = ['Words', 'Count','Frequency']
bold = book.add_format({'bold': 1})
format = book.add_format()
format.set_pattern(1)
format.set_bg_color('white')

##database for storage of word count data
conn=sqlite3.connect('mydb.db')
for s in rb.sheets():
    ##to drop existing tables with the same name
    dqr="drop table "+s.name+";"
    conn.execute(dqr)
    ##creation of tables
    qr="create table "+s.name+"(word text, count int);"
    conn.execute(qr)
    print ('Sheet:',s.name,s.nrows,s.ncols)
    url=s.cell(0,0).value
    r  = requests.get(url)
    soup = BeautifulSoup(r.text,'html.parser')
    type(soup)
    text=soup.get_text()
    s1 =  book.add_worksheet()
    print(s1.name)
    s1.write_row('A1',[url],bold)
    s1.write_row('A2', headings, bold)
    #print(text)
    sum1=0
    count1={}
    for row in range(s.nrows):
        if (row>1):
            print(s.cell(row,0).value)
            value  = (s.cell(row,0).value)
            print(text.count(value))
            count1[row+1]=text.count(value)
            sum1=sum1+count1[row+1]
            na='A'+str(row+1)
            nb='B'+str(row+1)
            v=[count1[row+1]]
##writing data into excel file
            s1.write_row(na,[value])
            s1.write_row(nb,v)
##inserin data into the database
            qr1="insert into "+s.name+"(word, count) values (\""+str(value)+"\","+str(count1[row+1])+")"
            #print(qr1)
            conn.execute(qr1)
            conn.commit()
##calculation of word count frequency
    #print(sum1)
    for val in count1:
            #print(count1[val])
            fq=count1[val]/sum1*100
            na='C'+str(val)
            s1.write_row(na,[fq])
##adding piechart for the data
    chart = book.add_chart({'type': 'pie'})
    print("adding chart "+s1.name)
    cat='='+s1.name+'!$A$3:$A$6'
    print(cat)
    va='='+s1.name+'!$C$3:$C$6'
    chart.add_series({'name': 'Pie Word frequency',
        'categories': cat,
        'values':  va,
        'points': [
        {'fill': {'color': '#5ABA10'}},
        {'fill': {'color': '#FE110E'}},
        {'fill': {'color': '#CA5C05'}},
        {'fill': {'color': 'yellow'}}
          ]})
    s1.insert_chart('D3', chart , {'x_offset': 25, 'y_offset': 10})
    
book.close()
        
            
