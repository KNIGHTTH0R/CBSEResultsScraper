import urllib
import urllib2
from urllib2 import URLError
from bs4 import BeautifulSoup
import xlwt
import re


numberstoscan = #100
start_roll = #9200000


subjects = []

url = 'http://www.cbseresults.nic.in/class12/cbse122015_all.asp'
user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
headers = { 'User-Agent' : user_agent,
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Content-Type' : 'application/x-www-form-urlencoded',
            'Connection':'keep-alive',
            'Referer': 'http://www.cbseresults.nic.in/class12/cbse122015_all.htm',
            'Origin': 'http://www.cbseresults.nic.in'
            }

wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
ws.write(0,0,"Roll No")
ws.write(0,1,"Name")
ws.write(0,2,"Mother's Name")
ws.write(0,3,"Father's Name")

basecol = 4

for inx in range (2,numberstoscan + 2):
    print "----------"
    rollno = start_roll - 2 + inx
    values = {'regno' : str(rollno) ,'B1':'Submit'}
    data = urllib.urlencode(values)
    req = urllib2.Request(url, data,headers)
    print rollno
    error = "[Errno 10060] A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond"
    while error == "[Errno 10060] A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond":
        print "Attempt" 
        try:
            response = urllib2.urlopen(req)
            the_page = response.read()
            soup = BeautifulSoup(the_page)
            table = soup.find_all("table", { "width" : "75%" })

            table1 = table[1]
            rows = table1.findAll("tr")
            rollno = rows[0].findAll("td")[1].find("font").text
            name = rows[1].findAll("td")[1].find("font").find("b").text
            mot  = rows[2].findAll("td")[1].find("font").text
            fat = rows[3].findAll("td")[1].find("font").text
            ws.write(inx,0,rollno)
            ws.write(inx,1,name)
            ws.write(inx,2,mot)
            ws.write(inx,3,fat)
            print name

            
            table2 = table[2]
            rows = table2.findAll("tr")
            rows.pop()
            rows.pop(0)
            rows[:] = [row for row in rows if ''.join(re.findall('[a-zA-Z]', row.find("td").find("font").text)) != 'AdditionalSubject']
            for row in rows:
                cols = row.find_all("td")
                scode = int(cols[0].find("font").text)
                i = subjects.index(scode) if scode in subjects else None
                if i == None:
                    i = len(subjects)
                    subjects.append(scode)
                    ws.write(0,basecol+(4*i),cols[1].find("font").text)
                    ws.write(1,basecol+(4*i)+0,"Theory")
                    ws.write(1,basecol+(4*i)+1,"Practical")
                    ws.write(1,basecol+(4*i)+2,"Marks")
                    ws.write(1,basecol+(4*i)+3,"Grade")
                theory = cols[2].text
                prac = cols[3].text
                marks = cols[4].text
                grade = cols[5].text
                col = basecol+(4*i)
                ws.write(inx,col+0,theory)
                ws.write(inx,col+1,prac)
                ws.write(inx,col+2,marks)
                ws.write(inx,col+3,grade)
            error = ""
        except URLError as e:
            print e.reason
    wb.save("Data.xls")
