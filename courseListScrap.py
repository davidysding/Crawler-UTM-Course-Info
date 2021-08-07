from bs4 import BeautifulSoup
import re
import urllib.request
import urllib.error
import xlwt
import sqlite3
import sys

findLink = re.compile(r'<a href="(.*?)">', re.S)
findCode = re.compile(r'<a href=".*">(.*?)</a>', re.S)
findTr = re.compile(r'<tr>(.*?)<tr>')

findTd = re.compile(r'<td>(.*)</td>', re.S)
findTdSec = re.compile(r'</td><td>(.*?)</td>', re.S)
findName = re.compile(r'(.*) <span.*')
findCateg = re.compile(r'.*<abbr title="(.*?)">.*')

def main():

    baseurl = 'https://student.utm.utoronto.ca/calendar/newdep_detail.pl?Depart=4'
    datalist = getData(baseurl)
    # for i in datalist:
    #      print(i)
    #      print("\n\n\n")
    saveData(datalist)


def getData(baseurl):
    datalist = []
    req = urllib.request.Request(baseurl)
    req.add_header(
        'User-Agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.112 Safari/537.36')
    try:
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'lxml')
        tr = soup.find_all('tr')
        for td in tr:
            realData = []  # realData = [code, link, name, categ]
            link = re.findall(findLink, str(td))

            code = re.findall(findCode, str(td))

            if (len(link) != 0 and len(code) != 0):
                realData.append(code[0])
                realData.append(
                    'https://student.utm.utoronto.ca/calendar/'+link[0])

            
            TdSec = re.findall(findTdSec, str(td))
            if (len(TdSec) != 0):
                name = re.findall(findName, str(TdSec[0]))
                realData.append(name[0])
                categ = re.findall(findCateg, str(TdSec[0]))
                realData.append(categ[0])

            datalist.append(realData)
    except urllib.error.HTTPError as e:
        print(e.code)
    return datalist

def saveData(datalist):
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('27', cell_overwrite_ok=True)
    sheet.write(0, 0, 'Course Code')
    sheet.write(0, 1, 'Description Link')
    sheet.write(0, 2, 'Course Name')
    sheet.write(0, 3, 'Credit Category')
    for i in range(len(datalist)):
        for j in range(len(datalist[i])):
            sheet.write(i+1, j, datalist[i][j])
    book.save('data.xls')


if __name__ == '__main__':
    main()
