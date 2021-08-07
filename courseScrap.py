from bs4 import BeautifulSoup
import re
import urllib.request
import urllib.error
import xlwt
import xlrd
import sqlite3
import sys

findContent = re.compile(r'(.*?)\[\d{0,99}L')
findContentAlt = re.compile(r'<span class="normaltext">(.*?)<br', re.S)

findExclusion = re.compile(r'Exclusion: (.*?)\n', re.DOTALL)
findPrereq = re.compile(r'Prerequisite: (.*?)\n', re.DOTALL)

def main():
    read_file()
    baseurl = 'https://student.utm.utoronto.ca/calendar/course_detail.pl?Depart=4&Course='
    datalist = []
    actual_url = baseurl+courseCode[38]
    #print(actual_url)
    for i in range(len(courseCode)):
        
        data = getData(baseurl+courseCode[i])
        data.append(courseCode[i])
        datalist.append(data)
    #print(datalist)
    write_file(datalist)

def read_file():
    excel_file_path = "data.xls"
    wb = xlrd.open_workbook(excel_file_path)
    sheet = wb.sheet_by_index(0)

    for row in sheet.col_values(0)[1:]:
        courseCode.append(row)

def write_file(datalist):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Sheet 1')
    for i in range(len(datalist)):
        for j in range(len(datalist[i])):
            worksheet.write(i, j, datalist[i][j])
    workbook.save('descriptions.xls')

def getData(url):
    req = urllib.request.Request(url)
    req.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.76 Safari/537.36')
    data = []
    try: 
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, "lxml")
        content = soup.find("span", {"class": "normaltext"})
        #print(content)
        description = re.findall(findContent, str(content.text))
        if len(description) != 0:
            data.append(description[0])
        else: 
            description = re.findall(findContentAlt, str(content))
            #print(description)
            data.append(description[0].replace('\n', ''))

        exclusion = re.findall(findExclusion, str(content.text))
        
        if len(exclusion) != 0:
            data.append(exclusion[0])
        else:
            data.append('None')
        prereq = re.findall(findPrereq, str(content.text))
        if len(prereq) != 0:
            data.append(prereq[0])
        else:
            data.append('None')
    except urllib.error.HTTPError as e:
        print(e.code)
        print(e.read())
    return data

if __name__ == '__main__':
    courseCode = []
    main()

    