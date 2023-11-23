import sqlite3
import xlsxwriter
import xlrd
import matplotlib.pyplot as plt
import bs4
import requests
url_link="https://www.business-standard.com/stocks/market-statistics/nse/index-components?indices=26771"
header={'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0'}
response=requests.get(url_link,headers=header)
soup=bs4.BeautifulSoup(response.content,'html.parser')
table=soup.find('table').find('tbody').findAll('tr')
indices=[]
level=[]
chg1=[]
chg2=[]
for row in table:

    value=[str(i.text).strip()for i in row .find_all('td')]

    indices.append(value[0])
    level.append(value[1])
    chg1.append(float(value[2]))
    chg2.append(float(value[3]))
Workbook=xlsxwriter.Workbook('BS.xlsx')
worksheet1=Workbook.add_worksheet()
bold=Workbook.add_format({'bold':True})
worksheet1.write('A1','indices',bold)
worksheet1.write('B1','level',bold)
worksheet1.write('C1','chg1',bold)
worksheet1.write('D1','chg2',bold)


row=1
col=0
for i in range(len(indices)):
    worksheet1.write(row,col,indices[i])
    worksheet1.write(row,col+1,level[i])
    worksheet1.write(row,col+2,chg1[i])
    worksheet1.write(row,col+3,chg2[i])
    row=row+1
chart1=Workbook.add_chart({'type':'pie'})
chart1.add_series({'categories':'=sheet1!$A2:$A$10','values':'sheet1!$C$2:$C$12'})
chart1.set_title({'name':'Business Standard'})
worksheet1.insert_chart('J4',chart1)
Workbook.close()

plt.plot(level[0:10],chg1[0:10],color='y',label='chg1')
plt.plot(level[0:10],chg2[0:10],color='r',label='chg2')
plt.title("Business Standard")
plt.xlabel("charge1")
plt.ylabel("charge2")
plt.legend()
plt.show()

wb=xlrd.open_workbook('Bs.xlsx')
worksheet1=wb.sheet_by_name("Sheet1")
num_rows=worksheet1.nrows
num_cols=worksheet1.ncols
coln_review=[]
for curr_row in range(0,num_rows,1):
    row_review=[]
    for curr_col in range(0,num_cols,1):
        review=worksheet1.cell_value(curr_row,curr_col)
        row_review.append(review)
    coln_review.append(row_review)
conn=sqlite3.connect("BS.db")
conn.execute("create table Business(indices Text Not Null,level int Not Null,chg1 int Not Null,chg2 int Not Null);")
cursor=conn.cursor()
cursor.executemany("insert into Business(indices,level,chg1,chg2)values(?,?,?,?)",coln_review)
conn.commit()
conn.close()





