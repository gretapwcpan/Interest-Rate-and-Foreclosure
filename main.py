#remember to set up your driver on local machine before run selenium
# OS: Windows 10

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import glob
import time
import shutil
import pandas as pd
import re
import os
import pyexcel
import re

# =============================================================================
# #get15-year foreclosure data, only need to run this part one time
# opts = Options()
# driver = webdriver.Chrome(executable_path='C:\\Users\\{your user name}\\AppData\\Local\\Continuum\\anaconda3\\Scripts\\chromedriver.exe',chrome_options=opts)
# driver.get('https://pip.moi.gov.tw/V3/A/SCRA0601.aspx')
# xpathlst = range(2, 24)
# raw = glob.glob(r'C:/*/*/Downloads/[*.xls')
# if len(raw) == 0:
#     pass
# else:
#     for item in raw:
#         os.remove(item)
#     
# def download():
#     for i in xpathlst:
#         driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlCity"]/option[{}]'.format(i)).click()
#         #driver.implicitly_wait(10) doesn't work!
#         time.sleep(2)
#         #works!
#         driver.find_element_by_xpath('/html/body/form/div[4]/div[2]/div[1]/div[3]/div/div[2]/div[2]/input[1]').click()
#         time.sleep(2)
#         driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_btnExportExcel"]').click()
#         time.sleep(2)
# test =  download()
# 
# =============================================================================
files = glob.glob(r'C:/*/*/Downloads/[*.xls')
main = pd.DataFrame()
for file in files:
    try:
        a = pd.read_excel(file, encoding = 'utf-8')
        main = main.append(a)
    except:
        break
main['拍定日期'] = main['拍定日期'].astype(str)
main[['年','月','日']] = main['拍定日期'].str.extract('(.*)(.{2})(.{2})') #split a column
main['年/月']=main['年'] + '/' + main['月']

#modify bank rate file with obselet excel format
sheet0 = pyexcel.get_sheet(file_name="C:\\Users\\{your user name}\\Desktop\\foreclosure\\bankrate.xls", name_columns_by_row=0)
xlsarray = sheet0.to_array()
sheet1 = pyexcel.Sheet(xlsarray)
sheet1.save_as("C:\\Users\\{your user name}\\Desktop\\foreclosure\\updatedbankrate.xlsx")
def filter():
    raw = pd.read_excel('updatedbankrate.xlsx', encoding = 'utf-8')
    rate = pd.DataFrame(raw[raw.columns[0]])
    rate['活存機動利率'] = raw[raw.columns[2]]
    rate = rate.drop([0,1,2,3])
    for index, element in enumerate(rate['臺灣銀行\u3000存放款利率歷史資料表']):
        if re.findall('^0',element)!=[]:
            rate['臺灣銀行\u3000存放款利率歷史資料表'][index] = element[1:]
        else:
            pass
    rate['活存機動利率'] = pd.to_numeric(rate['活存機動利率'])
    return rate
rate = filter()
main = pd.merge(main, rate, left_on = '年/月', right_on = '臺灣銀行\u3000存放款利率歷史資料表', how = 'left').drop('臺灣銀行\u3000存放款利率歷史資料表', axis = 1)
grouped = main[['總拍定價格','年/月','活存機動利率']].groupby('年/月').agg(成交量 =('年/月','count'),總拍定價格=('總拍定價格', "sum"),活存機動利率=('活存機動利率', 'mean')).reset_index()
grouped = grouped.reindex(list(range(113,len(grouped)))+list(range(0,113)))
#grouped = main[['總拍定價格','年/月','活存機動利率']].groupby('年/月').agg({'年/月': "count", '總拍定價格': "sum",'活存機動利率': 'mean'})
#.size().reset_index(name='counts')

# multiple line plot
# =============================================================================
# plt.plot( '年/月', '成交量', data=grouped, marker='o', markerfacecolor='blue', markersize=8, color='skyblue', linewidth=4)
# plt.plot( '年/月', '總拍定價格', data=grouped, marker='', color='green', linewidth=2)
# plt.plot( '年/月', '活存機動利率', data=grouped, marker='', color='olive', linewidth=2, linestyle='dashed')
# plt.legend()
# =============================================================================
#double y

fig = plt.figure()
ax1 = fig.add_subplot(111)
ax1.plot(grouped['年/月'], grouped['成交量'])
ax1.set_ylabel('y1')

ax2 = ax1.twinx()
ax2.plot(grouped['年/月'], grouped['活存機動利率'], 'r-')
ax2.set_ylabel('y2', color='r')
for tl in ax2.get_yticklabels():
    tl.set_color('r')

plt.savefig('images/two-scales-5.png')

