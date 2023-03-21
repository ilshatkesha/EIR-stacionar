import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

dr = webdriver.Firefox()
dr.get('http://10.2.213.165:90/User/Login')

#file_path = 'D:\\stat.txt'

dr.find_element("name", "Login").send_keys('ilshatkesha')
dr.find_element("name", "Password").send_keys('0020275')
dr.find_element("xpath", "//button[@type='submit']").click()

dr.find_element("xpath", "//a[contains(text(),'Подсистемы')]").click()
dr.find_element("xpath", "//a[contains(text(),'Направления на госпитализацию')]").click()
dr.find_element("xpath", "//a[contains(text(),'Направления на госпитализацию (Стационар)')]").click()
#dr.find_element("xpath", "//button[@id='hospital-find']").click()
#dr.find_element("xpath", "//input[@id='gs_DepLpu']").send_keys('пульмологи')



time.sleep(35)

#поиск отделение месяц

dr.find_element("xpath", "//td[@id='pacient-pager_center']/table/tbody/tr/td[8]/select").click()
dr.find_element("xpath", "(//option[@value='10000'])[2]").click()
time.sleep(12)

columns_list = ['ФИО пациента', 'Страховой полис', 'Дата и время Поступления', 'Дата и время Выписки']

df = pd.read_excel(r"D:\\0101-1003.xlsx", sheet_name=0, dtype=str)[columns_list]
#time.sleep(10)
df['Дата и время Поступления'] = pd.to_datetime(df['Дата и время Поступления'], dayfirst=True)
df['Дата и время Выписки'] = pd.to_datetime(df['Дата и время Выписки'], dayfirst=True)

for i in range(0,df.shape[0]):
    p = df['Страховой полис'][i]
    #p = str('0288099732000430')
    try:
        dr.find_element("xpath", f"//td[contains(.,{p})]").click()
    except Exception:
        print(i + 1, " ошибка ", df['ФИО пациента'][i])
        continue
    dr.find_element("xpath", "// td[ @ id = 'pacient-pager_left'] / table / tbody / tr / td / div").click()
    date_in = f"{df['Дата и время Поступления'][i]}".split()[0]
    time.sleep(10)
    dr.find_element("xpath", "// input[ @ id = 'hospital-pacient-edit-gosp-date-fact']").send_keys(date_in) #YYYY-MM-DD
    time_in = f"{df['Дата и время Поступления'][i]}".split()[1]
    time_in = time_in[:5]
    dr.find_element("xpath", "// input[ @ id = 'hospital-pacient-edit-gosp-time-fact']").send_keys(time_in) #HH:mm
    dr.find_element("xpath", "// button[ @ id = 'hospital-pacient-edit-date-time-fact']").click()
    time.sleep(20)
    dr.find_element("xpath", "(//button[@type='button'])[10]").click()
    date_out = f"{df['Дата и время Выписки'][i]}".split()[0]
    dr.find_element("id", "hospital-pacient-edit-out-date").send_keys(date_out)
    time.sleep(8)
    dr.find_element("xpath", "// button[ @ id = 'hospital-pacient-edit-out']").click()
    time.sleep(18)
    dr.find_element("xpath", "(// button[@ type='button'])[10]").click()
    time.sleep(8)
    dr.find_element("xpath", "(//button[@type='button'])[8]").click()
    time.sleep(8)
    print(i + 1, " ", df['ФИО пациента'][i])







