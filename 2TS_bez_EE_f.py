import time
import datetime
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from auth_data import id_password

path = "123.xlsx"  # имя файла

#  n = input(str("Введите номер протокола: "))
#  date_prot = input('введите дату проведения протокола (гггг-мм-дд): ')
#  path_protokol = input('введите путь к протоколу: ')
'''
date_now = datetime.date.today()
date_prot = date_prot.split('-')
day_count = datetime.date(int(date_prot[0]), int(date_prot[1]), int(date_prot[2]))
day_count = date_now - day_count
day_count = str(day_count)
day_count = (day_count.split()[0])
'''

wb_obj = openpyxl.load_workbook(path)  # Открываем файл
sheet_obj = wb_obj.active  # Выбираем активный лист таблицы(
m_row = sheet_obj.max_row

s = Service("/home/odinokov/PycharmProjects/pythonProject2/chromedriver")

url = "https://ideas.nlmk.com/"

options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0")
options.headless = False

driver = webdriver.Chrome(
    service=s,
    options=options
)


def iogin():
    email_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "login")))
    email_input.clear()
    email_input.send_keys("odinokov_da")
    password_input = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "password")))
    password_input.clear()
    password_input.send_keys(id_password)
    #  time.sleep(3)
    password_input.send_keys(Keys.ENTER)
    #  time.sleep(6)


def find_elment_class(n):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, n))).click()

    '''
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "myDynamicElement")))
    element.click()
    '''



def find_element_link(n):
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, n))).click()


def find_element_xpaths(n):
    ds = driver.find_elements(By.XPATH, n)
    return ds




def find_element_xpath(n):
    driver.find_elements(By.XPATH, n).click()


def link_lists():

    my_ul2 = driver.find_elements(By.XPATH, "//ul[@class='b-page']")  # блок со страницам
    if len(my_ul2)==1:
        my_ul = driver.find_element(By.XPATH, "//ul[@class='b-page']")  # блок со страницам
        time.sleep(7)
        all_li = my_ul.find_elements(By.TAG_NAME, "li")  # оперделяем кол-во страниц
        time.sleep(7)
        global link_list
        link_list = []
        for li in all_li:
            y = li.text
            link_list.append(y)

    return link_list
try:
    driver.get(url=url)
    time.sleep(15)
    iogin()
    print('вход')
    find_elment_class('work')
    #  time.sleep(5)
    find_element_link("Все идеи")
    print('все')
    #  time.sleep(5)
    find_element_link("2ТС без ТЭ")
    #  time.sleep(5)
    print('2ТС ТЭ')
    #  time.sleep(5)
    print(type(link_lists()))
    rec = 1
    while rec == 1:
        for i in range(2, m_row + 1):
            cell_obj = sheet_obj.cell(row=i, column=1)  # В column= подставляем номер нужной колонки
            number_ideas = cell_obj.value
            cell_ob = sheet_obj.cell(row=i, column=10)  # В column= подставляем номер нужной колонки
            name_ideas = cell_ob.value
            for y in link_list:
                ds = driver.find_elements(By.XPATH, " // *[ @ href = 'Offer.aspx?id=" + str(number_ideas) + "']")
                print(ds)
                time.sleep(7)


                if len(str(ds)) >= 1:
                    driver.find_element(By.XPATH, " // *[ @ href = 'Offer.aspx?id=" + str(number_ideas) + "']").click()
                    time.sleep(8)
except Exception as ex:
    print(ex)
finally:
    driver.close()
    driver.quit()
