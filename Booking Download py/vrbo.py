# main.py
import csv
import pickle
import settings

import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import *
import time
import glob
import os
from requests_html import HTMLSession
import pandas as pd
from bs4 import BeautifulSoup



def Main():

    options = webdriver.ChromeOptions()
    options.headless = False

    driver = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\chromedriver.exe', options=options)



    def login():
        print("Vrbo login....")
        driver.get(geturl(1))
        print("Webpage reached...")
        addCookies()

        credintals()
        driver.get('https://www.vrbo.com/en-nz/rm/reservations/page/1/sort/stay/asc/filter/all-reservations')
        download()
        time.sleep(5)
        convert()



    def download():

        dowbutton = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.DOWNBUTVR))
        )
        time.sleep(0.5)

        dowbutton.click()
        time.sleep(0.5)
        toggle = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.TOGGLEDVR))
        )
        time.sleep(0.5)
        subbut = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.SUBBUTTON))
        )
        time.sleep(4)
        toggle.click()
        time.sleep(1)
        subbut.click()





    def convert():
        list_of_files = glob.glob('C:/Users/harry/Downloads/*.csv')
        latest_file = max(list_of_files, key=os.path.getctime)
        print(latest_file)
        read_file = pd.read_csv(latest_file)
        read_file.to_excel(r'C:\Users\harry\PycharmProjects\pythonProject10\vrtest.xls', index=None, header=True)

    def addCookies():
        driver.add_cookie({"name":'bm_sv', "value":'8260E5DD57BDC9729F3E7915B099CD29~YAAQj6lgaFXF4OGAAQAA4+jOAg82a0yOJIWRL1uSed/ZmbZzPwQVHWdhkTr+rq65znk7CT0rYrvTQ+8s/S7+3kTnwMMdiScJPZ/OiUB61KO6evTWV7Ksoc12r5dEwnegMYqlHZeJ1U2iQtIVsJi2ZJ/umG3h5dzTRwhjNSW4eYDv2+jjDhttpA54Ua4U9zgl7mvCDHRUQq6NVoRP9G+r3fm0Wd8KC3UHV44AajO8wO8etw3cHcl2++LHY2RVv40=~1'})
        driver.add_cookie({"name":'_clsk' , "value":'plrei|1653609524595|2|0|b.clarity.ms/collect' })
        driver.add_cookie({"name":'site', "value":'homeaway_nz'})
        driver.add_cookie({"name":'DUAID', "value": '3be56189-fcac-469c-a38e-204d9661703d'})
        driver.add_cookie({"name":'xdid', "value":'981cc0ca-be43-4f70-aa0e-59474726a3ed|1653609492|vrbo.com'})
        driver.add_cookie({"name":'EG_SESSIONTOKEN', "value":'h5nAD5oYM_UNXdb4Qce9N_4c4tX_09zR5yfnFRq_qEM:SJVNYTYIbTxZrBbiVgk-WAEKReTQITjtn6SYEt1EL2U'})
        driver.add_cookie({"name":'_gcl_au', "value":'1.1.571117949.1653609490'})
        driver.add_cookie({"name":'HASESSIONV3', "value":'ae86e28d-dad5-480c-adc8-da0007ce0833'})
        driver.add_cookie({"name":'_uetsid', "value":'b1a5d0e0dd4f11eca5878d3b95522f2d'})
        driver.add_cookie({"name":'_clck', "value":'1eiq1oe|1|f1s|0'})
        driver.add_cookie({"name":'2008a337-b332-5e57-a558-0c88dc48b53fUAL', "value":'1'})
        driver.add_cookie({"name":'_fbp', "value":'fb.1.1653609490476.1119927072'})




    def credintals():
        print(settings.LOGIN_USERNAME_FIELDVR)
        login = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_USERNAME_FIELDVR))
        )
        print('hi')

        password = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_PASSWORD_FIELDVR))
        )
        print('1')

        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_BUTTONVR))
        )

        login.send_keys(settings.USERNAMEVR)
        password.send_keys(settings.PASSWORDVR)
        login_button.click()

        print(get_cookies())

    def get_cookies():
        cookies = {}
        selenium_cookies = driver.get_cookies()
        for cookie in selenium_cookies:
            cookies[cookie['name']] = cookie['value']

        return cookies


    def geturl(number):
        if number == 1:
            return 'https://www.bing.com/ck/a?!&&p=3eae8bd3720e48e271ce2d51e0917e72c04f8043a97ee06337f1b855e8cc75f9JmltdHM9MTY1MzM2MTIyOCZpZ3VpZD0zNjIyNzRlOC04NzJhLTQ3MjAtYmJjNy05ZTdmOWNmNmIyYmYmaW5zaWQ9NTE1OQ&ptn=3&fclid=aa897e43-db0d-11ec-a78a-3915893ddad4&u=a1aHR0cHM6Ly9hZG1pbi52cmJvLmNvbS9oYW9kLw&ntb=1'

        elif number == 2:
            return 'https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/search_reservations.html?source=nav&upcoming_reservations=1&hotel_id=554570&lang=xu&ses=850ac26632d089624d31f7080dec6c83&date_from=2022-05-13&date_to=2022-05-14&date_type=arrival'


    login()

