# main.py
import csv
import pickle
import settings
from openpyxl import Workbook
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
import pickle
from selenium.webdriver.chrome.options import Options
from datetime import date


def Main():

    try:
        os.remove(r"C:\Users\harry\Desktop\Rstatements\Reservations.csv")
    except:
        pass

    # Instantiate headless driver
    chrome_options = Options()

    # Windows path
    chromedriver_location = "C:/Program Files (x86)/chromedriver.exe"
    # Mac path. May have to allow chromedriver developer in os system prefs


    #chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    chrome_prefs = {"download.default_directory": r"C:\Users\harry\Desktop\Rstatements"}  # (windows)
    chrome_options.experimental_options["prefs"] = chrome_prefs

    driver = webdriver.Chrome(chromedriver_location, options=chrome_options)




    def addcookes(cookies):
        for cookie in cookies:
            print(cookie)
            driver.add_cookie(cookie)

    def login():

        #file = open("DictFile.pkl", "rb")
        #file_contents = pickle.load(file)
        #addcookes(file_contents)

        list_of_files = glob.glob('C:/Users/harry/Downloads/*.csv')
        lastDown = max(list_of_files, key=os.path.getctime)
        print("Vrbo login....")
        driver.get(geturl(1))
        print("Webpage reached...")
        cookieJar()
        credintals()
        time.sleep(4)
        download()
        time.sleep(10)



        convert()



    def download():
        today = date.today()

        oneyear = date.today().replace(year=date.today().year + 1)

        print(date.today())
        driver.get('https://www.vrbo.com/rm/proxies/v2/conversations/export?afterDate='+ str(today) + '&beforeDate='+ str(oneyear) + '&csrfToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0b2tlbiI6IjEwY2IyMzE4NGZhZWNiN2ExN2FmMzdlN2NmMzJjOTlkMjFjYTFkMmVhMGRkMmQyNjRkYTBlMWU0NmI0YzY0ZjM5OGNiM2E4MTJmZDMyNTMxNDk0ZmZhODg1NWE4MGMyODkyMzM5MjNmYTlkZGZlYTNjNmM4MjkzMDU1ZTM1ODhjZjBkOGUxZjE1ZDQxMTEwMTcyYzRmMWMwZWVkMzg1ZTY3ZjdmMjhjNGI1YjNlODQ2Nzg3ZDVhOGI3YjJmZGI3OWJiYTE0NDFjY2YwNTg1IiwiaWF0IjoxNjU0NzMyMjM4LCJleHAiOjE2NTUzMzcwMzh9.njbcO9AQLoZlqbwkF80TcQdPo1yfH1aMbnLLzvEZw7U&druid=&reservations=true&site=homeaway_nz&status=RESERVATION_DOWNLOADABLE')







    def convert():
        wb = Workbook()
        try:
            os.remove(r"C:\Users\harry\Desktop\Rstatements\Vrbo.xls")
        except:
            pass

        def create_workbook(path):
            workbook = Workbook()
            workbook.save(path)

        create_workbook(r"C:\Users\harry\Desktop\Rstatements\Vrbo.xlsx")


        read_file = pd.read_csv(r"C:\Users\harry\Desktop\Rstatements\Reservations.csv")
        read_file.to_excel(r"C:\Users\harry\Desktop\Rstatements\Vrbo.xlsx", index=None, header=True)
        



    def cookieJar():
        driver.add_cookie({"name":'bm_sv', "value":'8260E5DD57BDC9729F3E7915B099CD29~YAAQj6lgaFXF4OGAAQAA4+jOAg82a0yOJIWRL1uSed/ZmbZzPwQVHWdhkTr+rq65znk7CT0rYrvTQ+8s/S7+3kTnwMMdiScJPZ/OiUB61KO6evTWV7Ksoc12r5dEwnegMYqlHZeJ1U2iQtIVsJi2ZJ/umG3h5dzTRwhjNSW4eYDv2+jjDhttpA54Ua4U9zgl7mvCDHRUQq6NVoRP9G+r3fm0Wd8KC3UHV44AajO8wO8etw3cHcl2++LHY2RVv40=~1'})
        driver.add_cookie({"name":'_clsk' , "value":'plrei|1653609524595|2|0|b.clarity.ms/collect' })
        driver.add_cookie({"name":'site', "value":'homeaway_nz'})
        driver.add_cookie({"name":'DUAID', "value": '3be56189-fcac-469c-a38e-204d9661703d'})
        driver.add_cookie({"name":'xdid', "value":'981cc0ca-be43-4f70-aa0e-59474726a3ed|1653609492|vrbo.com'})
        #driver.add_cookie({"name":'EG_SESSIONTOKEN', "value":'h5nAD5oYM_UNXdb4Qce9N_4c4tX_09zR5yfnFRq_qEM:SJVNYTYIbTxZrBbiVgk-WAEKReTQITjtn6SYEt1EL2U'})
        driver.add_cookie({"name":'_gcl_au', "value":'1.1.571117949.1653609490'})
        #driver.add_cookie({"name":'HASESSIONV3', "value":'ae86e28d-dad5-480c-adc8-da0007ce0833'})
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

        file = open("DictFile.pkl", "wb")
        pickle.dump(driver.get_cookies(), file)
        file.close()

        file = open("DictFile.pkl", "rb")
        file_contents = pickle.load(file)
        print(file_contents)








    def geturl(number):
        if number == 1:
            return 'https://www.bing.com/ck/a?!&&p=3eae8bd3720e48e271ce2d51e0917e72c04f8043a97ee06337f1b855e8cc75f9JmltdHM9MTY1MzM2MTIyOCZpZ3VpZD0zNjIyNzRlOC04NzJhLTQ3MjAtYmJjNy05ZTdmOWNmNmIyYmYmaW5zaWQ9NTE1OQ&ptn=3&fclid=aa897e43-db0d-11ec-a78a-3915893ddad4&u=a1aHR0cHM6Ly9hZG1pbi52cmJvLmNvbS9oYW9kLw&ntb=1'

        elif number == 2:
            return 'https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/search_reservations.html?source=nav&upcoming_reservations=1&hotel_id=554570&lang=xu&ses=850ac26632d089624d31f7080dec6c83&date_from=2022-05-13&date_to=2022-05-14&date_type=arrival'

        elif number == 3:
            return 'https://www.vrbo.com/rm/proxies/v2/conversations/export?afterDate=2022-06-10&beforeDate=2022-07-15&csrfToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0b2tlbiI6IjEwY2IyMzE4NGZhZWNiN2ExN2FmMzdlN2NmMzJjOTlkMjFjYTFkMmVhMGRkMmQyNjRkYTBlMWU0NmI0YzY0ZjM5OGNiM2E4MTJmZDMyNTMxNDk0ZmZhODg1NWE4MGMyODkyMzM5MjNmYTlkZGZlYTNjNmM4MjkzMDU1ZTM1ODhjZjBkOGUxZjE1ZDQxMTEwMTcyYzRmMWMwZWVkMzg1ZTY3ZjdmMjhjNGI1YjNlODQ2Nzg3ZDVhOGI3YjJmZGI3OWJiYTE0NDFjY2YwNTg1IiwiaWF0IjoxNjU0NzMyMjM4LCJleHAiOjE2NTUzMzcwMzh9.njbcO9AQLoZlqbwkF80TcQdPo1yfH1aMbnLLzvEZw7U&druid=&reservations=true&site=homeaway_nz&status=RESERVATION_DOWNLOADABLE'

    login()



