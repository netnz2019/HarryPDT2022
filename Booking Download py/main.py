#Logs into Booking.com, uses cookies to bypass the 2 factor authentication.
#Downloads the Booking information using a url.
#Converts the xls to xlsx
#sends output to format.py

import settings
import format
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from requests_html import HTMLSession
from selenium.webdriver.chrome.options import Options
from datetime import date


def todaydate():
    return date.today()


def oneyear():
    return date.today().replace(year=date.today().year + 1)


def Main(debug=2, check_in=None, check_out=None, fro=todaydate(), to=oneyear()):
    print("Booking.com")

    try:
        os.remove(r"C:\Users\harry\Desktop\Rstatements\rstatement.xls")
    except:
        pass

    # Instantiate headless driver
    chrome_options = Options()

    # Windows path
    chromedriver_location = "C:/Program Files (x86)/chromedriver.exe"
    # Mac path. May have to allow chromedriver developer in os system prefs

    if debug == 0:
        chrome_options.add_argument("--headless")
    else:
        pass
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument(r"C:\Users\harry\Desktop\Rstatements\chrome_profiles\Default")
    # chrome_options.add_argument('--user-data-dir=C://Rstatements\Default')

    chrome_prefs = {"download.default_directory": r"C:\Users\harry\Desktop\Rstatements"}  # (windows)
    chrome_options.experimental_options["prefs"] = chrome_prefs

    driver = webdriver.Chrome(chromedriver_location, options=chrome_options)

    def Cookie(cookie):
        driver.add_cookie(cookie)

    def login():
        try:
            start = int(driver.current_url.find('ses'))
            end = int(driver.current_url[start:].find('&')) + start
            ses = str(driver.current_url[start + 4: end])

            startdate = fro
            enddate = to

            download_link = 'https://admin.booking.com/fresa/extranet/reservations/download?date_type=arrival&date_to=' + str(
                enddate) + '&date_from=' + str(
                startdate) + '&reservation_status[]=ok&lang=xu&ses=' + ses + '&hotel_id=554570'
            print('downloading from: ' + download_link)
            addcookies()
            driver.get(download_link)

            Download(download_link)
        except:
            print(f'Logging in...')
            driver.get(geturl(1))
            print("Webpage reached...")
            addcookies()
            credintals()

            driver.get(geturl(2))
            time.sleep(2)

            start = int(driver.current_url.find('ses'))
            end = int(driver.current_url[start:].find('&')) + start
            ses = str(driver.current_url[start + 4: end])

            startdate = fro
            enddate = to
            if check_in != None:
                startdate = check_in
                enddate = check_out
                custom = True

            download_link = 'https://admin.booking.com/fresa/extranet/reservations/download?date_type=arrival&date_to=' + str(
                enddate) + '&date_from=' + str(
                startdate) + '&reservation_status[]=ok&lang=xu&ses=' + ses + '&hotel_id=554570'
            print('downloading from: ' + download_link)
            driver.get(download_link)

            Download(download_link)



        print('Successfully logged in!')
        format.main()
        return driver.get_cookies()

    def get_cookies():
        cookies = {}
        selenium_cookies = driver.get_cookies()
        for cookie in selenium_cookies:
            cookies[cookie['name']] = cookie['value']

        return cookies

    def Download(download_link):

        session = HTMLSession()
        response = session.get(download_link, cookies=get_cookies())
        if response.status_code == 200:
            print("SUCCESS")
        else:
            print("Failed to download")

        with open('test4.xls', 'wb') as output:
            output.write(response.content)

    def addcookies():
        # Cookie({"name": 'bkng_sso_ses', "value": "e30"})
        # Cookie({"name": 'auth_token', "value": "5380097289"})
        # Cookie({"name": "bkng_sso_session", "value": "'e30'"})
        # Cookie({"name": "bkng_bfp", "value": '2c0e1aa245bea2d0fc31f83d09399c67'})
        # Cookie({"name": 'ecid', "value": 'zDFhjOTN7BG0pp0j1AwiRQsf'})
        Cookie({"name": 'hauavc', "value": '2EC81835QrxgZ7K4S1o4tCsRNco%2FMczXiGZIHAntIkAmTTGwG2c'})
        # Cookie({"name": 'extranet_cors_js', "value": '1'})
        # Cookie({"name": 'liteha',
        # "value": '%5B%7B%22actions%22%3A%5B%5D%2C%22page%22%3A%22home%22%7D%2C%7B%22actions%22%3A%5B%5D%2C%22page%22%3A%22search_reservations%22%7D%5D'})
        # Cookie({"name": '_ga', "value": 'GA1.2.1575159258.1651914208'})
        print("Cookies added 🍪🍪")

    def credintals():
        login = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_USERNAME_FIELD))
        )

        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_BUTTON))
        )

        login.send_keys(settings.USERNAME)

        login_button.click()

        password = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_PASSWORD_FIELD))
        )

        login_button2 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LOGIN_BUTTON2))
        )

        password.send_keys(settings.PASSWORD)
        login_button2.click()

        link1 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, settings.LINK))
        )

    def geturl(number):
        if number == 1:
            return 'https://account.booking.com/sign-in?op_token=EgVvYXV0aCKnAQoUNlo3Mm9IT2QzNk5uN3prM3BpcmgSCWF1dGhvcml6ZRoaaHR0cHM6Ly9hZG1pbi5ib29raW5nLmNvbS8qYnsicGFnZSI6Ii9ob3RlbC9ob3RlbGFkbWluL2V4dHJhbmV0X25nL21hbmFnZS9ob21lLmh0bWw_bGFuZz14dSZtb2JpbGVfZXh0cmFuZXQ9JmhvdGVsX2lkPTU1NDU3MCJ9QgRjb2RlKhIwoejAk9bHJToAQgBY7_XTkwY'

        elif number == 2:
            return 'https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/search_reservations.html?source=nav&upcoming_reservations=1&hotel_id=554570&lang=xu&ses=850ac26632d089624d31f7080dec6c83&date_from=2022-05-13&date_to=2022-05-14&date_type=arrival'

    login()

    format.main()
