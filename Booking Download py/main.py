# main.py

import settings
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from requests_html import HTMLSession



options = Options()

driver = webdriver.Firefox(executable_path=r'C:\Users\harry\Downloads\geckodriver-v0.31.0-win64\geckodriver.exe', options=options)

def login():
    print(f'Logging in...')
    driver.get('https://account.booking.com/sign-in?op_token=EgVvYXV0aCKnAQoUNlo3Mm9IT2QzNk5uN3prM3BpcmgSCWF1dGhvcml6ZRoaaHR0cHM6Ly9hZG1pbi5ib29raW5nLmNvbS8qYnsicGFnZSI6Ii9ob3RlbC9ob3RlbGFkbWluL2V4dHJhbmV0X25nL21hbmFnZS9ob21lLmh0bWw_bGFuZz14dSZtb2JpbGVfZXh0cmFuZXQ9JmhvdGVsX2lkPTU1NDU3MCJ9QgRjb2RlKhIwoejAk9bHJToAQgBY7_XTkwY')

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

    time.sleep(60)




    print('Successfully logged in!')

def get_cookies():
    cookies = {}
    selenium_cookies = driver.get_cookies()
    for cookie in selenium_cookies:
        cookies[cookie['name']] = cookie['value']
    return cookies

def get_posts():
    session = HTMLSession()

    cookies = {'bkng_sso_session': 'e30', 'bkng_sso_ses': 'e30', 'bkng_bfp': '2c0e1aa245bea2d0fc31f83d09399c67', 'ecid': 'CmqPXS7N7BGm5dGpbos3BAEN', 'pxcts': '618537c8-cd2e-11ec-b29a-45434b4c7967', '_pxvid': '61852b1b-cd2e-11ec-b29a-45434b4c7967', 'hauavc': 'B2D7B87E%2Fk8KXe427k9XnSXXGywhkL6fXLwXt14GZ49r5ujCmQY', 'external_host': 'account.booking.com', 'auth_token': '6040060806', 'uz': 'e', 'extranet_cors_js': '1', 'liteha': '%5B%7B%22actions%22%3A%5B%5D%2C%22page%22%3A%22home%22%7D%5D', 'OptanonConsent': 'isIABGlobal=false&datestamp=Fri+May+06+2022+23%3A19%3A49+GMT%2B1200&version=6.13.0&hosts=&consentId=4f54f154-e2c3-4530-b669-0e8886ac04c7&interactionCount=0&landingPath=https%3A%2F%2Fadmin.booking.com%2Fhotel%2Fhoteladmin%2Fextranet_ng%2Fmanage%2Fhome.html%3Flang%3Dxu%26mobile_extranet%3D%26hotel_id%3D554570%26ses%3Da83d4221fd546bb3566369140f01693e&groups=C0001%3A1%2CC0002%3A1%2CC0004%3A1', '_ga': 'GA1.2.1346332001.1651835964', '_gid': 'GA1.2.1362213366.1651835964', '_gat': '1', '_mkto_trk': 'id:261-NRZ-371&token:_mch-booking.com-1651835990621-28829', '_pxff_ddtc': '1', 'esadm': '02UmFuZG9tSVYkc2RlIyh9YbxZGyl9Y5%2BPTaVrIMNtqKbKeQtDbd4G5WiKEs4tcUaKtPFwP%2FJnw9s%3D', '_px3': '6e0a79e9f56ece0a9a5c1f93e091e68a40622982ca9d3d9c749d139e435cd7a5:kNigu2fK7B6zwrHS2p91WgbTdQE9uD5pjkyrI2T8jTYdE8L/We4j7iU3m/ZerApnHZklwjWvbCoONAd39nUNnQ==:1000:o8rgWPLndZIr/C3BKj9aJTvPIymwXeh79O0Qs3Y2T/bFpbc3pWBynEES0JVD6SXIiT33/B8+KsvJmVKBhbTsb/IdNf3dwS4AnLoWTceEWpnteJjh68pob1a+bf8BcvUaCGkSGbX+/yJR1uAZ5KUeKjp+kcr/5+xiNVuo12mXle916ZCn6N3kRfdGljRMY8Nt/JC7HiXh8W1YbfMy6jmYQw==', '_pxde': 'e7652e6e1e29a2f8f5e2a36372742904b8a63398e722e00b3cdb1fb1a82234c2:eyJ0aW1lc3RhbXAiOjE2NTE4MzU5OTAzOTYsImZfa2IiOjAsImlwY19pZCI6W119'}
    response = session.get('https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/search_reservations.html?upcoming_reservations=1&source=nav&hotel_id=554570&lang=en&ses=552f7ffe6ca62a35aa824d8e6eb15b38&date_from=2022-05-06&date_to=2022-05-07&date_type=arrival', cookies=cookies)
    links = response.html.absolute_links
    response = session.get('https://admin.booking.com/fresa/extranet/reservations/download?date_type=&date_to=2022-05-07&date_from=2022-05-06&hotel_id=554570&ses=a83d4221fd546bb3566369140f01693e&lang=en', cookies=cookies)

    with open('test.xls', 'wb') as output:
        output.write(response.content)

    print(links)
    print(response.status_code)
    return response.text





print(get_posts())

