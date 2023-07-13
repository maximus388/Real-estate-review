"""
Выгрузка готовых обзоров на сайт Этажи 
Ссылка: https://www.etagi.com/analytics/
"""


import locale
import datetime
import re, ast
from sqlalchemy import create_engine
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from pdf2image import convert_from_path
import progressbar

locale.setlocale(locale.LC_ALL, '')

MONTH = 6     # Месяц обзора
YEAR = 2023
review_date = datetime.datetime(YEAR, MONTH, 1)

URL = "https://ries3.etagi.com/sitebuilder/analitics/pdf/view/"
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)             # Оставлять браузер открытым

MONTH = '%02d' % (MONTH)
YEAR = str(YEAR)

REVIEW_FOLDER_PATH = os.path.join("*\\Обзоры", YEAR, MONTH + '. ' + review_date.strftime("%B"))
COVER_FOLDER_PATH = os.path.join("*\\Обложки", YEAR, MONTH + '. ' + review_date.strftime("%B"))


"""
Генерация обложек обзоров из их первых страниц
"""
def generate_covers(REVIEW_FOLDER_PATH, COVER_FOLDER_PATH):
    os.chdir(REVIEW_FOLDER_PATH)
    if not os.path.exists(COVER_FOLDER_PATH):
        os.makedirs(COVER_FOLDER_PATH)
    with progressbar.ProgressBar(max_value = len(os.listdir())) as bar:
        for review in os.listdir():
            if not os.path.exists(COVER_FOLDER_PATH + '\\' + review.split(' - ')[0] + '.jpg'):
                writer = PdfFileWriter()
                reader = PdfFileReader(os.getcwd() + '\\' + review)
                page = reader.getPage(0)
                writer.addPage(page)
                with open(COVER_FOLDER_PATH + '\\' + review.split(' - ')[0] + '.pdf', 'wb') as output:
                    writer.write(output)    
                image = convert_from_path(COVER_FOLDER_PATH + '\\' + review.split(' - ')[0] + '.pdf', poppler_path = r"C:\Program Files\poppler-22.04.0\Library\bin")
                for i in range(len(image)):
                    image[i].save(COVER_FOLDER_PATH + '\\' + review.split(' - ')[0] +'.jpg', 'JPEG')
                    os.remove(COVER_FOLDER_PATH + '\\' + review.split(' - ')[0] + '.pdf')
            bar.update(os.listdir().index(review))

"""
Выгрузка обзоров на сайт
"""
def upload_reviews(url, REVIEW_FOLDER_PATH, COVER_FOLDER_PATH):
    try:
        with open(r'*\connection str RIES.txt', 'r') as f:
            CONNECT = f.read()
            CONNECT = ast.literal_eval(re.sub(r'\n','', CONNECT))
    except FileNotFoundError:
        with open(fr'*\connection str RIES.txt', 'r') as f:
            CONNECT = f.read()
            CONNECT = ast.literal_eval(re.sub(r'\n','', CONNECT))
    con = create_engine(f"mysql+pymysql://{CONNECT['user']}:{CONNECT['password']}@{CONNECT['host']}/{CONNECT['database']}")
    
    df_site_reviews = pd.read_sql("""<QUERY>""", con = con)
    
    os.chdir(REVIEW_FOLDER_PATH)
    caps = DesiredCapabilities().CHROME
    caps["pageLoadStrategy"] = "none"
#     browser = webdriver.Chrome(executable_path = r'C:\Program Files (x86)\Google\ChromeDriver\chromedriver 108.exe', chrome_options = chrome_options, desired_capabilities = caps) 
    browser = webdriver.Chrome(service = Service(ChromeDriverManager().install()), chrome_options = chrome_options, desired_capabilities = caps)
    browser.get(url)
    username = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@name='username']"))).send_keys('<username>')
    password = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@name='password']"))).send_keys('<password>')
    submit_auth_btn = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="input-form"]/div[2]/button[1]"""))).click()
    sorted_reviews = sorted(os.listdir(), reverse = True)
    with progressbar.ProgressBar(max_value = len(sorted_reviews)) as bar:
        for review in sorted_reviews[]:
            if review not in list(df_site_reviews['filename']):
                new_review_btn = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="content"]/a/span"""))).click()
                review_name = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, """//*[@id="pdf_form"]/label[1]/input"""))).send_keys(review)
                review_city = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, """//*[@id="pdf_form"]/label[2]/select"""))).send_keys(review.split(' - ')[0])
                class_btn = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, """//*[@id="pdf_form"]/label[3]/select"""))).send_keys('Вторичное жилье')
                upload_review =  WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, """//*[@id="uploadFile"]"""))).send_keys(REVIEW_FOLDER_PATH + '\\' + str(review))
                upload_cover = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, """//*[@id="uploadImg"]"""))).send_keys(COVER_FOLDER_PATH + '\\' + review.split(' - ')[0] + '.JPG')
                review_date = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, """//*[@id="date_w"]"""))).send_keys(f"""01.{MONTH}.{YEAR}""")
                upload_btn = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="pdf_form"]/button/span"""))).click()
                bar.update(sorted_reviews.index(review))
            else:
                bar.update(sorted_reviews.index(review))
                pass
        
    back_to_homepage = WebDriverWait(browser, 20).until(EC.element_to_be_clickable((By.XPATH, """//*[@id="head_sideLeft"]/a/img"""))).click()
    
def main():
    print('Генерация обложек обзоров...')
    generate_covers(REVIEW_FOLDER_PATH, COVER_FOLDER_PATH)
    print('Выгрузка обзоров в RIES...')
    upload_reviews(URL, REVIEW_FOLDER_PATH, COVER_FOLDER_PATH)


if __name__ == '__main__':
    main()
    print('Скрипт выполнен!')
