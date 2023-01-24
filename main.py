from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from src.utils.countrylist import country, exception
import pandas as pd
import subprocess
import os
import multiprocessing as mp
from concurrent.futures import ProcessPoolExecutor

options = Options()
options.add_argument("--headless")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
WINDOW_SIZE = "1920,1080"
options.add_argument("--window-size=%s" % WINDOW_SIZE)
driver = webdriver.Chrome(options=options)

outputfile = 'temp.csv'
complete = f'即時匯率.xlsx'

def exp_country(exception):
    for i, (exception_link, exception_chinese) in enumerate(exception.items()):
        func_num = i + 3
        driver.get(f'https://www.bestxrate.com/card/mastercard/{exception_link}.html')
        visa = driver.find_element(By.ID, "comparison_huilv_Visa").text
        visadate = driver.find_element(By.CSS_SELECTOR, ".even .gray_font").text
        visadate = visadate.replace('(', '').replace(')', '')
        master = (driver.find_element(By.CSS_SELECTOR, ".odd:nth-child(1) > td:nth-child(2)").text)
        masterdate = master.split(' ')[-1].replace('(', '').replace(')', '')
        master = master.split(' ')[0]
        with open(outputfile, 'a+', encoding = 'utf-8') as f:
            f.seek(3)
            f.write(
                str(exception_chinese) + ',' + str(exception_link) + ',' + str('=C2') + ',,' + 
                str(visadate) + ',' + str(visa) + ',' + str(f'=B2*C{func_num}*F{func_num}') + ',,' +
                str(masterdate) + ',' + str(master) + ',' + str(f'=B2*C{func_num}*J{func_num}') + ',,,' +
                '\n')
        print(exception_chinese, ' has been processed')

def normal_country(country):
    for i, (country_link, country_chinese) in enumerate(country.items()):
        func_num = i + 5
        driver.get(f'https://www.bestxrate.com/card/mastercard/{country_link}.html')
        visa = driver.find_element(By.ID, "comparison_huilv_Visa").text
        visadate = driver.find_element(By.CSS_SELECTOR, ".even .gray_font").text
        visadate = visadate.replace('(', '').replace(')', '')
        master = (driver.find_element(By.CSS_SELECTOR, ".odd:nth-child(1) > td:nth-child(2)").text)
        masterdate = master.split(' ')[-1].replace('(', '').replace(')', '')
        master = master.split(' ')[0]
        jcb = (driver.find_element(By.CSS_SELECTOR, ".odd:nth-child(3) > td:nth-child(2)").text)
        jcbdate = jcb.split(' ')[-1].replace('(', '').replace(')', '')
        jcb = jcb.split(' ')[0]
        with open(outputfile, 'a+', encoding = 'utf-8') as f:
            f.seek(4)
            f.write(
                str(country_chinese) + ',' + str(country_link) + ',' + str('=C2') + ',,' +
                str(visadate) + ',' + str(visa) + ',' + str(f'=B2*C{func_num}*F{func_num}') + ',,' +
                str(masterdate) + ',' + str(master) + ',' + str(f'=B2*C{func_num}*J{func_num}') + ',,' +
                str(jcbdate) + ',' + str(jcb) + ',' + str(f'=B2*C{func_num}*N{func_num}') + '' +
                '\n')
        print(country_chinese, ' has been processed')

outputfile = 'temp.csv'
complete = f'即時匯率.xlsx'
def start():
    if os.path.exists(outputfile):
        os.remove(outputfile)
    with open(outputfile, 'a+', encoding = 'utf-8') as f:
        f.write(('國家') + ',' + str('手續費') + ',' + str('金額') + ',,,' + str('VISA') + ',,,,' + str('Master')  + ',,,,' + str('JCB') + ',' + '\n')
        f.write(str('手動輸入') + ',' + str('1.5') + ',' + str('金額輸入...')  + ',,' + str('更新日期') + ',' + str('匯率') + ',' + str('試算結果') + ',,' + str('更新日期') + ',' + 
                str('匯率') + ',' + str('試算結果') + ',,' + str('更新日期') + ',' + str('匯率') + ',' + str('試算結果') + '\n')

def convert():
    data = pd.read_csv(outputfile, sep=',', encoding='utf-8', header = None)
    data.to_excel(complete, index=False, header=False)
    os.remove(outputfile)

def driver_clean():
    subprocess.call("taskkill /im chrome.exe /f")

if __name__ == '__main__':
    start()
    exp_country(exception)
    normal_country(country)
    convert()
    driver_clean()