from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

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