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
import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

country = {'USD': '美金', 'EUR': '歐元', 'JPY': '日圓', 'CNY': '人民幣', 'HKD': '港幣', 'SGD': '新加坡幣', 'GBP': '英鎊'
, 'AUD': '澳幣', 'CAD': '加拿大幣', 'NZD': '紐西蘭幣', 'MOP': '澳門幣', 'THB': '泰幣', 'KRW': '韓幣', 'PHP': '菲律賓幣'
, 'IDR': '印尼幣', 'MYR': '馬來幣', 'BND': '汶萊幣', 'INR': '印度盧比', 'CHF': '瑞士法郎', 'SEK': '瑞典幣', 'NOK': '挪威克朗'
, 'DKK': '丹麥克朗', 'RUB': '俄國盧布', 'ZAR': '南非幣', 'AED': '阿拉伯聯合酋長國幣', 'AFN': '阿富汗幣', 'ALL': '阿爾巴尼亞幣'
, 'AMD': '亞美尼亞幣', 'ANG': '荷屬安地列斯幣', 'AOA': '安哥拉幣', 'ARS': '阿根廷比索', 'AWG': '阿魯巴幣', 'AZN': '亞塞拜然幣'
, 'BAM': '波斯尼亞及黑塞哥維那幣', 'BBD': '巴巴多斯幣', 'BDT': '孟加拉幣', 'BGN': '保加利亞幣', 'BHD': '巴林幣', 'BIF': '布隆迪法郎'
, 'BMD': '百慕達幣', 'BOB': '玻利維亞幣', 'BRL': '巴西幣', 'BSD': '巴哈馬幣', 'BTN': '不丹幣', 'BWP': '博茨瓦納幣', 'BZD': '貝里斯幣'
, 'CDF': '剛果法郎', 'CLP': '智利比索', 'COP': '哥倫比亞比索', 'CRC': '哥斯大黎加幣', 'CVE': '佛得角幣', 'CZK': '捷克克朗', 'DJF': '吉布地法郎'
, 'DOP': '多米尼加比索', 'DZD': '阿爾及利亞幣', 'EGP': '埃及鎊', 'ETB': '衣索比亞幣', 'FJD': '斐濟幣', 'FKP': '福克蘭群島鎊', 'GEL': '喬治亞幣'
, 'GHS': '迦納幣', 'GIP': '直布羅陀鎊', 'GMD': '甘比亞幣', 'GNF': '幾內亞法郎', 'GTQ': '瓜地馬拉幣', 'GYD': '圭亞那幣', 'HNL': '宏都拉斯幣'
, 'HRK': '克羅埃西亞幣', 'HTG': '海地幣', 'HUF': '匈牙利幣', 'ILS': '以色列幣', 'IQD': '伊拉克幣', 'ISK': '冰島克朗', 'JMD': '牙買加幣'
, 'JOD': '約旦幣', 'KES': '肯尼亞幣', 'KGS': '吉爾吉斯斯坦幣', 'KHR': '柬埔寨幣', 'KMF': '科摩羅法郎', 'KWD': '科威特幣', 'KYD': '開曼群島幣'
, 'KZT': '哈薩克堅戈', 'LAK': '寮幣', 'LKR': '斯里蘭卡盧比', 'LRD': '賴比瑞亞幣', 'LSL': '賴索托幣', 'LYD': '利比亞幣', 'MAD': '摩洛哥幣'
, 'MDL': '摩爾多瓦幣', 'MGA': '馬達加斯加幣', 'MKD': '馬其頓幣', 'MMK': '緬甸幣', 'MNT': '蒙古幣', 'MUR': '模里西斯盧比', 'MVR': '馬爾地夫幣'
, 'MWK': '馬拉威幣', 'MXN': '墨西哥比索', 'MZN': '莫三比克幣', 'NAD': '納米比亞幣', 'NGN': '尼日利亞幣', 'NIO': '尼加拉瓜幣', 'NPR': '尼泊爾盧比'
, 'OMR': '阿曼幣', 'PAB': '巴拿馬幣', 'PEN': '秘魯幣', 'PGK': '巴布亞紐幾內亞幣', 'PKR': '巴基斯坦盧比', 'PLN': '波蘭幣', 'PYG': '巴拉圭幣'
, 'QAR': '卡塔爾幣', 'RON': '羅馬尼亞幣', 'RSD': '塞爾維亞幣', 'RWF': '盧旺達法郎', 'SAR': '沙烏地阿拉伯幣', 'SBD': '索羅門群島幣', 'SCR': '塞席爾盧比'
, 'SHP': '聖赫勒拿鎊', 'SOS': '索馬利亞幣', 'SRD': '蘇利南幣', 'SSP': '南蘇丹鎊', 'SVC': '薩爾瓦多科朗', 'SZL': '史瓦濟蘭幣', 'TMT': '土庫曼幣'
, 'TND': '突尼西亞幣', 'TOP': '東加幣', 'TRY': '土耳其里拉', 'TTD': '特立尼達及多巴哥幣', 'TZS': '坦桑尼亞幣', 'UAH': '烏克蘭幣', 'UGX': '烏干達幣'
, 'UYU': '烏拉圭比索', 'UZS': '烏玆別克斯坦幣', 'VND': '越南幣', 'VUV': '瓦努阿圖幣', 'WST': '薩摩亞幣', 'XAF': '中非法郎', 'XCD': '東加勒比海幣'
, 'XOF': '西非法郎', 'XPF': '太平洋法郎', 'YER': '葉門幣', 'ZMW': '尚比亞克瓦查'}

exception = {'LBP': '黎巴嫩鎊', 'TJS': '塔吉克斯坦幣'}

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

def on_exchange():
    progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="indeterminate")
    progress.pack()
    progress.start()
    start()
    print('Start Finding Exchange Rate for Excepted Countries')
    exp_country(exception)
    print('Start Finding Exchange Rate for Normal Countries')
    normal_country(country)
    convert()
    progress.stop()
    progress.destroy()
    print('Task has been completed')
    label.config(text='以更新匯率列表')

def on_exist():
    progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="indeterminate")
    progress.pack()
    progress.start()
    print('Cleaning Chrome Driver')
    driver_clean()
    progress.stop()
    progress.destroy()
    print('Done!')
    label.config(text='已關閉瀏覽器並清理記憶體')

if __name__ == '__main__':
    # root = tk.Tk()
    root = ttk.Window(
            title="Rate Exchange",
            themename="darkly",
            size=(1066,600),
            position=(100,100),
            minsize=(0,0),
            maxsize=(1920,1080),
            resizable=None,
            alpha=1.0,
            )
    # root.place_window_center()    #居中
    # root.resizable(False,False)   #不可更改大小
    # root.wm_attributes('-topmost', 1)#其它窗口之上

    root.iconbitmap("src\ico.ico")
    root.title("Rate Exchange")
    root.geometry("400x200")
    root.wm_attributes('-topmost', 1)
    root.place_window_center()
    
    start_button = tk.Button(root, text="開始查詢匯率", command=lambda: root.after(0, on_exchange))
    start_button.pack()
    
    exist_button = tk.Button(root, text="清理記憶體/Chrome殘留", command=lambda: root.after(0, on_exist))
    exist_button.pack()

    # label = tk.Label(root, text="請選擇動作")
    frame = tk.Frame(root, width=200, height=100)
    frame.pack()

    label = tk.Label(frame)
    label.pack()

    root.mainloop()

    # start()
    # exp_country(exception)
    # normal_country(country)
    # convert()
    # driver_clean()