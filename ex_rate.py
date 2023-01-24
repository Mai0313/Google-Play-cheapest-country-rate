from src.model import exp_country, normal_country
from src.utils.countrylist import country, exception
import pandas as pd
import subprocess
import os
import multiprocessing as mp

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
    

def google_drive():
    pass
    # pool = mp.Pool(mp.cpu_count())
    # with mp.Pool(processes = mp.cpu_count()) as p:
    #     print(list(tqdm(p.imap(exp_country, exception), total=len(exception), desc = '我是進度條')))

def exrate():
    start()
    exp_country(exception)
    normal_country(country)
    convert()

def clean():
    driver_clean()

if __name__ == '__main__':
    start()
    exp_country(exception)
    normal_country(country)
    convert()
    driver_clean()