import numpy as np
import datetime
import pandas as pd                        
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
import time
import math
from datetime import date
from io import BytesIO
warnings.filterwarnings("ignore") 
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import glob
import os

dt = date.today()
end = (dt - timedelta(days=dt.weekday()+2)).strftime('%Y-%m-%d')
start = str(int(end.split('-')[0])-1)+'-'+str(int(end.split('-')[1].split('-')[0])).zfill(2)+'-01'
dfk = pd.read_excel('keywordlist.xlsx').drop('Unnamed: 0',1)



def enable_headless_download(browser, download_path):
    browser.command_executor._commands["send_command"] = \
        ("POST", '/session/$sessionId/chromium/send_command')
 
    params = {'cmd': 'Page.setDownloadBehavior',
              'params': {'behavior': 'allow', 'downloadPath': download_path}}
    browser.execute("send_command", params)

    
def download_files(ind, keyword, t= 15):
    download_path = os.getcwd() + str(ind)
    chrome_options = Options()
    download_prefs = {'download.default_directory' : download_path,
                      'download.prompt_for_download' : False,
                      "download.directory_upgrade": True,
                      "safebrowsing.enabled": False,
                      'profile.default_content_settings.popups' : 0}

    chrome_options.add_experimental_option('prefs', download_prefs)
#     chrome_options.add_argument('--headless')
    chrome_options.add_argument('--window-size=1920x1080')
    url = f'https://trends.google.com/trends/explore?date={str(start)}%20{str(end)}&geo=FI&q=' + keyword
    browser = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=chrome_options)
    browser.get(url) 
    enable_headless_download(browser, download_path)
    browser.get(url)
    time.sleep(t)
    button = browser.find_element("css selector",'.widget-actions-item.export')
    button.click()
    time.sleep(15)
    browser.quit()
    
    list_of_files = glob.glob(download_path+"\\*.csv") # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    df = pd.read_csv(latest_file).reset_index()
    if df.shape[0]>1:
        df1=df[1:]
        df1.columns=[i.split(':')[0] if i not in ['Viikko','Week'] else 'date' for i in df.iloc[0]]
        df1['date']=pd.to_datetime(df1['date'])
        df1=df1.replace('<1','1')
        return df1.astype({col: float for col in set(df1.columns)-set(['date'])})
    
def merge2df(df1, df2):
    df=df1.merge(df2, on='date')
    df['trans']=df[df.filter(regex='_x').columns].values/df[df.filter(regex='_y').columns].values
    df['trans']=df['trans'].fillna(df[df.filter(regex='_x').columns].mean()/df[df.filter(regex='_y').columns].mean())
    print(df['trans'].describe())
    for col in df.columns[-(df2.shape[1]-1):-1]:
        if col != 'trans':
            df[col] = df['trans']*df[col]
            df[col] = df[col].fillna(0).apply(lambda x: int(round(x,0)))
    df.drop(df.filter(regex='_y').columns,1, inplace=True)
    df.columns = [ x.split('_x')[0] for x in df.columns]
    return df.drop('trans',1)

def df_result(kw_dict, category):
    df=pd.DataFrame()
    search = list(kw_dict.keys())
    ind = 0
    t=20
    while len(search)>0:
        if df.shape[0]>0:
            key1=(df.set_index('date') == 0).astype(int).sum(axis=0).sort_values(ascending=True).index[0]
            print(f"searching keywords:\n{search[:4]+[key1]}")
#             try:
            df1 = download_files(ind, (',').join(search[:4]+[key1]))
            t+=10
            time.sleep(t)
#             except:
#                 print(f"Retrying... taking 120 more seconds to download")
#                 df1 = download_files(ind, (',').join(search[:4]+[key1]), t=120)
#                 t+=10
#                 time.sleep(t)
            if df1 is not None:
                df = merge2df(df, df1)
            else:
                print(f'No data from below keywordlist:\n{search[:4]+[key1]}')
            ind+=1
            search = list(set(search)-set(search[:4]))
        else:
            print(f"searching keywords:\n{search[:5]}")
#             try:
            df1=download_files(ind, (',').join(search[:5]))
            t+=10
            time.sleep(t)
#             except:
#                 print(f"Retrying... taking 120 more seconds to download")
#                 df1=download_files(ind, (',').join(search[:5]), t=120)
#                 t+=10
#                 time.sleep(t)
            if df1 is not None:
                df = df1.copy()
            else:
                print(f'No data from below keywordlist:\n{search[:5]}')
            ind+=1
            search = list(set(search)-set(search[:5]))
      
    df=df.set_index('date')
    df.columns=df.columns.map(kw_dict)
    df = df.groupby(lambda x:x, axis=1).sum()
    df1 = df.rolling(4).mean().dropna() # smoothing weeklystr
    df2 = df.reset_index()
    df2['date']=df2['date'].dt.to_period('M').astype('str')
    df2 = df2.groupby('date').mean().dropna() # monthly agg
    df3 = df2.rolling(12).mean().dropna() # smooth monthly agg
    df['aggregate'] = 'Weekly'
    df['smooth'] = 0
    df1['aggregate'] = 'Weekly'
    df1['smooth'] = 1
    df2['aggregate'] = 'Monthly'
    df2['smooth'] = 0
    df3['aggregate'] = 'Monthly'
    df3['smooth'] = 1
    dfa = df.append(df1).append(df2).append(df3)
    dfa['category'] = category
    return dfa


dft = dfk[dfk.flag != 1]
print(dft)
try:
    result = pd.read_excel('result.xlsx').drop('Unnamed: 0',1)
except:
    result = pd.DataFrame()



for cat in dft.category.unique():
    print(cat)
    kw_dict = dict(zip(dft[dft.category==cat].keywords, dft[dft.category==cat].afflix))
    df1 = df_result(kw_dict, category = cat)
    dfk.loc[dfk.category == cat, 'flag'] = 1
    if result.shape[0] == 0:
        result = pd.melt(df1.reset_index(), id_vars=['category','aggregate','smooth','date'])
    else:
        result = result.append(pd.melt(df1.reset_index(), id_vars=['category','aggregate','smooth','date']))
    time.sleep(10)
result.to_excel('result.xlsx')
dfk.to_excel('keywordlist.xlsx')