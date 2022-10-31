import streamlit as st
import numpy as np
import datetime
import pandas as pd                        
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
import time
import math
from datetime import date
from io import BytesIO
warnings.filterwarnings("ignore") 

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import glob
import os


st.set_page_config(layout="wide")
st.markdown("<h1 style='text-align: center; color: black;'>Google Trends</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: darkgrey;'>Automates search and save the needed keywords for modelling</h2>", unsafe_allow_html=True)


start = st.date_input(
"start date of comparison",
datetime.date(2020, 10, 1))
end = st.date_input(
    "end date of comparison",
    date.today()- datetime.timedelta(days=date.today().weekday()+2))



st.markdown("<h5 style='text-align: left; color: black;'>\nInput here keywords and their afflixes for Google trends comparison</h1>", unsafe_allow_html=True)

category = st.text_input('Category name for the keywords set for search', value = 'Mobiililaajakaistaliittymät')
st.text(" ")

col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
kw_dict = {}
with col1:
    word = st.text_input('Key Word1', value = "Elisa mobiililaajakaista")
    afflix = st.text_input('Afflix1', value = "Elisa")
    kw_dict[word]=afflix
with col2:
    word = st.text_input('Key Word2', value = "Saunalahti mobiililaajakaista")
    afflix = st.text_input('Afflix2', value = "Elisa")
    kw_dict[word]=afflix
with col3:
    word = st.text_input('Key Word3', value = "DNA mobiililaajakaista")
    afflix = st.text_input('Afflix3', value = "DNA")
    kw_dict[word]=afflix
with col4:
    word = st.text_input('Key Word4', value = "Telia mobiililaajakaista")
    afflix = st.text_input('Afflix4', value = "Telia")
    kw_dict[word]=afflix
with col5:
    word = st.text_input('Key Word5', value = "Sonera mobiililaajakaista")
    afflix = st.text_input('Afflix5', value = "Telia")
    kw_dict[word]=afflix
with col6:
    word = st.text_input('Key Word6', value = "Moi mobiililaajakaista")
    afflix = st.text_input('Afflix6', value = "Moi")
    kw_dict[word]=afflix
with col7:
    word = st.text_input('Key Word7')
    afflix = st.text_input('Afflix7')
    kw_dict[word]=afflix

with st.expander("Input more Keywords"):
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        word = st.text_input('Key Word8')
        afflix = st.text_input('Afflix8')
        kw_dict[word]=afflix
        st.text(" ")
        word = st.text_input('Key Word13')
        afflix = st.text_input('Afflix13')
        kw_dict[word]=afflix
    with col2:
        word = st.text_input('Key Word9')
        afflix = st.text_input('Afflix9')
        kw_dict[word]=afflix
        st.text(" ")
        word = st.text_input('Key Word14')
        afflix = st.text_input('Afflix14')
        kw_dict[word]=afflix
    with col3:
        word = st.text_input('Key Word10')
        afflix = st.text_input('Afflix10')
        kw_dict[word]=afflix
        st.text(" ")
        word = st.text_input('Key Word15')
        afflix = st.text_input('Afflix15')
        kw_dict[word]=afflix
    with col4:
        word = st.text_input('Key Word11')
        afflix = st.text_input('Afflix11')
        kw_dict[word]=afflix
        st.text(" ")
        word = st.text_input('Key Word16')
        afflix = st.text_input('Afflix16')
        kw_dict[word]=afflix
    with col5:
        word = st.text_input('Key Word12')
        afflix = st.text_input('Afflix12')
        kw_dict[word]=afflix
        st.text(" ")
        word = st.text_input('Key Word17')
        afflix = st.text_input('Afflix17')
        kw_dict[word]=afflix

        

kw_dict = {k: v for k, v in kw_dict.items() if v!=''}

def enable_headless_download(browser, download_path):
    # Add missing support for chrome "send_command" to selenium webdriver
    browser.command_executor._commands["send_command"] = \
        ("POST", '/session/$sessionId/chromium/send_command')
 
    params = {'cmd': 'Page.setDownloadBehavior',
              'params': {'behavior': 'allow', 'downloadPath': download_path}}
    browser.execute("send_command", params)

# Add arguments telling Selenium to not actually open a window

def download_files(ind, keyword, t= 10):
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
    # Start up browser
    browser = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=chrome_options)
    browser.get(url) 
    enable_headless_download(browser, download_path)
    # Load webpage
    browser.get(url)
    time.sleep(t)
    button = browser.find_element("css selector",'.widget-actions-item.export')
    button.click()
    time.sleep(5)
    browser.quit()
    
    list_of_files = glob.glob(download_path+"\\*.csv") # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    df = pd.read_csv(latest_file).reset_index()
    print(df)
    if df.shape[0]>1:
        df1=df[1:]
        df1.columns=[i.split(':')[0] if i not in ['Viikko','Week'] else 'date' for i in df.iloc[0]]
        df1['date']=pd.to_datetime(df1['date'])
        return df1.astype({col: float for col in set(df1.columns)-set(['date'])})
    
def merge2df(df1, df2):
    df=df1.merge(df2, on='date')
    df['trans']=df[df.filter(regex='_x').columns].values/df[df.filter(regex='_y').columns].values
    df['trans']=df['trans'].fillna(df['trans'].mean())
    for col in df.columns[-(df2.shape[1]-1):-1]:
        if col != 'trans':
            df[col] = df['trans']*df[col]
            df[col] = df[col].fillna(0).apply(lambda x: int(round(x,0)))
    df.drop(df.filter(regex='_y').columns,1, inplace=True)
    df.columns = [ x.split('_x')[0] for x in df.columns]
    return df.drop('trans',1)

def sos_calculator(kw_dict, category = category):
    df=pd.DataFrame()
    search = list(kw_dict.keys())
    ind = 0
    t=20
    while len(search)>0:
        if df.shape[0]>0:
            key1=(df.set_index('date') == 0).astype(int).sum(axis=0).sort_values(ascending=True).index[0]
            st.text(f"searching keywords:\n{search[:4]+[key1]}")
            try:
                df1 = download_files(ind, (',').join(search[:4]+[key1]))
                t+=10
                time.sleep(t)
            except:
                st.text(f"Retrying... taking 120 more seconds to download")
                df1 = download_files(ind, (',').join(search[:4]+[key1]), t=120)
                t+=10
                time.sleep(t)
            if df1 is not None:
                df = merge2df(df, df1)
            else:
                st.text(f'No data from below keywordlist:\n{search[:4]+[key1]}')
            ind+=1
            search = list(set(search)-set(search[:4]))
        else:
            st.text(f"searching keywords:\n{search[:5]}")
            try:
                df1=download_files(ind, (',').join(search[:5]))
                t+=10
                time.sleep(t)
            except:
                st.text(f"Retrying... taking 120 more seconds to download")
                df1=download_files(ind, (',').join(search[:5]), t=120)
                t+=10
                time.sleep(t)
            if df1 is not None:
                df = df1.copy()
            else:
                st.text(f'No data from below keywordlist:\n{search[:5]}')
            ind+=1
            search = list(set(search)-set(search[:5]))
      
    df=df.set_index('date')
    df0 = df.copy()
    df2 = df0.reset_index()
    df2['date']=df2['date'].dt.to_period('M').astype('str')
    dfm = df2.groupby('date').mean().dropna()
    
    df.columns=df.columns.map(kw_dict)
    df = df.groupby(lambda x:x, axis=1).sum()
    df2 = df.reset_index()
    df2['date']=df2['date'].dt.to_period('M').astype('str')
    df2 = df2.groupby('date').mean().dropna() # monthly agg
    return df0.reset_index(), dfm.reset_index(), df2.reset_index()


if st.button('Calculate Google trends'):
    df_weekly, df_monthly, df = sos_calculator(kw_dict, category = category )
    col1, col2 = st.columns(2)
    with col1:
        fig = px.line(df, x="date", y=[i for i in set(kw_dict.values()) if i in df.columns], title='Share of Search Over Time (Monthly aggregated)')
        
        st.plotly_chart(fig, use_container_width=True)
        ax = sns.lineplot(data=df.set_index('date'))
        plt.xticks(rotation=90)
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, 1.05), ncol=3, fancybox=True)
        plt.savefig("python.png")
        
    with col2:
        df1=pd.DataFrame()
        df1['share']=(df.mean()/(df.mean().sum())).values
        df1['names']=(df.mean()/(df.mean().sum())).index
        fig1 = px.pie(df1, values='share', names='names', title='Share of Search in %',hole=.65)
        #fig1.show()  
        st.plotly_chart(fig1, use_container_width=True)
        

    def to_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df_weekly.to_excel(writer, index=False, sheet_name='weekly_orign')
        df_monthly.to_excel(writer, index=False, sheet_name='monthly_orign')
        workbook = writer.book
        worksheet = workbook.add_worksheet("Plot(Monthly aggregated)")
        worksheet.insert_image('B2', 'python.png', {'x_offset': 15, 'y_offset': 1})
        writer.save()
        processed_data = output.getvalue()
        return processed_data
    df_xlsx = to_excel(df)
    file = st.text_input('Name your file', value='5G google trends')
    st.download_button(label='📥 Download trends data from the selected keywords',
                                    data=df_xlsx ,
                                    file_name= f'{file}.xlsx')
# if st.button('♞ Add current keywords to BI reports! Be sure you add the right one, once it is added, it will be in reporting'):
#     def keyword_excel(kw_dict, cat):
#         df = pd.DataFrame(kw_dict.items(), columns=['keywords','afflix'])
#         df['category'] = cat
#         df['flag'] = 0
#         return df
#     if category not in keyw.category.unique():
#         keyw = keyw.append(keyword_excel(kw_dict, category))
#         keyw.to_excel('keywordlist.xlsx')
#     else:
#         st.markdown("<h5 style='text-align: left; color: black;'>\nCategory already exists, input a new one and add again!</h5>", unsafe_allow_html=True)
# def to_excel2(df):
#     output = BytesIO()
#     writer = pd.ExcelWriter(output, engine='xlsxwriter')
#     keyw.to_excel(writer, index=False, sheet_name='keywords')
#     writer.save()
#     processed_data = output.getvalue()
#     return processed_data
# df_xlsx = to_excel2(keyw)
# st.download_button(label='Download keywords lists for BI report',
#                                 data=df_xlsx ,
#                                 file_name= 'keywordlist.xlsx')
# rm_cat = st.text_input('Remove this category name from reporting list', value='Mobiililaajakaistaliittymät1')
# if st.button('\26 When click this button the category will be removed from BI report! '):
#     if rm_cat not in keyw.category.unique():
#         st.markdown("<h5 style='text-align: left; color: black;'>\nCategory not found from reporting list!</h5>", unsafe_allow_html=True)
#     else:
#         keyw = keyw[keyw.category!=rm_cat]
#         st.markdown("<h5 style='text-align: left; color: black;'>\nCategory removed.\n If you wish to download the updated keywordlist, please refresh the page.</h5>", unsafe_allow_html=True)
#         keyw.to_excel('keywordlist.xlsx')

