# sos_local

This repository contains web scraper to scrap google trends data locally. There are two functions created:
- sos_excel: make change in excel file (keywordlist) and will scrap required keywords from Google Trends and make result in to result.xlsx
    The result is in a format that is very easy to create a Power BI report to visualize the trends
    This is used when there is multiple sets of keywords and updates for existing defined keywords
- sos_app: an app interface (with streamlit) allowing users to input keywords by themselves and download result from the UI
    This is used when testting single sets of keywords



## Before you start

It is recommended that you install anaconda navigator https://www.anaconda.com/products/distribution
*Anaconda has python already set up in its environment. If you have already installed python & pip and your cmd/terminal recognize them, you can skip this step.

It is recommended that you install Power BI desktop https://powerbi.microsoft.com/en-us/downloads/, to visualize the result data.


## Structure

```
/sos_excel
/sos_app
/requirements.txt
/keywordlist.xlsx
```

## Usage
clone/copy this repo to your local folder

Install first packages required
```pip install -r requirements.txt```

And then run the script (sos_excel)
```python sos_excel.py```

or (sos_app)
```streamlit run sos_app.py```


Current this line is commented out in both scripts
```chrome_options.add_argument('--headless')```
if you don't wish to have the chrome driver poping out, please uncomment this line in the python script.


## Project contacts

* Litingaalto
