from selenium.webdriver.support.ui import Select
from dateutil.relativedelta import relativedelta
from selenium.webdriver.common.by import By       
from django.http import HttpResponse 
from django.shortcuts import render           
from selenium import webdriver          
from bs4 import BeautifulSoup
from datetime import datetime
from pandas import DataFrame
from datetime import date 
import pandas as pd
import time

from webdriver_manager.chrome import ChromeDriverManager


def UI(request):
    if request.method == 'POST':
        try:

            BFM = date.today() + relativedelta(months=-5)
            Chrome_options = webdriver.ChromeOptions()
            # Chrome_options.add_argument('--headless')
            # options = Chrome_options.headless
            # chromedriver_path = 'C:/Users/Henry/Downloads/chromedriver_win32/chromedriver.exe'
            driver = webdriver.Chrome(ChromeDriverManager().install(), options=Chrome_options)
            driver.maximize_window()
            driver.get('https://www.sec.gov/edgar/search/')
            time.sleep(2)
            driver.find_element(By.ID, 'entity-short-form').send_keys('Greenlight Capital')
            time.sleep(2)
            driver.find_element(By.ID, 'search').click()
            time.sleep(5)
            driver.find_element(By.ID, 'keywords').clear()
            time.sleep(2)
            driver.find_element(By.ID, 'keywords').send_keys('13F')
            time.sleep(3)
            driver.find_element(By.ID, 'entity-full-form').send_keys('Greenlight Capital')
            time.sleep(5)
            select = Select(driver.find_element(By.ID, 'date-range-select'))
            select.select_by_value('custom')
            time.sleep(2)
            driver.find_element(By.ID, 'date-from').click()
            time.sleep(2)
            select = Select(driver.find_element(By.CLASS_NAME, 'ui-datepicker-month'))
            select.select_by_value(str(int(str(BFM).split('-')[1]) - 1))
            time.sleep(2)
            select = Select(driver.find_element(By.CLASS_NAME, 'ui-datepicker-year'))
            select.select_by_value(str(BFM).split('-')[0])
            time.sleep(2)
            for d in driver.find_elements(By.CLASS_NAME, 'ui-state-default'):
                if d.text == str(int(str(BFM).split('-')[2])):
                    d.click()
                    break
            time.sleep(2)
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            time.sleep(4)
            if len(driver.find_elements(By.CLASS_NAME, 'preview-file')) > 0:
                IFF = []
                IRF = []
                IFEP = []
                for (ff, rf, en) in zip(driver.find_elements(By.CLASS_NAME, 'preview-file'),
                                        driver.find_elements(By.CLASS_NAME, 'enddate')[1:],
                                        driver.find_elements(By.CLASS_NAME, 'entity-name')[1:]):
                    if '13F-HR ' in ff.text:
                        IFF.append(ff.text)
                        IRF.append(rf.text)
                        IFEP.append(en.text)
                if len(IFF) > 0:
                    FF = []
                    RF = []
                    FEP = []
                    for (i, j) in zip(IFEP, range(len(IFEP))):
                        if IFEP.count(i) > 1:
                            if FEP.count(i) == 0:
                                FF.append(IFF[j])
                                FEP.append(IFEP[j])
                                DL = []
                                for k in range(len(IRF)):
                                    if IFEP[k] == i and IFEP[k] != '':
                                        DL.append(
                                            datetime(int(str(IRF[k]).split('-')[0]), int(str(IRF[k]).split('-')[1]),
                                                     int(str(IRF[k]).split('-')[2])))
                                RF.append(str(max(DL)).split(' ')[0])
                        else:
                            FF.append(IFF[j])
                            RF.append(IRF[j])
                            FEP.append(IFEP[j])
                    time.sleep(1)
                    df = DataFrame({'Col1': FF, 'Col2': RF, 'Col3': FEP})
                    response = HttpResponse(content_type='application/xlsx')
                    response['Content-Disposition'] = f'attachment; filename=' + str('Greenlight Capital').replace(' ',
                                                                                                                   '_') + '.xlsx'
                    with pd.ExcelWriter(response) as writer:
                        df.to_excel(writer, sheet_name=str('Greenlight Capital').replace(' ', '_'))
                    return response
                else:
                    return render(request, "index.html", {'message': 'No results for this company!'})
            else:
                return render(request, "index.html", {'message': 'No results for this company!'})        
        except Exception as e:
            print(e)
            return render(request, "index.html", {'message': 'Something bad happened!'})
    else:    
        return render(request, "index.html")