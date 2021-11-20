from os import write
from selenium import webdriver
import time
import pandas as pd
import numpy as np

PATH = "C:\Program Files (x86)\chromedriver.exe"
##driver = webdriver.Chrome(PATH)
##driver.maximize_window()

ef = pd.read_excel('List of Stocks case study (1).xlsx')
s = ef['SYMBOL'].tolist()
for i in range(len(s)):
    s[i] = s[i].strip()


for f in range(len(s)):
   
    driver = webdriver.Chrome(PATH)
    driver.get("https://finance.yahoo.com/quote/{}.NS/history?p={}.NS".format(s[f],s[f]))
    d = []
    p = []
    
    

    #link = driver.find_element_by_xpath('//*[@id="subtab-equity"]/div/div[3]/nav/div/div/a[5]')
    #link.click()
    #time.sleep(3)
    #link1 = driver.find_element_by_xpath('//*[@id="historical-trade"]/section/div/div[1]/div/div[1]/ul/li[4]/a')

    #link1.click()
    #link2 = driver.find_element_by_xpath('//*[@id="equity-historical-Date-filter"]/div[3]/button')
    #driver.execute_script("arguments[0].click();", link2)

        #search = driver.find_element_by_id('priceInfoTable')
    i = 0
    for i in range(89):
        try:
            z = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[2]/table/tbody/tr[' + str(i+1) +']/td[1]/span')
            x = driver.find_element_by_xpath('//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[2]/table/tbody/tr[' + str(i+1) +']/td[5]/span')
        except:
            continue
        t = x.text
        t = t.replace(",","")
        p.append(float(t))
        d.append(z.text)

       
    driver.delete_all_cookies()
    
    df = pd.DataFrame(({'DATE':d,'PRICE':p}))

    data = pd.ExcelWriter(s[f] + '.xlsx')
    df.to_excel(data, 'Sheet1',index=False)
    data.save()
    driver.quit()


##//*[@id="Col1-1-HistoricalDataTable-Proxy"]/section/div[2]/table/tbody/tr[1]/td[5]/span
driver.quit()
    


#print(s[2])

