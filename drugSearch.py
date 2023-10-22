# import module
# install module in TERMINAL first (pip is the package installer for Python. You can use pip to install packages)
import pandas as pd 
import requests #requests 抓取網頁的資料
import time


# 抽參數
input = "inputBrandName.xlsx"
output = "outputListPrice.xlsx"


# main code
df = pd.read_excel(input) # 強制吃input excel

result = [] #開空list

for row in df.index: #查詢 each row (以index標記)
#for col in df.columns:
    # 整理request payload
    drugName = df['Name'][row] #pd寫法，彈性寫法是 df[col][row]
    requestPayload = {'Name': drugName,'Type[]':'迄今'}
    url = 'https://www.nhi.gov.tw/QueryN_New/QueryN/Query1List' #request POST url

    # request
    response = requests.post(url, data = requestPayload)
    data = response.json().get('data') # data是preview response; response.JSON() converts http content into data to get()

    result.extend(data) # for loop iterable result, result放外面

    time.sleep(0.5) 
    # 很重要，要防止DDOS


# Output path
dfResult = pd.DataFrame(result)
dfResult.to_excel(output, index=False)
print('執行腳本完成')
