import pandas as pd
import numpy as np
import os
import re
from datetime import date

# 設置路徑
domainPath = os.path.dirname( os.path.realpath( __file__ ))
dataPath = domainPath + '\\data\\'
# print(dataPath)

# 讀入科系_學院配對excel
deptCollegeDict = pd.read_excel(dataPath + '0_科系_學院分類.xlsx')
deptCollegeDict = deptCollegeDict.set_index('系所')['院'].to_dict()
# print(deptCollegeDict)

# 讀入科系_學院_例外 配對excel   # 條列例外
deptCollegeDictExcept = pd.read_excel(dataPath + '0_科系_學院分類_例外.xlsx')
deptCollegeDictExcept = deptCollegeDictExcept.set_index('系所')['院'].to_dict()
# print(deptCollegeDictExcept)

# 讀入公司_次產業_銀行分類excel
subInd = pd.read_excel(dataPath +'0_金融業與次產業分類.xlsx', usecols = ['公司代碼簡稱','金融次產業','銀行分類'])
# subInd.head()

# 讀入TSE_學院配對excel  #注意沒有None、M9900 其他
industryCollegeDict = pd.read_excel(dataPath + '0_產業_學院配對_刪除其他None存託憑證.xlsx')
industryCollegeDict = industryCollegeDict.set_index('TSE新產業別')['學院'].to_dict()
# industryCollegeDict

# 執行日期
mmdd = date.today().strftime("%m%d")

# 只需要電子業&金融業 共9個產業
needIndustry = ['M2324 半導體','M2325 電腦及週邊','M2326 光電業','M2327 通信網路業',
                'M2328 電子零組件','M2329 電子通路業','M2330 資訊服務業','M2331 其他電子業',
                'M2800 金融業']