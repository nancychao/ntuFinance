import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath 

stockRatioDf = pd.read_excel(dataPath +'電子業內部人持股變化(月資料).xlsx')
stockRatioDf.rename(columns = {'公司代碼' : '公司代碼簡稱'}, inplace = True)
stockRatioDf['公司代碼'] = stockRatioDf['公司代碼簡稱'].apply(lambda x : int(x[:4]))
stockRatioDf['資料源年'] = stockRatioDf['年月日'].apply(lambda x : int(x.strftime('%Y')))
# stockRatioDf

stockRatioDf['身份別'].fillna('無', inplace = True)
stockRatioDf.sort_values(by = ['年月日'], inplace = True)
stockAgg = stockRatioDf.groupby(['公司代碼簡稱', '持股人姓名','公司代碼', '資料源年']).agg({'期初股數': lambda x : x.iloc[0],
                                                                               '期末股數': lambda x : x.iloc[-1],
                                                                               '期初股數(合計)':lambda x:x.iloc[0],
                                                                               '期末股數(合計)': lambda x : x.iloc[-1],
                                                                               '年月日': lambda x :x.iloc[-1],
                                                                               '身份別': lambda x : '&'.join(x)
    
}).reset_index()
# stockAgg

stockAgg.to_excel(dataPath + '2.0_電子業內部人持股變化(年資料).xlsx',
                        encoding = 'utf_8_sig', index = False
            )
print('------------------------------完成-----------------------------------------------')
