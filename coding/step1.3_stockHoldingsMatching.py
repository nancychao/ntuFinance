import pandas as pd
import numpy as np
import os
import re
from datetime import date
from step0_settings import domainPath, dataPath, mmdd, needIndustry

i = 0
for industry in needIndustry:
    i += 1
    print(f'第{i}個產業：{industry}')
    df = pd.read_excel(dataPath + f'1.2_學院更新_{industry}_0502.xlsx')
    dfStockHoldings = pd.read_excel(dataPath + '0.1_電子業內部人持股變化(年資料).xlsx')
    dfStockHoldings.rename(columns = {'持股人姓名':'董監經理人姓名'}, inplace = True)
    dfStockHoldings['公司代碼'] = dfStockHoldings['公司代碼'].apply(lambda x : str(x))
    df['公司代碼'] = df['公司代碼'].apply(lambda x : str(x))
    final = df.merge(dfStockHoldings[['公司代碼','董監經理人姓名','資料源年','期初股數','期末股數','期初股數(合計)','期末股數(合計)']], how='left', on=['公司代碼','資料源年','董監經理人姓名'])
    final.to_excel(dataPath + f'1.3_新增經理人持股數_{industry}_{mmdd}.xlsx',
                    encoding = 'utf_8_sig', index = False
        )
print('------------------------------完成-----------------------------------------------')
