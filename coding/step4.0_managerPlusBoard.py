import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, mmdd, needIndustry 


pd.options.display.max_columns = 100
pd.options.display.max_rows = 2300  #調整pandas輸出display列數
pd.set_option('display.max_colwidth',1000)  #設定每個欄位dsiplay字數長度


i = 0
df = pd.DataFrame()
# df.columns = ['公司代碼', '公司代碼簡稱', '上市別', 'TSE新產業名', '金融次產業', '銀行分類', '資料源年月', '資料源年']
for industry in needIndustry:
#     industry = 'M2800 金融業'
    i += 1
    print('第'+str(i)+'個產業：'+industry)
    manager = pd.read_excel(dataPath + f'2_高階經理人_{industry}_{mmdd}.xlsx') 
    board = pd.read_excel(dataPath + f'3_獨立董事_{industry}_0502.xlsx') 
    mix = manager.merge(board, how = 'outer', on = list(manager)[:8])
    print('第'+str(i)+'個產業：'+industry+'  '+str(mix.shape))
    df = df.append(mix)
    
df['經理人獨董_總分'] = df['高階經理人_總分'] + df['獨立董事_總分']
managerCols = list(manager.columns)
boardCols = list(board.columns)[8:]
managerCols.extend(boardCols)
managerCols.append('經理人獨董_總分')
df = df.loc[:, managerCols]
    
df.to_excel(dataPath + '4.0_經理人+獨董_'+mmdd+'.xlsx',
                                    encoding = 'utf_8_sig', index = False
                                 )
print('-------------------------完成-----------------------------------')