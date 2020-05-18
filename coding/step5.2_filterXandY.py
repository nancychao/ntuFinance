import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, mmdd

# read excel
df = pd.read_excel(dataPath + f'5.1_電子業X+Y_{mmdd}.xlsx')
df['TSE'] = df['TSE新產業名'].apply(lambda x : x[:5])
df['總經理_姓名代碼'].replace('無',np.nan, inplace = True)
df['財務長_姓名代碼'].replace('無',np.nan, inplace = True)
df['獨董_姓名代碼'].replace('無',np.nan, inplace = True)


cols = [
 '公司代碼',
 '上市別',
 'TSE',
 '資料源年',
 '總經理_姓名代碼',
 '總經理_教育程度_代碼',
 '總經理_經理人年資max',
 '總經理_分數',
 '總經理_股數增減比率',
 '總經理_股數增減比率(合計)',
 '財務長_姓名代碼',
 '財務長_教育程度_代碼',
 '財務長_經理人年資max',
 '財務長_分數',
 '財務長_股數增減比率',
 '財務長_股數增減比率(合計)',
 '獨董_姓名代碼',
 '獨董_教育程度_代碼',
 '獨董_董監事年資max',
 '獨立董事_總分',
 '董監經理人變化_燈號',
 'CRISISCOR',
 '研究發展費用率'
        ]


df = df[cols]
df.to_excel(dataPath + f'5.2_電子業X+Y篩選_{mmdd}.xlsx',
            encoding = 'utf_8_sig', index = False
                                 )
print('-------------------------完成-----------------------------------')