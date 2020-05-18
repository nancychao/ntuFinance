import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, mmdd

# Y1 = 危機公司(危機 = 1，其他 = 0)
# Y2 = 研究發展費用率(研發費用/營收)


# read excel
df = pd.read_excel(dataPath + f'5.0_電子業X_{mmdd}.xlsx')

# 危機公司資料
crisisDf = pd.read_excel(dataPath + '0_20200503-危機公司翊茹.xlsx')
crisisDf['公司代碼'] = crisisDf['COMCOR'].apply(lambda x : int(x[:4]))
crisisDf.rename(columns = {'YEAR' : '資料源年'}, inplace = True)
# print(crisisDf.shape)
# crisisDf

# df['資料源年'] = df['資料源年月'].apply(lambda x : x//100)

## 危機公司資料共57筆
# inner join 兩邊皆有的公司
# dfPlusCrisis = df.merge(crisisDf[['公司代碼','CRISISCOR','資料源年']], how='inner', on=['公司代碼','資料源年'])
# print(dfPlusCrisis.shape)


# left join
dfPlusCrisis = df.merge(crisisDf[['公司代碼','CRISISCOR','資料源年']], how='left', on=['公司代碼','資料源年'])
dfPlusCrisis['CRISISCOR'].fillna(0,inplace = True)
# print(dfPlusCrisis.shape)
# dfPlusCrisis


# 研究發展費用率
devDf = pd.read_excel(dataPath+'0_全產業研究發展費用率.xlsx')
devDf['公司代碼'] = devDf['公司'].apply(lambda x : int(x[:4]))
devDf['資料源年'] = devDf['年月'].apply(lambda x : int(x.strftime('%Y')))
# devDf

# left join 兩邊皆有的公司
dfPlusDev = dfPlusCrisis.merge(devDf[['公司代碼','研究發展費用率','資料源年']], how='left', on=['公司代碼','資料源年'])
# print(dfPlusDev.shape)
# dfPlusDev

dfPlusDev.to_excel(dataPath + f'5.1_電子業X+Y_{mmdd}.xlsx',
            encoding = 'utf_8_sig', index = False
                                 )
print('-------------------------完成-----------------------------------')


