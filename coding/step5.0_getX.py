import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, mmdd

# 計算年資max
def getMaxNum(string):
    string = str(string)
    if '&' not in string:
        return(float(string))
    else:
        if '-' in string:
            list = [np.nan]
            for i in string.split('&'):
                if i == '     -':
                    continue
                list.append(float(i))
            return max(list)
        
        else:
            list = [float(i) for i in string.split('&')]
            return max(list)

# 計算CEO、CFO、獨立董事與前值是否有變化
def change(df):  
    df['總經理_姓名代碼'].fillna('無', inplace = True)
    df['財務長_姓名代碼'].fillna('無', inplace = True)
    df['獨董_姓名代碼'].fillna('無', inplace = True)
    df['總經理_changed'] = df['總經理_姓名代碼'].ne(df['總經理_姓名代碼'].shift().bfill()).astype(int)
    df['財務長_changed'] = df['財務長_姓名代碼'].ne(df['財務長_姓名代碼'].shift().bfill()).astype(int)
    df['獨董_changed'] = df['獨董_姓名代碼'].ne(df['獨董_姓名代碼'].shift().bfill()).astype(int)
    return df


# read excel
df = pd.read_excel(dataPath + f'4.1_上市上櫃_經理人+獨董_{mmdd}.xlsx')

# 分類 "電子業"&"金融業"
df.insert(4, '產業',np.where(df['TSE新產業名']=='M2800 金融業','金融業','電子業'))
df = df[df['產業'] == '電子業']  # 僅選擇電子業data

df = df.sort_values(['公司代碼','資料源年月'])  # 排序
df['總經理&財務長_總分'] = df['總經理_分數'] + df['財務長_分數']
df['總經理&財務長&獨董_總分'] = df['總經理&財務長_總分'] + df['獨立董事_總分']
# df.head()


# 建立教育程度代碼，並計算年資max
positions = ['總經理','財務長','獨董']
for position in positions:
    print(position)
    df[position+'_教育程度'].fillna('缺', inplace = True)
    df[position+'_教育程度_代碼'] = np.where(df[position+'_教育程度'].str.contains('博士'), 1,
                                           np.where(df[position+'_教育程度'].str.contains('碩士'), 2,
                                               np.where(df[position+'_教育程度'].str.contains('大學'), 3,
                                                   np.where(df[position+'_教育程度'].str.contains('高中職以下'), 4,
                                                       np.where(df[position+'_教育程度'].str.contains('無'), 5, 6)))))
df['總經理_經理人年資max'] = df['總經理_經理人年資'].apply(lambda x : getMaxNum(x))
df['財務長_經理人年資max'] = df['財務長_經理人年資'].apply(lambda x : getMaxNum(x))
df['獨董_董監事年資max'] = df['獨董_董監事年資'].apply(lambda x : getMaxNum(x))


# 僅選擇12月份的資料，並計算與前值是否有差異，(1 = 有差異，0 = 無差異)
df = df[df['資料源年月'].apply(lambda x : x%100) == 12]
df = df.groupby(['公司代碼']).apply(func = change).reset_index(drop=True)

# 計算差異總分
df['changed_sum'] = df['總經理_changed'] +df['財務長_changed'] + df['獨董_changed']
# 計算董監經理人變化燈號
df['董監經理人變化_燈號'] = np.where(df['changed_sum'] == 0, 0, 1)


# CEO、CFO持股變化
positions = ['總經理','財務長']
for position in positions:
    df[f'{position}_股數增減比率'] = (df[f'{position}_期末股數'] - df[f'{position}_期初股數']) / df[f'{position}_期初股數']
    df[f'{position}_股數增減比率(合計)'] = (df[f'{position}_期末股數(合計)'] - df[f'{position}_期初股數(合計)']) / df[f'{position}_期初股數(合計)']
    df[f'{position}_股數增減比率'].fillna(0, inplace = True)
    df[f'{position}_股數增減比率(合計)'].fillna(0, inplace = True)

df.to_excel(dataPath + f'5.0_電子業X_{mmdd}.xlsx',
            encoding = 'utf_8_sig', index = False
                                 )
print('-------------------------完成-----------------------------------')



