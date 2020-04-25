import pandas as pd
import numpy as np
import os
import re
from datetime import date
from step0_settings import domainPath, dataPath, mmdd, needIndustry




def mergeAll(series):
    major_all = []
    for majors in series.to_list():
        for major in majors.split('&'):
            if major not in major_all:
                major_all.append(major)
    return ('&'.join(major_all))

i = 0
for industry in needIndustry:
    i += 1
    print(f'第{i}個產業：{industry}')
    df = pd.read_excel(dataPath + f'1.0_學歷配對_{industry}_{mmdd}.xlsx')
    df['教育程度'].fillna('無', inplace = True)
    df = df[['董監經理人姓名','姓名代碼','教育程度','學經歷及目前兼任說明','學院']].drop_duplicates()
    df.reset_index(drop = True, inplace = True)

    final = df.groupby(['姓名代碼','董監經理人姓名','教育程度']).agg({'學經歷及目前兼任說明':lambda x: mergeAll(x),
                                                                    '學院': lambda x: mergeAll(x)}  
                                                                    ).reset_index()
    final.to_excel(dataPath + f'1.1_姓名學院_{industry}_{mmdd}.xlsx',
                        encoding = 'utf_8_sig', index = False
            )
print('------------------------------完成-----------------------------------------------')
