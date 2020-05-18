import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, mmdd 


df = pd.read_excel(dataPath + f'4.0_經理人+獨董_{mmdd}.xlsx') 
df = df[(df['上市別'] == 'TSE')|(df['上市別'] == 'OTC')]
df.to_excel(dataPath + f'4.1_上市上櫃_經理人+獨董_{mmdd}.xlsx',
                                    encoding = 'utf_8_sig', index = False
                                 )
print('-------------------------完成-----------------------------------')