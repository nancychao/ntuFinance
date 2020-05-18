import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, mmdd

# read excel
df = pd.read_excel(dataPath + f'5.2_電子業X+Y篩選_{mmdd}.xlsx')
train = df[df['董監經理人變化_燈號'] == 0]
test = df[df['董監經理人變化_燈號'] == 1]

train.to_excel(dataPath + f'5.3_電子業X+Y篩選train_{mmdd}.xlsx',
            encoding = 'utf_8_sig', index = False
                                 )

test.to_excel(dataPath + f'5.3_電子業X+Y篩選test_{mmdd}.xlsx',
            encoding = 'utf_8_sig', index = False
                                 )
                    
print('-------------------------完成-----------------------------------')