import pandas as pd
import numpy as np
import os
import re
from datetime import date
from step0_settings import domainPath, dataPath, deptCollegeDict, deptCollegeDictExcept, subInd, industryCollegeDict, mmdd, needIndustry

pd.options.display.max_columns = 100
pd.options.display.max_rows = 2000  #調整pandas輸出display列數
pd.set_option('display.max_colwidth',1000)  #設定每個欄位dsiplay字數長度


def FindMajor(string):  #清除"大學"或其他學校名稱以前的字 
    global keyword
    index = keyword.index('藥專')
    for school in keyword[: (index+1)]:
        if school not in string:
            continue
        return string.split(school)[-1]
    return string

def MatchMajorCollege(string):   #配對科系&相對應學院
    global deptCollegeDict
    for keys in deptCollegeDictExcept:
        if keys in string:
            return (deptCollegeDictExcept[str(keys)])
        
    major_name = FindMajor(string)
    for keys in deptCollegeDict:
        if keys in major_name:
            return(deptCollegeDict[str(keys)])
    return("無法辨識")



keyword = ['大學',                 #會先使用index較前面字串去比對，所以越常出現的index會越前面
           '學校',
           '科大','University','Univ','應大',
           '補校',
           '中學','初中','中學',
           '女中','高中',
           '學院',
           
           '商專','商職','商工','高商','家商','工職','高工','工商','工農','家職','高職',
           '專科','專校',
           '家專','企專','空專','海專','醫專','工專','農專','藥專',
           
           
           '研究所','博士','學士','碩士','所','系','Ph\.D',
           'MBA','College','School','u\.'
           ] 

#可以辨別是否是學歷的關鍵字，有出現這些字則高度可能是經歷
delete = ['教授','主任','講師','事務所','所長','院長','校長','董事','執行長','委員',
          '處長','協理','總務長','館長','副總裁',
           '總經理','邀請','主管','株式會社','兼任','曾任','理事長','研究員','副理',
          '股份有限公司','證券交易所','所羅門(股)公司財務處專案經理','工研院電子所課長',
          '期貨交易所','台科視訊系統(蘇州)有限公司經理','工研院電子所會計室管理師','艾睿電子經理',
          '中國鋼鐵\(股\)公司成本帳務系統組','專任助教','資深經理','本行系統開發營運部經理',
          '司法官訓練所',
          '公司監察人','大學學務長','期交所監察人','中華民國工商協進會顧問','中央研究所',
          '教務處註冊組','學生事務長','工研院','工業技術研究院','大學任教','司法官訓練所第',
          '交通部電信研究所','大學助教','公司監察人','大學教務長',
          '\(股\)公司','中油煉製研究所專案經理','校友會會長','臺灣工商聯合會會務顧問','校友總會',
          '中科院航空研究所',
          '高職教師','鳳和中學教師','中心診所醫療財團法人中心綜合醫院顧問','大學總務處文書組長',
          '大學產學合作處產學合作長','德州儀器公司','中山科學院副組長','工商建研會理事',
          '\[無職稱\]','醫藥專業行銷人員認證','研討會主講人','計畫總主持人','工商時報專欄記者',
          '醫院婦產科醫師','中泰電腦系統部經理','傑出校友','Professor',
          'Chair of Council and Senior Pro-Chancellor',
          '英國Loughborough大學Council and Senior Pro-Chancellor主席','訪問學者','訪問學人'
         ]




class CleanData():
#     global datapath,industryCollegeDict,industry, industryProCollege, industryData
    
    @staticmethod
    def readIndsData(industry): # 讀入特定產業的檔案
        df = pd.read_excel(dataPath + '\\TEJ輸出檔案\\0318版\\依產業分\\' + industry + '.xlsx')
        return df

    @staticmethod
    def dropDupAndMatchCollege(df):
        # 行次： A:主要學經歷說明 B:目前兼任說明 (刪除 目前兼任說明)
        df = df[df['行次'].str.contains("A")]

        # 刪除重複值 (因為有可能有同年、同公司、同姓名代碼 的 學經歷)
        df = df.drop_duplicates(subset = ['公司代碼','資料源年月','姓名代碼','學經歷及目前兼任說明'])


        # 篩選"學歷"：學經歷中符合篩選標準的(含有"大學"等字詞，且不包含"教授"等字詞)
        df = df[(df['學經歷及目前兼任說明'].str.contains("|".join(keyword),flags=re.IGNORECASE)) &
            (~df['學經歷及目前兼任說明'].str.contains("|".join(delete),flags=re.IGNORECASE, na = False))]  #刪除['學經歷及目前兼任說明']欄位為空值之資料

        # 科系_學院 配對
        df['學院'] = df['學經歷及目前兼任說明'].map(MatchMajorCollege)
        
        return df
    
    @staticmethod    
    def stripColumns(df):
        
#         def formatCmpId(id):
#             if type(id) == int:
#                 return '{:04d}'.format(id)
#             else:
#                 return id.strip(' ')

        df['公司代碼'] = df['公司代碼'].apply(lambda x :str(x).strip())
        df['簡稱'] = df['簡稱'].apply(lambda x :x.strip())
        df['上市別'] = df['上市別'].apply(lambda x :x.strip())
        df['董監經理人姓名'] = df['董監經理人姓名'].apply(lambda x :x.strip())
        df['資料源'] = df['資料源'].apply(lambda x :str(x).strip())
        df['職稱'] = df['職稱'].apply(lambda x :x.strip())
        df['身份別'] = df['身份別'].apply(lambda x :x.strip())
        df['教育程度'] = df['教育程度'].apply(lambda x :x.strip())
        df['學經歷及目前兼任說明'] = df['學經歷及目前兼任說明'].apply(lambda x :x.strip())
        
        return df

    @staticmethod
    def joinIntoOneRow(df):
        cols = list(df)
        cols.remove('學經歷及目前兼任說明')
        cols.remove('學院')
        cols.remove('行次')       
        
        # 將多筆學歷整合在同一列中
        df = df.groupby(cols).agg({'學經歷及目前兼任說明':lambda x: '&'.join(x),
                                        '學院':lambda x: '&'.join(x)}).reset_index()
        return df

    @staticmethod
    def crtOrMergeNewColumns(df):
        # 合併"公司代碼" & "簡稱"
        df.insert(1,'公司代碼簡稱', df['公司代碼'].map(str) + df['簡稱'].apply(lambda x: x.strip()))
        df.drop(columns = ['簡稱'], inplace = True)
        
        # 插入"資料源年"欄位
        df.insert(5,'資料源年', df['資料源年月'].apply(lambda x : x//100) )  #每年可能有多筆資料，不一定只有12月一筆資料

        # 依照"資料源年"排序
        df.sort_values(by=["資料源年月", "公司代碼"], inplace = True)

        # 配對公司_次產業 & 銀行分類
        df = df.merge(subInd, how='left', on=['公司代碼簡稱'])
        
        # 新增「職位代碼_排序」欄位
        conditions = [
            (df['職位代碼'] == 0),  # 總裁
            (df['職位代碼'] == 1) | (df['職位代碼'] == 4),  # 總經理、執行長
            (df['職位代碼'] == 7),  #副總裁
            (df['職位代碼'] == 3) | (df['職位代碼'] == 5), #部門總經理、副執行長
            (df['職位代碼'] == 2), #副總經理
            (df['職位代碼'] == 9), #協理
            (df['職位代碼'] == 8), #經理
            (df['職位代碼'] == 'Z'), #副理
            (df['職位代碼'] == 'P'), #其他職務
            ]

        choices = [1, 2, 3, 4, 5,6,7,8,9]
        df['職位代碼_排序'] = np.select(conditions, choices, default=10)
        
        return df
    
    @staticmethod
    def sortColumns(df):
        cols = ['公司代碼', '公司代碼簡稱', '上市別', 'TSE新產業名', '金融次產業', '銀行分類',
                '資料源年月', '資料源年', '董監經理人姓名', '姓名代碼', '資料源', '資料截止日',
                '職稱', '身份別', '初任日期--董監事', '初任日期--經理人', '年資--董監事',
                '年資--經理人', '教育程度', '財務', '會計', '法務', '身份代碼', '獨立代碼',
                '職位代碼', '職位代碼_排序', '會計主管', '財務主管', '類別代碼', '重整代碼',
                '初次選任日期--董監事', '選就任日', '沿用日', '學經歷及目前兼任說明', '學院'
               ]
        # 重新排列dataframe
        df = df.loc[:, cols]

        return df


i = 0
for industry in needIndustry: 
    i += 1
    print('第{i}個產業：{industry}'.format(i=i,industry=industry))
    industryData = CleanData.readIndsData(industry)    # 讀取產業TEJ輸出raw data
    df = CleanData.dropDupAndMatchCollege(industryData) # 刪除重複值&配對學院
    df = CleanData.stripColumns(df) # 刪除各欄位前後的空白
    df = CleanData.joinIntoOneRow(df) # 合併同年度、同公司、同姓名代碼成一列
    df = CleanData.crtOrMergeNewColumns(df) #新增欄位或merge
    df = CleanData.sortColumns(df) # 將columns依想要的依序排列
    df.to_excel(dataPath + f'1_學歷配對_{industry}_{mmdd}.xlsx',
                            encoding = 'utf_8_sig', index = False
                )
# industryData
print('------------------------------完成-----------------------------------------------')



