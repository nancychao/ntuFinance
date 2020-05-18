import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, industryCollegeDict, mmdd, needIndustry 

pd.options.display.max_columns = 100
pd.options.display.max_rows = 2300  #調整pandas輸出display列數
pd.set_option('display.max_colwidth',1000)  #設定每個欄位dsiplay字數長度


managerTable = pd.DataFrame()

class ManagerFunc():
#     global industryCollegeDict
#     global industryData,managerDict,position
    global position
    
    @staticmethod
    def openIndsData(industry):  # 讀取學歷配對資料
        df = pd.read_excel(dataPath + f'1.3_新增經理人持股數_{industry}_{mmdd}.xlsx')           
        return df
    


    @staticmethod
    def getManagerDict():  # 各產業高階經理人的關鍵字、專業學院
        global industryProCollege
        managerDict = {'總經理':
                           {'包含':'總經理|執行長|總裁',
                            '不包含':'副總經理|總經理室|總經理辦公室|執行長室|總裁室|事業群執行長|處執行長|事業執行長|執行長特助|總經理特助|分公司總經理|中心總經理|營運總經理|功能性總經理|處總經理|區總經理|部總經理|群總經理|部代總經理|部門總經理|廠總經理|體總經理|單位總經理|業務總經理|事業總經理|處1總經理|處I總經理|處II總經理|行政總經理|文創總經理|駐地總經理|營運總經理|嘉鎂總經理|電子總經理|廣禾總經理|新力總經理|重慶總經理|深圳總經理|和進總經理|和昌總經理|資訊總經理|GWC總經理|中山上詮總經理|台北總經理|昆山總經理|HPIW總經理|子公司長瑞總經理|子公司蘇州長瑞光電總經理|上海合晶總經理|SBEH總經理|客戶總經理|SAC總經理|上海鈺太芯總經理|享慶科技總經理|關係企總經理',
                            '專業學院':'管理學院|{industryProCollege}'.format(industryProCollege=industryProCollege)
                            },
                       '財務長':
                           {'身份別':'財務主管|財會主管',
                            '包含':'財務長',
                            '不包含':'副財務長|財務長室|財務長辦公室',
                            '專業學院':'管理學院'
                            },
                       '營運長':
                           {'職稱':'營運',
                            '包含':'營運長',
                            '不包含':'副營運長|營運長室|營運長辦公室|中心營運長', 
                            '專業學院':'管理學院|{industryProCollege}'.format(industryProCollege=industryProCollege)
                            },
                       '法遵長':
                           {'職稱':'法遵|法令遵循',
                            '包含':'法遵長|法遵處處長|法令遵循處處長|法遵/法務處處長|法遵暨法務處處長|法令遵循主管|法令遵循處負責人|法令遵循處代處長',
                           '不包含':'副法遵長|法遵長室|法遵長辦公室|秘書處',
                           '專業學院':'法學院|{industryProCollege}'.format(industryProCollege=industryProCollege)
                           },
                       '法務長':
                           {'職稱':'法務|法律事務',
                            '包含':'法務長|法務處處長|法務暨法令遵循處處長',
                            '不包含':'副法務長|法務室|秘書處',
                            '專業學院':'法學院'
                            },
                       '資訊長':
                           {'職稱':'資訊',
                            '包含':'資訊長|資訊處處長|資訊科技處處長|資訊執行長',
                            '不包含':'副資訊長|資訊長室|資訊長辦公室',
                            '專業學院':'電機資訊學院|工學院'
                            },
                       '技術長':
                           {'職稱':'技術',
                            '包含':'技術長',
                            '不包含':'副技術長|技術長室|技術長辦公室',
                            '專業學院':'電機資訊學院|工學院'
                            },
                       '資安長':
                            {'職稱':'資安|資訊安全',
                             '包含':'資安長|資訊安全處處長|資安處處長',
                             '不包含':'副資安長|資安長室|資安長辦公室',
                             '專業學院':'電機資訊學院'
                             }

                  }
        return managerDict
    
    @staticmethod
    def crtManagerTable(df): # 建table儲存資料
        df = df[list(df)[:8]].drop_duplicates()
        df.reset_index(drop = True, inplace = True)
        return df
    
    @staticmethod
    def getMgsTopFilter(df):  # 選擇高階經理人篩選器
        global managerDict


        if len(df.index) == 1:  #如果該職稱只有一個人，則return，例如：法遵/法令遵循
            return df
        
        else: # 如果該職稱有很多人

            df = df[df['職位代碼_排序'] == df['職位代碼_排序'].min()]  # 選擇職位排序較前面者(較高階)

            if len(df.index) == 1:   #如果同位階只有一筆
                return df

            else:   #如果同位階有多筆，檢查職稱中有沒有特別的高階經理人關鍵字，例如：法遵長
                positionContains = managerDict[position]['包含']
                positionContainsDel = managerDict[position]['不包含']

                df2 = df[(df['職稱'].str.contains(positionContains)) 
                             & (~df['職稱'].str.contains(positionContainsDel, na = True))]

                if len(df2.index) > 0:  #如果有人的職稱中包含關鍵字
                    return df2

                else:    #如果都沒有人的職稱包含關鍵字：就return全部
                    return df

    # @staticmethod
    # def getCeoTopFilter(df): # 選擇CEO篩選器
        
    #     if len(df.index) == 1:  #如果該職稱只有一個人，則return，例如：總裁/法遵長
    #         return df
        
    #     else: # 如果該職稱有很多人
    #         df = df[df['職位代碼_排序'] == df['職位代碼_排序'].min()]  # 選擇職位排序較前面者(較高階)
    #         return df


    @staticmethod        
    def getMgsTopMain(df):
        global managerDict

        # 依照職位不同進行初步篩選
        if position == '總經理':
            df_position = df[(df['職位代碼_排序'] == 1)|(df['職位代碼_排序'] == 2)|(df['職位代碼_排序'] == 3) ]   # 此職位以總裁(1)為優先  # 如果沒有總裁職位者，再選擇總經理、執行長(2)
            df_final = df_position.groupby(['資料源年月','公司代碼']).apply(func = ManagerFunc.getMgsTopFilter).reset_index(drop=True)
            
            return df_final.sort_values(["資料源年月", "公司代碼"])


        elif position == '財務長':
            title = managerDict['財務長']['身份別']
            df_position = df[df['身份別'].str.contains(title, na=False)]
            
            
        else:
            df_position = df[df['職稱'].str.contains(managerDict[position]['職稱'], na=False)]

            
        # 選擇各年度各公司最高長官
        if len(df_position.index) == 0:
            return df_position
            
        else:    
            df_final = df_position.groupby(['資料源年月','公司代碼']).apply(func = ManagerFunc.getMgsTopFilter).reset_index(drop=True)

            return df_final.sort_values(["資料源年月", "公司代碼"])   
    
    @staticmethod
    def countAndMergeNumPeople(df):
        global managerTable
  
        if len(df.index) == 0:
            managerTable[position+'_總人數'] = 0
            return managerTable
        
        else:
            dfPeople = df.groupby(['上市別','公司代碼簡稱','資料源年月']).size().reset_index(name= position+'_總人數')
            managerTable = managerTable.merge(dfPeople, how='left', on=['上市別','公司代碼簡稱','資料源年月'])
    #         managerTable.update(managerTable[list(managerTable)[8:]].fillna(0)) 
            return managerTable

    @staticmethod
    def getAndMergeInfo(df):
        global managerTable
        
        if len(df.index) == 0:
            managerTable[position+'_姓名代碼'] = np.nan
            managerTable[position+'_姓名'] = np.nan
            managerTable[position+'_職位代碼_排序'] = np.nan
            managerTable[position+'_教育程度'] = np.nan
            managerTable[position+'_經理人年資'] = np.nan
            managerTable[position+'_職稱'] = np.nan
            managerTable[position+'_身份別'] = np.nan
            managerTable[position+'_學經歷及目前兼任說明'] = np.nan
            managerTable[position+'_學院'] = np.nan
            managerTable[position+'_期初股數'] = np.nan
            managerTable[position+'_期末股數'] = np.nan
            managerTable[position+'_期初股數(合計)'] = np.nan
            managerTable[position+'_期末股數(合計)'] = np.nan
            
            return managerTable        
        
        
        else:
            df['教育程度'].fillna('無揭露', inplace=True)
            df['年資--經理人'] = df['年資--經理人'].apply(lambda x : str(x))
            managerNameId = df.groupby(['上市別','公司代碼簡稱','資料源年月']).agg({'姓名代碼':lambda x: '&'.join(x),
                                                                                  '董監經理人姓名':lambda x: '&'.join(x),
                                                                                  '職位代碼_排序': lambda x : min(x),
                                                                                  '教育程度': lambda x : '&'.join(x),
                                                                                  '年資--經理人': lambda x :'&'.join(x),
                                                                                  '職稱' : lambda x : '&'.join(x),
                                                                                  '身份別': lambda x : '&'.join(x),
                                                                                  '學經歷及目前兼任說明' : lambda x : '#'.join(x),
                                                                                  '學院': lambda x : '#'.join(x),
                                                                                  '期初股數': lambda x : sum(x),
                                                                                  '期末股數': lambda x : sum(x),
                                                                                  '期初股數(合計)': lambda x : sum(x),
                                                                                  '期末股數(合計)': lambda x : sum(x),
                                                                                  
                                                                                }).reset_index()
            managerNameId.rename(columns = {'姓名代碼' : position+'_姓名代碼',
                                            '董監經理人姓名':position+'_姓名',
                                            '職位代碼_排序' : position+'_職位代碼_排序',
                                            '教育程度' : position+'_教育程度',
                                            '年資--經理人' : position+'_經理人年資',
                                            '職稱' : position+'_職稱',
                                            '身份別' : position+'_身份別',
                                            '學經歷及目前兼任說明' : position+'_學經歷及目前兼任說明',
                                            '學院' : position+'_學院',
                                            '期初股數':position+'_期初股數',
                                            '期末股數':position+'_期末股數',
                                            '期初股數(合計)':position+'_期初股數(合計)',
                                            '期末股數(合計)':position+'_期末股數(合計)',

                                           }, inplace = True)
            managerTable = managerTable.merge(managerNameId, how='left', on=['上市別','公司代碼簡稱','資料源年月'])
            return managerTable
    
    @staticmethod
    def countAndMergeNumPro(df):
        global managerTable

        if len(df.index) == 0:  #如果此類高階經理人，無資料
            managerTable[position+'_符合人數'] = 0
            managerTable[position+'_不符合人數'] = 0
            managerTable[position+'_無法辨識學院人數'] = 0
            
            return managerTable       
        
        else:
            positionNeedCollege = managerDict[position]['專業學院']
            df['專業度'] = np.where(df['學院'] == '無法辨識', '無法辨識學院',
                                np.where(df['學院'].str.contains(positionNeedCollege), '是', '否'))
            # if == '無法辨識'
            #    return '無法辨識'
            # else:
            #    if 符合專業標準:
            #         return '是'
            #    else:
            #         return '否'
            
            # 若學院是無法辨識，且教育程度：高中職以下 → 不符合，因為覺得高中或以下學歷並沒有專業科目課程
            df.loc[(df['學院'] == '無法辨識')&(df['教育程度'] == '高中職以下'),'專業度' ] = '否'


            managerPro2 = df.groupby(['上市別','公司代碼簡稱','資料源年月','專業度']).size().reset_index(name='人數')
            managerPro2 = managerPro2.pivot_table(index=['上市別','公司代碼簡稱','資料源年月'], columns='專業度', values='人數').reset_index()
            tmpDf = managerPro2.copy()[['上市別','公司代碼簡稱','資料源年月']]
            tmpDf['是'] = np.nan
            tmpDf['否'] = np.nan
            tmpDf['無法辨識學院'] = np.nan
            for status in ['是','否','無法辨識學院']:
                if status in list(managerPro2):
                    tmpDf = tmpDf.merge(managerPro2[['上市別', '公司代碼簡稱', '資料源年月', status]], how='left', on=['上市別','公司代碼簡稱','資料源年月'])
                    tmpDf[status] = tmpDf[status + '_x'].fillna(tmpDf[status+'_y'])
                    tmpDf.drop([status+'_x', status+'_y'], 1, inplace=True)

            tmpDf.rename(columns = {'是' : position+'_符合人數',
                                     '否' : position+'_不符合人數',
                                     '無法辨識學院' : position+'_無法辨識學院人數'
                                    }, inplace = True)
            tmpDf.fillna(0, inplace = True)
    #         tmpDf

            managerTable = managerTable.merge(tmpDf, how='left', on=['上市別','公司代碼簡稱','資料源年月'])
    #         managerTable.update(managerTable[list(managerTable)[8:]].fillna(0)) 
            return managerTable
    
    @staticmethod
    def getLabelAndScore():
        global managerTable
        managerTable[position+'_燈號'] = np.where(managerTable[position+'_符合人數'] > 0, '符合',
                                           np.where(managerTable[position+'_不符合人數'] > 0, '不符合',
                                               np.where(managerTable[position+'_無法辨識學院人數'] > 0, '無法辨識學院',
                                                   '沒有揭露')))
        managerTable[position+'_分數'] = np.where(managerTable[position+'_燈號'] =='符合', 1, 0)
        return managerTable
    
    @staticmethod
    def getAllMgsScore():
        global managerTable
        managerTable['高階經理人_總分'] = managerTable['總經理_分數']+managerTable['財務長_分數']+managerTable['營運長_分數']+managerTable['法遵長_分數']+managerTable['法務長_分數']+managerTable['資訊長_分數']+managerTable['技術長_分數']+managerTable['資安長_分數']
        return managerTable




# from datetime import date
# mmdd = date.today().strftime("%m%d")
# mmdd


i = 0
# needIndustry = ['M2324 半導體','M2325 電腦及週邊','M2326 光電業','M2327 通信網路業',
#                 'M2328 電子零組件','M2329 電子通路業','M2330 資訊服務業','M2331 其他電子業',
#                 'M2800 金融業']

for industry in needIndustry:
    i += 1
    print(f'第{i}個產業：{industry}')
    industryProCollege = industryCollegeDict[industry] # 各產業專業
    managerDict = ManagerFunc.getManagerDict()   # 各高階經理人職位條件
    
    industryData = ManagerFunc.openIndsData(industry)  # 讀檔
    managerTable = ManagerFunc.crtManagerTable(industryData) # 建立總表
    
    for position in managerDict:
        print(position)
        df = ManagerFunc.getMgsTopMain(industryData)  # 找各職位的高階經理人
        managerTable = ManagerFunc.countAndMergeNumPeople(df) # 計算該職位人數
        managerTable = ManagerFunc.getAndMergeInfo(df)  # merge該經理人其他資訊
        managerTable = ManagerFunc.countAndMergeNumPro(df) # 計算符合、不符合、無法辨識學院人數
        managerTable = ManagerFunc.getLabelAndScore() # 取得該職位燈號&分數
    managerTable = ManagerFunc.getAllMgsScore() # 計算所有高階經理人分數


    managerTable.to_excel(dataPath + f'2_高階經理人_{industry}_{mmdd}.xlsx',
                                encoding = 'utf_8_sig', index = False
                             )
#     managerTable
