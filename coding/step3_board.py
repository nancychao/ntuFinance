import pandas as pd
import numpy as np
import os
from step0_settings import domainPath, dataPath, industryCollegeDict, mmdd, needIndustry 

pd.options.display.max_columns = 100
pd.options.display.max_rows = 3000  #調整pandas輸出display列數
pd.set_option('display.max_colwidth',1000)  #設定每個欄位dsiplay字數長度


boardTable = pd.DataFrame()

class BoardFunc():
    @staticmethod
    def openIndsData(industry):
        df = pd.read_excel(dataPath + f'1.2_學院更新_{industry}_{mmdd}.xlsx')        
        return df

    @staticmethod
    def getProDict():
        global industry, proDict
        
        proDict = {'財務':'管理學院',
                   '會計':'管理學院',
                   '法務':'法學院',
                   'IT':'電機資訊學院',
                   '產業專業':industryCollegeDict[industry]
                  }

        return proDict

    @staticmethod
    def crtBoardTable(df):  #建立空table存取所有公司各年月的資料
        df = df[list(df)[:8]].drop_duplicates()
        df.reset_index(drop = True, inplace = True)
        return df

    @staticmethod
    def getIndependentDirector(df): # 帶入TEJ中，職稱有包含"獨立董事"的資料
        df = df[(df['職稱'].str.contains('獨立董事'))|(df['身份別'].str.contains('獨立董事')) ].reset_index(drop = True)
        return df 
    
    @staticmethod
    def countBoardNumPeople(df): # 計算每家公司各年月的獨董人數
        global boardTable
        boardPeople = df.groupby(['上市別','公司代碼簡稱','資料源年月']).size().reset_index(name='獨立董事_總人數')
        boardTable = boardTable.merge(boardPeople, how='left', on=['上市別','公司代碼簡稱','資料源年月'])
        
        return boardTable
    
    @staticmethod
    def getAndMergeInfo(df):
        global boardTable
        
        df['教育程度'].fillna('無揭露', inplace=True)
        df['身份別'].fillna('-', inplace=True)
        df['年資--董監事'] = df['年資--董監事'].apply(lambda x : str(x))
        directorInfo = df.groupby(['上市別','公司代碼簡稱','資料源年月']).agg({'姓名代碼':lambda x: '&'.join(x),                                                        
                                                                              '董監經理人姓名':lambda x: '&'.join(x),
                                                                              '教育程度': lambda x : '&'.join(x),
                                                                              '年資--董監事': lambda x :'&'.join(x),
                                                                              '職稱' : lambda x : '&'.join(x),
                                                                              '身份別': lambda x : '&'.join(x),
                                                                              '學經歷及目前兼任說明' : lambda x : '#'.join(x),
                                                                              '學院': lambda x : '#'.join(x)
                                                                            }).reset_index()
        directorInfo.rename(columns = {'姓名代碼' : '獨董_姓名代碼',
                                       '董監經理人姓名':'獨董_姓名',
                                        '教育程度' : '獨董_教育程度',
                                        '年資--董監事' : '獨董_董監事年資',
                                        '職稱' : '獨董_職稱',
                                        '身份別' : '獨董_身份別',
                                        '學經歷及目前兼任說明' : '獨董_學經歷及目前兼任說明',
                                        '學院' : '獨董_學院'

                                       }, inplace = True)
        boardTable = boardTable.merge(directorInfo, how='left', on=['上市別','公司代碼簡稱','資料源年月'])
        return boardTable        
        
        
    
    @staticmethod
    def countAndMergeSpecialty(df):
        global boardTable, industry,industryCollegeDict,proDict
        

        for pro in proDict: # 依序整理所有獨董評估項目
            
            # 帶入TEJ整理的三個獨立董事欄位(財務、會計、法務)
            if pro == '財務':
                df['專業度'] = np.where((df['財務'].str.contains('V'))|(df['財務主管'].str.contains('Y')), '是', '否')  # np.where沒辦法用np.nan
                df['專業度'].replace('否', np.nan, inplace = True)

            elif pro == '會計':
                df['專業度'] = np.where(df['會計'].str.contains('V')|(df['會計主管'].str.contains('Y')), '是', '否')  # np.where沒辦法用np.nan
                df['專業度'].replace('否', np.nan, inplace = True)

            elif pro == '法務':
                df['專業度'] = np.where(df['法務'].str.contains('V'), '是', '否')  # np.where沒辦法用np.nan
                df['專業度'].replace('否', np.nan, inplace = True)

            else:
                df['專業度'] = np.nan
            

            # 建立新table儲存學院專業結果
            df2 = df.copy()
            needCollege = proDict[pro]
            df2['專業度'] = np.nan
            df2['專業度'] = np.where(df2['學院'] == '無法辨識', '無法辨識學院',
                                np.where(df2['學院'].str.contains(needCollege), '是', '否'))
            # if == '無法辨識'
            #    return '無法辨識'
            # else:
            #    if 符合專業標準:
            #         return '是'
            #    else:
            #         return '否'

            # 若學院是無法辨識，且教育程度：高中職以下 → 不符合，因為覺得高中或以下學歷並沒有專業科目課程
            df2.loc[(df2['學院'] == '無法辨識')&(df2['教育程度'] == '高中職以下'),'專業度' ] = '否'


            
            # 整合TEJ整理結果&學院專業結果
            df = df.merge(df2, how='left', on=list(df)[:-1])
            df['專業度'] = df['專業度_x'].fillna(df['專業度_y'])
            df.drop(['專業度_x', '專業度_y'], 1, inplace=True)


            # 計算符合、不符合、無法辨識學院人數
            df_group = df.groupby(['上市別','公司代碼簡稱','資料源年月','專業度']).size().reset_index(name='人數')
            df_group = df_group.pivot_table(index=['上市別','公司代碼簡稱','資料源年月'], columns='專業度', values='人數').reset_index()
            
            #解決沒有"是"、"否"、"無法辨識學院"的情形
            for status in ['是','否','無法辨識學院']:
                if status not in list(df_group):
                    df_group[status] = np.nan
                    
            df_group = df_group[['上市別','公司代碼簡稱','資料源年月','是','否','無法辨識學院']]
            df_group.rename(columns = {'是' : '獨立董事_'+pro+'_符合人數',
                                     '否' : '獨立董事_'+pro+'_不符合人數',
                                     '無法辨識學院' : '獨立董事_'+pro+'_無法辨識學院人數'
                                    }, inplace = True)
            df_group.fillna(0, inplace = True)

            # merge 統計表 & 計算後的人數資料
            boardTable = boardTable.merge(df_group, how='left', on=['上市別','公司代碼簡稱','資料源年月'])

            
            # 計算專業燈號
            boardTable['獨立董事_'+pro+'_燈號'] = np.where(boardTable['獨立董事_'+pro+'_符合人數'] > 0, '符合',
                                               np.where(boardTable['獨立董事_'+pro+'_不符合人數'] > 0, '不符合',
                                                   np.where(boardTable['獨立董事_'+pro+'_無法辨識學院人數'] > 0, '無法辨識學院',
                                                       '沒有揭露')))
            # 計算分數
            boardTable['獨立董事_'+pro+'_分數'] = np.where(boardTable['獨立董事_'+pro+'_燈號'] =='符合', 1, 0)

        return boardTable
    
    @staticmethod
    def getAllBoardScore():
        global boardTable
        boardTable['獨立董事_總分'] = boardTable['獨立董事_財務_分數']+boardTable['獨立董事_會計_分數']+boardTable['獨立董事_法務_分數']+boardTable['獨立董事_IT_分數']+boardTable['獨立董事_產業專業_分數']
        return boardTable


i = 0
for industry in needIndustry:
    i += 1
    print('第'+str(i)+'個產業：'+industry)
    df = BoardFunc.openIndsData(industry)
    proDict = BoardFunc.getProDict()
    boardTable = BoardFunc.crtBoardTable(df)
    directors = BoardFunc.getIndependentDirector(df)
    boardTable = BoardFunc.countBoardNumPeople(directors)
    boardTable = BoardFunc.getAndMergeInfo(directors)
    boardTable = BoardFunc.countAndMergeSpecialty(directors)
    boardTable = BoardFunc.getAllBoardScore()
    boardTable
    boardTable.to_excel(dataPath + f'3_獨立董事_{industry}_{mmdd}.xlsx',
                                    encoding = 'utf_8_sig', index = False
                                 )


