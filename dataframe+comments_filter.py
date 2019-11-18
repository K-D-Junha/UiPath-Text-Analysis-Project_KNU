#!/usr/bin/env python
# coding: utf-8

# In[24]:


import pandas as pd
from pandas import DataFrame as df


#데이터프레임 생성, csv파일로 저장하고 다시 불러올 경우 index가 변환되어서 우선 그냥 사용
df1 = df(data = {'Galaxy Note10':['O','O','X', 'X', 'O', 'O'],
                'Galaxy S10':['O','O','X', 'O', 'O', 'O'],
                'Galaxy Note9':['O','O','X', 'X', 'O', 'O'],
                'iPhone11':['O', 'X', 'X', 'X', 'O', 'O'],
                'iPhoneX':['O','X', 'X', 'X', 'O', 'O']},
 
        index=['Facial', 'Fingerprint','Iris', 'Earphone', 'NFC', 'HDR'],
        columns = ['Galaxy Note10', 'Galaxy S10','Galaxy Note9', 'iPhone11', 'iPhoneX'])
#df1.to_csv("list.csv")


# In[25]:


#print(df1)


# In[26]:


def get_keyword(DataFrame):
    
    keyword = []
    
    product_name = df1.columns[0]       #'갤럭시 노트10' 제품명 이름 받기
                                        #다른 제품으로 변경 시 배열숫자 변경
                                        # 0.갤럭시노트10, 1.갤럭시s10, 2. 갤럭시노트9, 3.아이폰11, 4.아이폰x
    
    for i in range(0,len(df1[product_name])):
        if df1[product_name][i] == "O":     #기능 리스트 중 있는 리스트만 추출
            keyword.append(df1.index[i])
            
    return keyword


# In[35]:


def keyword_comments(trans_list = None, DataFrame = None):
    result = []
    
    keyword = get_keyword(DataFrame)    #키워드 목록 받아오기
    
    for i in range(0,len(trans_list)):
        if (keyword[0].lower() in trans_list[i]) or (keyword[0] in trans_list[i]):
            #키워드 리스트 중 0=facial, 1=..., 임시로 하나만 추가해둠, 키워드 리스트 별로 모두 나타내는 것 필요하다면 추가가능
            result.append(trans_list[i])
    
    return result


# In[36]:


test = ["hello hi","fingerprint recognition is bad", "facial bad","Facial good", "good i want you hi", "dd", "으으으으"]



print(keyword_comments(test, df1))


# In[ ]:





# In[ ]:




