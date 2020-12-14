# -*- coding: utf-8 -*-
"""
Created on Sun Mar  8 00:17:10 2020

@author: Dausyy
"""

# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import requests
import json
import xlwt
import datetime
import os

CheckPath1 = os.path.exists('Results')
if CheckPath1 == False:
    os.mkdir('Results')
    
CheckPath2 = os.path.exists('Results/Internalflow')
if CheckPath2 == False:
    os.mkdir('Results/Internalflow')

################函数定义################
#############################################################

#用城市代码搜索城市名称
CityCodeLibrary_json = open('CityCodeLibrary.json', encoding ='utf-8-sig') 
RawCityCodeLibrary=json.load(CityCodeLibrary_json)
CityCodeLibrarySheet = RawCityCodeLibrary['工作表 1 - 行政区划乡镇清单201910 (3)']

def CityCode2Name(CityCode): 
    
    CityName = ''
    
    CodeKey = '地级市名称'
    
    SearchKey_1 = '地级市代码'
    
    SearchKey_2 = '省代码'
    
    if CityCode in ['310000', '110000', '120000', '500000']:
        for rownum in range(len(CityCodeLibrarySheet)):
            if CityCodeLibrarySheet[rownum][SearchKey_2] == CityCode:            
                CityName=CityCodeLibrarySheet[rownum][CodeKey]
    else:        
        for rownum in range(len(CityCodeLibrarySheet)):    
            if CityCodeLibrarySheet[rownum][SearchKey_1] == CityCode:            
                CityName=CityCodeLibrarySheet[rownum][CodeKey]
        
    return CityName


#web to jsontext
def JsonTextConvert(text): 

    """Text2Json

    Arguments:

        text {str} -- webContent

    Returns:

        str -- jsonText

    """    

    text = text.encode('utf-8').decode('unicode_escape')

    head, sep, tail = text.partition('(')

    tail=tail.replace(")","")

    return tail

#URL formatting

today_datetime = datetime.datetime.today()

today_date = today_datetime.strftime('%Y%m%d')


CityCode_json = open('CityCode.json', encoding ='utf-8-sig') 

RawCityCode=json.load(CityCode_json)

CityCodeSheet = RawCityCode['Sheet1']




def UrlFormat_internalflow(CityCode,today_date):
    
    url_internalflow='https://huiyan.baidu.com/migration/internalflowhistory.jsonp?dt=city&id={0}&date={1}&callback=jsonp_'.format(CityCode,today_date)
    return url_internalflow


################执行################
#############################################################

for h in range(len(CityCodeSheet)):
    workbook = xlwt.Workbook(encoding='utf-8') #Create workbook
    
    CityCode = CityCodeSheet[h]['CityCode']

    worksheet = workbook.add_sheet('internalflow',cell_overwrite_ok=True)

    table_head = ['Date_ref','value_ref','Date_new','value_new']

    url_internalflow=UrlFormat_internalflow(CityCode,today_date)
    req_internalflow = requests.get(url_internalflow)
    text_internalflow = req_internalflow.text
    rawData_internalflow=json.loads(JsonTextConvert(text_internalflow))
    sort_internalflow=json.dumps(rawData_internalflow,sort_keys=True)
    sortedData_internalflow=json.loads(sort_internalflow)
    
    data_internalflow= sortedData_internalflow['data']
    list_internalflow = data_internalflow['list']

    key = list_internalflow.keys()
    

    index = 1
          
    for i in range(len(table_head)):
        worksheet.write(0,i,table_head[i])

    for l in key:
        if l < '20200101':
            worksheet.write(index,0,l)
            worksheet.write(index,1,list_internalflow[l])
            numday_2019 = index
            index = index + 1
        else:                            
            worksheet.write(index-numday_2019,2,l)
            worksheet.write(index-numday_2019,3,list_internalflow[l])
            index = index + 1
                                       

    CityName = CityCode2Name(CityCode)
    
    filename = CityName+'.xls'

    workbook.save('Results/Internalflow/'+filename) 


    
    
    



