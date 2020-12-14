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
import os

CheckPath1 = os.path.exists('Results')
if CheckPath1 == False:
    os.mkdir('Results')
    
CheckPath2 = os.path.exists('Results/Cityrank')
if CheckPath2 == False:
    os.mkdir('Results/Cityrank')

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
Date_json = open('Date.json', encoding ='utf-8-sig') 

RawDate=json.load(Date_json)

DateSheet = RawDate['Sheet1']


CityCode_json = open('CityCode.json', encoding ='utf-8-sig') 

RawCityCode=json.load(CityCode_json)

CityCodeSheet = RawCityCode['Sheet1']


in_out = ['in', 'out']

def UrlFormat_daily(CityCode,in_out,date):
    
    url_daily='https://huiyan.baidu.com/migration/cityrank.jsonp?dt=city&id={0}&type=move_{1}&date={2}&callback=jsonp_'.format(CityCode,in_out,date)
    return url_daily

################执行################
#############################################################

for h in range(len(CityCodeSheet)):
    workbook = xlwt.Workbook(encoding='utf-8') #Create workbook
    
    CityCode = CityCodeSheet[h]['CityCode']
    
    for i in range(len(DateSheet)):
        date = DateSheet[i]['Date']
        worksheet = workbook.add_sheet(date,cell_overwrite_ok=True) #Create worksheet
    
        for j in in_out:        
            url_daily=UrlFormat_daily(CityCode,j,date)
            #print(url_daily)
            req = requests.get(url_daily)
            text = req.text
            rawData = json.loads(JsonTextConvert(text))
            data = rawData['data']
            list = data['list']
        
        
            table_head = ['city_name','value']#表头
        
            worksheet.write(0,0,'In')
            worksheet.write(0,3,'Out')
        
            index = 2
        
            if j == 'in':
                for a in range(len(table_head)):
                    worksheet.write(1,a,table_head[a])
                    for b in list:
                        worksheet.write(index,0,b['city_name'])
                        worksheet.write(index,1,b['value'])
                        index = index + 1
            
            else:
                for a in range(len(table_head)):
                    worksheet.write(1,a+3,table_head[a])
                    for b in list:
                        worksheet.write(index,3,b['city_name'])
                        worksheet.write(index,4,b['value'])                                              
                        index = index + 1
   

    CityName = CityCode2Name(CityCode)
    
    filename = CityName+'.xls'

    workbook.save('Results/Cityrank/'+filename) #保存表


    
    
    



