# -*- coding: utf-8 -*-
"""
Created on Sun Mar  8 00:17:10 2020

@author: Dausyy
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Mar  8 00:01:51 2020

This is a temporary script file.
"""

import requests
import json
import xlwt
import os

CheckPath1 = os.path.exists('Results')
if CheckPath1 == False:
    os.mkdir('Results')
    
CheckPath2 = os.path.exists('Results/Hist')
if CheckPath2 == False:
    os.mkdir('Results/Hist')
    
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


CityCode_json = open('CityCode.json', encoding ='utf-8-sig') 

RawCityCode=json.load(CityCode_json)

CityCodeSheet = RawCityCode['Sheet1']


in_out = ['in', 'out']

def UrlFormat_hist(CityCode,in_out):
    
    url_hist='https://huiyan.baidu.com/migration/historycurve.jsonp?dt=city&id={0}&type=move_{1}&callback=jsonp_'.format(CityCode,in_out)
    return url_hist


################执行################
#############################################################

for h in range(len(CityCodeSheet)):
    workbook = xlwt.Workbook(encoding='utf-8') #Create workbook
    
    CityCode = CityCodeSheet[h]['CityCode']

    worksheet = workbook.add_sheet('hist',cell_overwrite_ok=True)

    table_head = ['Date_ref','value_ref','Date_new','value_new']

    worksheet.write(0,0,'In')
    worksheet.write(0,5,'Out')

    for j in in_out:
        url_hist=UrlFormat_hist(CityCode,j)
        req_hist = requests.get(url_hist)
        text_hist = req_hist.text
        rawData_hist=json.loads(JsonTextConvert(text_hist))

        data_hist= rawData_hist['data']
        list_hist = data_hist['list']

        key = list_hist.keys()
    

        index = 1
        numday_2019 = 62

        if j == 'in':            
            for i in range(len(table_head)):
                worksheet.write(1,i,table_head[i])

            for l in key:
                if l < '20200101':
                    worksheet.write(index+1,0,l)
                    worksheet.write(index+1,1,list_hist[l])
                    numday_2019 = index
                    index = index + 1
                else:                              
                    worksheet.write(index-numday_2019+1,2,l)
                    worksheet.write(index-numday_2019+1,3,list_hist[l])
                    index = index + 1
                            
        else:
            for i in range(len(table_head)):
                worksheet.write(1,i+5,table_head[i])

            for l in key:
                if l < '20200101':
                    worksheet.write(index+1,0+5,l)
                    worksheet.write(index+1,1+5,list_hist[l])
                    numday_2019 = index
                    index = index + 1
                else:           
                    worksheet.write(index-numday_2019+1,2+5,l)
                    worksheet.write(index-numday_2019+1,3+5,list_hist[l])
                    index = index + 1            

    CityName = CityCode2Name(CityCode)
    
    filename = CityName+'.xls'

    workbook.save('Results/Hist/'+filename) 


    
    
    



