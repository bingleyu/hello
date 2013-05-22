# -*- coding: utf-8 -*-
""" Find the value searched """

import re
import xdrlib, sys ,os
import xlrd
import codecs
import os.path


fileName=u"C:/Zope/Instance/2.11.4/Extensions/bingle/员工通讯录.xls"
sheetName=u"总通讯录"



def stripChar(data='',char=".0"):
    # strip the tail of float numbers from excel file opend sheet
    sdata=[]
    for item in data.strip().split(' '):
        if item.endswith(char):
            sdata.append(str(item.strip("0").strip(".")))# twice strip incase stripped zero in phone number
        else:
            sdata.append(item.encode("utf-8"))
    return " ".join(sdata)

def stripSpace(s=''):
    us=unicode(s,"gbk").strip()
    return us
    
def find(value='yub'):
    "Returns the search result"
        
    filename=os.path.normcase(fileName)                       
    try:
        data = xlrd.open_workbook(filename)
    except Exception,e:
        return str(e)
    try:
        table=data.sheet_by_name(sheetName)
    except Exception,e:
        return str(e)
    nrows=table.nrows
    txlist=[]

    for i in range(0,nrows):
        result = " ".join(map(unicode,table.row_values(i)))
        txlist.append(result)
        
    searchitem=[]    
    value=value.strip()
    if value=="*":
        value=''
    if len(value.split())>1:
        value=''.join(value.split())
        searchitem.append('你找的是不是:')
    uvalue=unicode(value,"utf-8").strip()
    p=re.compile(uvalue,re.I)
    found=False
    
    for item in txlist:
        if p.search(item):
            # searchitem.append(uvalue)
            searchitem.append(stripChar(item))
            found=True
        else:
            continue
    if found: 
        return searchitem
    else:
        return ['Search failed for: '+value+'.'+'Try another word!']

def matchlist(filelist,value):
    value=value.strip()
    result=[]
    p=re.compile(value,re.I)
    found=False

    for item in filelist:
        if p.search(item):
            result.append(item)
            found=True
        else:
            continue
    if found:
        return result
    else:
        return False
    
def match(file,value):
    """ match value once at a time and return True if matched"""
    value=value.strip()
    if value in ["*","^","?"]:
        return False
    p=re.compile(value,re.I)
    if p.search(file):
        return True
    else:
        return False
        

