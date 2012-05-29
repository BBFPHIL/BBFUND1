#!/usr/bin
import os
import sys
import xlrd
import xlwt



'''
Goal: statementdata.py is a library that is going to be built in order to
extract financial statement and other R2 data from an excel sheet that is constantly querying
data from a web page.
This information will be the key data api to building the R2BBF algorithm
as yqldata.py is core to the R1BBF algorithm
'''

'''
Open up all the approriate workbooks and sheets to create handles
'''

file = 'IncomeData3.xls'

wb = xlrd.open_workbook(file)
print "File opened! \n"

#####Grab file sheets 1 - 7 ######
sh1 = wb.sheet_by_name(u'Sheet1')#Grab sheet number 1 aka Income Statement
sh2 = wb.sheet_by_name(u'Sheet2')#Grab sheet number 2 aka Balance Statement
sh3 = wb.sheet_by_name(u'Sheet3')#Grab sheet number 3 aka Cash Flow Statement
sh4 = wb.sheet_by_name(u'Sheet4')#Grab Sheet number 4 aka Holders
sh5 = wb.sheet_by_name(u'Sheet5')#Grab sheet number 5 aka Insider Transactions
sh6 = wb.sheet_by_name(u'Sheet6')#Grab sheet number 6 aka Analyst Estimates
sh7 = wb.sheet_by_name(u'Sheet7')#Grab sheet number 7 aka Key Statistics
#### END ###### SHEET HANDLES #########

#Grab important data with the following functions
'''
Functions that grab all data from income statement
'''

def income_yrs():
    
    #Grab last three years of income totals
    year1 = sh1.cell(38, 4)
    year2 = sh1.cell(38, 3)
    year3 = sh1.cell(38, 2)
    
    #convert to string --> then strip junk --> convert to float
    year1 = str(year1) 
    year1 = year1.strip('number, :')
    year1 = float(year1)
    
    year2 = str(year2)
    year2 = year2.strip('number, :')
    year2 = float(year2)
    
    year3 = str(year3)
    year3 = year3.strip('number, :')
    year3 = float(year3)
    
    source = {}
    source["year1"] = year1
    source["year2"] = year2
    source["year3"] = year3
    
    return source


#Save results of income statement years into income dictionary
incomeYears = {}
incomeYears = income_yrs()
#print incomeYears["year3"]

'''
END OF Income Statement extraction
'''
########################################
'''
Balance Sheet Statement
Data collection
'''
def balance_sheet():
    #Grab last three years of income totals
    year1 = sh2.cell(49, 4)
    year2 = sh2.cell(49, 3)
    year3 = sh2.cell(49, 2)
    
    #convert to string --> then strip junk --> convert to float
    year1 = str(year1) 
    year1 = year1.strip('number, :')
    year1 = float(year1)
    
    year2 = str(year2)
    year2 = year2.strip('number, :')
    year2 = float(year2)
    
    year3 = str(year3)
    year3 = year3.strip('number, :')
    year3 = float(year3)
    
    source = {}
    source["year1"] = year1
    source["year2"] = year2
    source["year3"] = year3

    return source

#Save balance sheet years to global variable for testing purposes
balanceYears = {}
balanceYears = balance_sheet()
print balanceYears["year1"]
'''
END OF Balance Sheet Statement extraction
'''
####################################
'''
Extract Cash Flow Years 1 - 3 data
'''
def cashflow_sheet():
    #Grab last three years of income totals
    year1 = sh3.cell(30, 4)
    year2 = sh3.cell(30, 3)
    year3 = sh3.cell(30, 2)
    
    #convert to string --> then strip junk --> convert to float
    year1 = str(year1) 
    year1 = year1.strip('number, :')
    year1 = float(year1)
    
    year2 = str(year2)
    year2 = year2.strip('number, :')
    year2 = float(year2)
    
    year3 = str(year3)
    year3 = year3.strip('number, :')
    year3 = float(year3)
    
    source = {}
    source["year1"] = year1
    source["year2"] = year2
    source["year3"] = year3

    return source

#Save cashflow years to a global dictionary... example ofr actual program
cashflowYears = {}
cashflowYears = cashflow_sheet()
print cashflowYears["year1"]

'''
END OF CASH FLOW SHEET EXTRACTION
'''
#####################################
'''
Extract the big name holders in the company
'''
def holder_sheet():
    holderList = []
    
    for rownum in range(sh4.nrows):
        holderList.append(sh4.row_values(rownum))
    
    return holderList 
    


#Save information into new list
holderNames = []
holderNames = holder_sheet()
print holderNames[8:-1]

'''
END of extracting major holder information
'''
####################################
'''
Extract insider transactions!
'''
def insider_sheet(): 
    insiderList = []
    
    for rownum in range(sh5.nrows):
        insiderList.append(sh5.row_values(rownum))
    
    return insiderList

'''
Note: The item of sale or buy is located in the fifth item of the sublist
'''

#Place insider list inside of a new list
insiderNames = []
insiderNames = insider_sheet()
print insiderNames

'''
END OF EXTRACTING INSIDER SALES AND BUYS
'''
#######################################
'''
Extracting strong buy, buy, hold, and sell for
current month, and last month for each
'''
def analyst_sheet():
    
    source = {}
    
    #number of current and last month recommendations of strong buy from analysts
    cur_mn_sb = sh6.cell(26,1)
    lst_mn_sb = sh6.cell(26,2)
    
    #convert to floats
    cur_mn_sb = str(cur_mn_sb) 
    cur_mn_sb = cur_mn_sb.strip('number, :')
    cur_mn_sb = float(cur_mn_sb)
    ##### LST MN CONVERT TO FLOAT #######
    lst_mn_sb = str(lst_mn_sb) 
    lst_mn_sb = lst_mn_sb.strip('number, :')
    lst_mn_sb = float(lst_mn_sb)
    
    #number of current and last month recommendations of buy from analysts
    cur_mn_b = sh6.cell(27,1)
    lst_mn_b = sh6.cell(27,2)
    
    cur_mn_b = str(cur_mn_b) 
    cur_mn_b = cur_mn_b.strip('number, :')
    cur_mn_b = float(cur_mn_b)
    ##### LST MN CONVERT TO FLOAT #######
    lst_mn_b = str(lst_mn_b) 
    lst_mn_b = lst_mn_b.strip('number, :')
    lst_mn_b = float(lst_mn_b)

    #number of current and last month recommendations of strong buy from analysts
    cur_mn_h = sh6.cell(28,1)
    lst_mn_h = sh6.cell(28,2)
    
    cur_mn_h = str(cur_mn_h) 
    cur_mn_h = cur_mn_h.strip('number, :')
    cur_mn_h = float(cur_mn_h)
    ##### LST MN CONVERT TO FLOAT #######
    lst_mn_h = str(lst_mn_h) 
    lst_mn_h = lst_mn_h.strip('number, :')
    lst_mn_h = float(lst_mn_h)
    
    #number of current and last month recommendations of strong buy from analysts
    cur_mn_s = sh6.cell(29,1)
    lst_mn_s = sh6.cell(29,2)
    
    cur_mn_s = str(cur_mn_s) 
    cur_mn_s = cur_mn_s.strip('number, :')
    cur_mn_s = float(cur_mn_s)
    ##### LST MN CONVERT TO FLOAT #######
    lst_mn_s = str(lst_mn_s) 
    lst_mn_s = lst_mn_s.strip('number, :')
    lst_mn_s = float(lst_mn_s)
    
    #Put inside dictionary
    source["cur_sb"] = cur_mn_sb
    source["lst_sb"] = lst_mn_sb
    
    source["cur_b"] = cur_mn_b
    source["lst_b"] = lst_mn_b
    
    source["cur_h"] = cur_mn_h
    source["lst_h"] = lst_mn_h
    
    source["cur_s"] = cur_mn_s
    source["lst_s"] = lst_mn_s
    
    return source

#Store in a new dictionary
analystData = {}
analystData = analyst_sheet()
#print analystData

'''
END OF EXTRACTING ANALYST DATA 
'''
###############################
'''
Begin extracting key statistics
'''

'''
def get_profitmargin():
    
    pm = sh7.cell(16, 1)
    
    #convert to str --> float
    pm = str(pm) 
    pm = pm.strip('number, :')
    pm = float(pm)

    
    return pm

profitMargin = get_profitmargin()
print profitMargin
    
    
def get_qtr_growth():

    qtr_growth = sh7.cell(24, 1)

    return qtr_growth

q = get_qtr_growth()
print q
'''
print "Successful!! \n"






#####collect all excel data into dictionary for querying later####
#data = collectAllDataList(sh)

