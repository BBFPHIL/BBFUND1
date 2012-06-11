#!/usr/bin

import os
import sys
import urllib
import re
import yqldata
import rm1


'''
Version 1.1
R1BBF Characteristics of an Attractive Investment
Name: R1BBF Algo (Research Algorithm 2)
Type: Fundamental Research Algorithm
Goal: Filter entire stock exchanges for stocks that fit our specific investment model
Expected Mortality Rate: 50-60 symbols (2%)
*This algorithm evolves as the assets under management fluctuate. i.e. if fund grows to
$10,000 AUM, then price may be < $20.00 for investment.*
'''


data = []
data = yqldata.all_symbols()
c = 0

#print data[:-1] debugging                                                                                                         
dataQueryList = []
dataQueryList = data[:500] #Set this variable as you see fit                                                                                                               
#print dataQueryList

#for sym in dataQueryList[:-1]:  #######Testing only first two for time's sake and server capabilities####### MANUAL MAINTENANCE ######                                     
for num, sym in enumerate(dataQueryList[:-1]): #Enumerate is more efficient 
    
    #Set variables for decision making                                                                                                                                     
    one_year =  rm1.create_float(yqldata.get_one_year_target(sym["symbol"]))
    price = float(yqldata.get_price(sym["symbol"]))
    week_high_52 = rm1.create_float(yqldata.get_52_week_high(sym["symbol"]))
    week_low_52 = rm1.create_float(yqldata.get_52_week_low(sym["symbol"]))
    pe = rm1.create_float(yqldata.get_pe(sym["symbol"]))
    eps = rm1.create_float(yqldata.get_earnings_per_share(sym["symbol"]))
    short_ratio = rm1.create_float(yqldata.get_short_ratio(sym["symbol"]))


    '''                                                                                                                                                                    
    ***** ALL COMMENTED OUT CODE IS FOR DEBUGGING PURPOSES *******

    print sym["symbol"]                                                                                                                                                    
    print one_year, " one year target\n"                                                                                                                                   
    print price, " price \n"                                                                                                                                               
    print week_high_52, " high \n"                                                                                                                                         
    print week_low_52, " low \n"                                                                                                                                           
    print pe, " pe \n"                                                                                                                                                     
    print eps, " eps\n"                                                                                                                                                    
    print short_ratio, "\n"                                                                                                                                                
    

    #Begin filtering process                                                                                                                                               
    #print sym["symbol"], "\n"                                                                                                                                             
    #Filter 1 .. find all appropriately priced stocks under 10.00                                                                                                          
    #For farther examination                                                                                                                                               
    #print price_req(price), "\n"                                                                                                                                          
    '''

    if rm1.price_req(price) == False:
        dataQueryList.pop(c)
        continue
    #Filter 2 -- checking PE                                                                                                                                               
    #print pe_req(pe), "\n"           

    if rm1.pe_req(pe) == False:
        dataQueryList.pop(c)
        continue
    #Filter 3 -- checking eps                                                                                                                                              
    #print eps_req(eps), "\n"                                                                                                                                              

    if rm1.eps_req(eps) == False:
        dataQueryList.pop(c)
        continue
    #Filter 4 -- checking Resistance and arbitrage potential                                                                                                               
    #print resistance_req(price, one_year, week_high_52, week_low_52)                                                                                                      

    if rm1.resistance_req(price, one_year, week_high_52) == False:
        dataQueryList.pop(c)
        continue


    c += 1
    #Increment count for popping                                                                                                                                           

print "The following are the results from R1BBF analysis. Please change IncomeData3.xls \n"
print "symbols to the ones below and run R2BBF for each of them"

print dataQueryList[:-1]

