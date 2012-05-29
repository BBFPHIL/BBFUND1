#!/usr/bin
import os
import sys
#sys.path.append('/Users/pwalker/BBF_PROP_ALGOS/FINAL_PRODUCT_ALGOS/')
import yqldata
import re



#Rule based functions                                                                                                                                                      
#Purpose: to determine of these rules adhere to our trading strategy                                                                                                       

#Check if price of stock is True or not                                                                                                                                    
def price_req(price):
    max_price = 10.00
    return price < max_price

#Check if pe requirement is true or not                                                                                                                                    
def pe_req(pe):
    max_pe = 25.00
    return pe < max_pe and pe > 0


# Check if eps is true or not                                                                                                                                              
def eps_req(eps):
    min_eps = 0
    return eps > min_eps

#Check for resistance levels to facilitate breakout possibility in near future                                                                                             
#Parameters required:                                                                                                                                                      
# Price, one year target, 52 week high, 52 week low                                                                                                                        
def resistance_req(price, oneyr, weekhi):
    return oneyr > price*1.85 and weekhi > price*1.95

#Create floats from basic api inputs collected                                                                                                                             
#Disregard N/A inputs as '1000' string and then convert                                                                                                                    
#That to a float                                                                                                                                                           
def create_float(num):
    if num == 'N/A':
        num = '-10000'
        return float(num)
    else:
        return float(num)



