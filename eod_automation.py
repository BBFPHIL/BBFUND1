#!/usr/bin/env python

import os
import sys
import win32com.client
import xlrd
#import xlutils
import xlwt
import glob
import time
from xlrd import open_workbook
import datetime
from xlrd import open_workbook,empty_cell



##################
'''
EOD_AUTO_REPORT

Date: 5/13/2012

By:Phillip Walker

Goal: This software aims to automate the daily trading P&L results for Operations.
This software will automatically collect all important data (specified by the VP) from orginal excel and
then automatically format the information into a new excel sheet.

'''
#################



#################
'''
Collect all important cells from orginal excel file
*****The Work Books need to be automated with user input!******
'''
#################
def collectAllDataList(sh):
	data = []
	for rownum in range(sh.nrows):
		data.append(sh.row_values(rownum))
	
	return data


##############
#Fucntion that converts .xlsx to .xls for xlrd/xlwt/xlutil capatibility
##############

def convert_xlsx(wb):


	xlsx_files = glob.glob('wb') #This must be changed often


	if len(xlsx_files) == 0: 
    		print "The file is already in xls format"
		#raise RuntimeError('No XLSX file to convert.') 

	xlApp = win32com.client.Dispatch('Excel.Application') 

	for file in xlsx_files: 
    		xlWb = xlApp.Workbooks.Open(os.path.join(os.getcwd(), file))
 
    		xlWb.SaveAs(os.path.join(os.getcwd(), file.split('.xlsx')[0] + '.xls'), FileFormat=1) 


	#xlApp.Quit() 

	# Delete or comment out the following lines if you want to preserve the 
	# original XLSX files. 

	time.sleep(2) # give Excel time to quit, otherwise files may be locked 

	for file in xlsx_files: 
    		os.unlink(file)
 

### CREATE EMPTY quotes to no quote empty

def cell_type(cellvar):  #DO NOT APPLY TO CELLS THAT ARE SUPPOSED TO BE TEXT!!!

	if cellvar == '\'\'':
		cellvar.strip("'")
		cellvar = "-"
		return cellvar
	else:
		return float(cellvar)
	


############
############
############
									
						
			


#TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING TESTING
if __name__ == "__main__":


	#####Creating variables for program in global scope#####
	data = []


	#####Grab file name#####
	file = raw_input("Hello Superior, please enter file name: \n")
	
	convert_xlsx(file) #Convert to xls if it is not already in xls

	wb = xlrd.open_workbook(file)
	print "File opened! \n"

	#####Grab file sheet!######
	sh = wb.sheet_by_index(0) #Grab the sheet in number 1
	
	#####collect all excel data into dictionary for querying later####
	data = collectAllDataList(sh)
	

	###################### DELETE THIS #######################

	#Now grab list within list where sublist[0] == "ASIAN PRODUCTS"

	nicheSum = [] #Summary of Niches to be exported to new excel write material
	nicheData = [] #The long list of data correlating to the niche sectors
	nicheHeader = []#Header files for niche summary

	#Write summary to excel!
		
	#Sweet bold styling text for overall summary
	style = xlwt.XFStyle()
	font = xlwt.Font()
	font.name = 'Arial'
	font.bold = True
	font.height=20*16
	style.font = font

	report = xlwt.Workbook()

	#Adding a sheet	
	reportSheet = report.add_sheet('Sheet1')

	#Make Columns wide enough to see all numbers easily
	x_width = 0
	for each in range(4):
		column_width = reportSheet.col(x_width)
		column_width.width = 440 * 20 #20 characters wide
		x_width += 1

	#Grab niche summary header!
	nicheHeader.append(sh.row_values(0,0))

	#Create headers in file
	#reportSheet.write(0,0, str(nicheHeader[:5]))


	#CHANGE OR DELETE THESE VARIABLES FOR CUSTOMIZATION OF REPORT SUMMARY
	search1 = 'ASIAN PRODUCTS'
	search3 = 'US SPEC SITS'
	search4 = 'CREDIT (WALSH)'
	search5 = 'L/S_TECH (SOLOMON)'
	search6 = 'L/S_IND&TRAN (MORT)'
	search7 = 'L/S_HEALTH (AMIEL)'
	
	
	#Create header rows
	reportSheet.write(0,1, 'P&L', style)
	reportSheet.write(0,2, 'LMV (MM)', style)
	reportSheet.write(0,3, 'SMV (MM)', style)

		

		
	for sublist in data:
		if sublist[0] == search1:
			
			
			#Name of Niche is first row
			reportSheet.write(1,0, str(sublist[0:1]).strip('[,],u,"'), style) #Create prototypes Name ASIAN PRODUCTS

			#P&L
			c1 = str(sublist[4:5]).strip('[,],u')

			c1a = cell_type(c1)
			
			reportSheet.write(1,1, c1a, style) #Create P&L Slice
		
			
			########## END ###########

			c2 = str(sublist[6:7]).strip('[,],u') #Convert to floats

			c2a = cell_type(c2)
			if c2a == '-':
				reportSheet.write(1,2, c2a, style)#Create LMV

			else:
				c2b = c2a/1000
				reportSheet.write(1,2, round(c2b), style)#Create LMV
				
			########### END ##########
			
			c3 = str(sublist[7:8]).strip('[,],u') #CONVERT
			
			c3a = cell_type(c3)
			if c3a == '-':
				reportSheet.write(1,3, c3a, style) #Create SMV

			else:
				c3b = c3a/1000
				reportSheet.write(1,3, round(c3b), style) #Create SMV
	

		
		if sublist[0] == search3: #Watch out... 2 is not included in list!

			reportSheet.write(2,0, str(sublist[0:1]).strip('[,],u,"'), style) #Create prototypes
	
			########### END ############

			#reportSheet.write(2,1, str(sublist[4:5]).strip('[,],u,"'), style) #Create P&L Slice
			#P&L
			c4 = str(sublist[4:5]).strip('[,],u')

			c4a = cell_type(c4)
			
			reportSheet.write(2,1, c4a, style) #Create P&L Slice
			############ END ##############


			#reportSheet.write(2,2, str(sublist[6:7]).strip('[,],u,"'), style) #Create LMV
			#LMV
			c5 = str(sublist[6:7]).strip('[,],u') #Convert to floats

			c5a = cell_type(c5)
			if c5a == '-':
				reportSheet.write(2,2, c5a, style)#Create LMV

			else:
				c5b = c5a/1000
				reportSheet.write(2,2, round(c5b), style)#Create LMV

			############## END #############
				

			#reportSheet.write(2,3, str(sublist[7:8]).strip('[,],u,"'), style) #Create SMV
			
			c6 = str(sublist[7:8]).strip('[,],u') #CONVERT
			c6a = cell_type(c6)
			if c6a == '-':
				reportSheet.write(2,3, c6a, style) #Create SMV

			else:
				c6b = c6a/1000
				reportSheet.write(2,3, round(c6b), style) #Create SMV

		
		if sublist[0] == search4:
		
			reportSheet.write(3,0, str(sublist[0:1]).strip('[,],u,"'), style) #Create prototypes
			

						########### END ############
			#reportSheet.write(3,1, str(sublist[4:5]).strip('[,],u,"'), style) #Create P&L Slice
			#P&L
			c7 = str(sublist[4:5]).strip('[,],u')

			c7a = cell_type(c7)
			
			reportSheet.write(3,1, c7a, style) #Create P&L Slice

			################### END ############

			#reportSheet.write(3,2, str(sublist[6:7]).strip('[,],u,"'), style) #Create LMV
			#LMV
			c8 = str(sublist[6:7]).strip('[,],u') #Convert strings

			c8a = cell_type(c8)
			if c8a == '-':
				reportSheet.write(3,2, c8a, style)#Create LMV
			else:
				c8b = c8a/1000
				reportSheet.write(3,2, round(c8b), style)#Create LMV

			###################### END #############################

			#reportSheet.write(3,3, str(sublist[7:8]).strip('[,],u,"'), style) #Create SMV
			#SMV
			c9 = str(sublist[7:8]).strip('[,],u') #CONVERT
			
			c9a = cell_type(c9)
			
			if c9a == '-':
				reportSheet.write(3,3, c9a, style) #Create SMV

			else:
				c9b = c9a/1000
				reportSheet.write(3,3, round(c9b), style) #Create SMV



		if sublist[0] == search5:

			reportSheet.write(4,0, str(sublist[0:1]).strip('[,],u,"'), style) #Create prototypes
			
			#reportSheet.write(4,1, str(sublist[4:5]).strip('[,],u,"'), style) #Create P&L Slice
			#P&L
			c10 = str(sublist[4:5]).strip('[,],u')

			c10a = cell_type(c10)
			
			reportSheet.write(4,1, c10a, style) #Create P&L Slice
			
			###################### END ########################

			#reportSheet.write(4,2, str(sublist[6:7]).strip('[,],u,"'), style) #Create LMV
			#LMV
			c11 = str(sublist[6:7]).strip('[,],u') #Convert strings

			c11a = cell_type(c11)
			if c11a == '-':
				reportSheet.write(4,2, c11a, style)#Create LMV
			else:
				c11b = c11a/1000
				reportSheet.write(4,2, round(c11b), style)#Create LMV
				
			############## END ################
			
			#reportSheet.write(4,3, str(sublist[7:8]).strip('[,],u,"'), style) #Create SMV	
			#SMV
			c12 = str(sublist[7:8]).strip('[,],u') #CONVERT
			c12a = cell_type(c12)
			if c12a == '-':
				reportSheet.write(4,3, c12a, style) #Create SMV
			else:
				c12b = c12a/1000
				reportSheet.write(4,3, round(c12b), style) #Create SMV


		if sublist[0] == search6:
		
			reportSheet.write(5,0, str(sublist[0:1]).strip('[,],u,"'), style) #Create prototypes
				

			#reportSheet.write(5,1, str(sublist[4:5]).strip('[,],u,"'), style) #Create P&L Slice
			#P&L
			c13 = str(sublist[4:5]).strip('[,],u')

			c13a = cell_type(c13)
			
			reportSheet.write(5,1, c13a, style) #Create P&L Slice
			################END#############
			

			#reportSheet.write(5,2, str(sublist[6:7]).strip('[,],u,"'), style) #Create LMV
			#LMV
			c14 = str(sublist[6:7]).strip('[,],u') #Convert strings

			c14a = cell_type(c14)
			if c14a == '-':
				reportSheet.write(5,2, c14a, style)#Create LMV
			else:
				c14b = c14a/1000
				reportSheet.write(5,2, round(c14b), style)#Create LMV
			################## end ########################

			#reportSheet.write(5,3, str(sublist[7:8]).strip('[,],u,"'), style) #Create SMV
			#SMV
			c15 = str(sublist[7:8]).strip('[,],u') #CONVERT
			c15a = cell_type(c15)
			if c15a == '-':
				reportSheet.write(5,3, c15a, style) #Create SMV
			else:
							
				c15b = c15a/1000
				reportSheet.write(5,3, round(c15b), style) #Create SMV

		if sublist[0] == search7:
		
			reportSheet.write(6,0, str(sublist[0:1]).strip('[,],u,"'), style) #Create prototypes
			
			#reportSheet.write(6,1, str(sublist[4:5]).strip('[,],u,"'), style) #Create P&L Slice
			#P&L
			c16 = str(sublist[4:5]).strip('[,],u')

			c16a = cell_type(c16)
			
			reportSheet.write(6,1, c16a, style) #Create P&L Slice
			##################### END ###############

			#reportSheet.write(6,2, str(sublist[6:7]).strip('[,],u,"'), style) #Create LMV
			#LMV
			c17 = str(sublist[6:7]).strip('[,],u') #Convert strings

			c17a = cell_type(c17)
			
			if c17a == '-':
				reportSheet.write(6,2, c17a, style)#Create LMV
			else:
				c17b = c17a/1000
				reportSheet.write(6,2, round(c17b), style)#Create LMV
			################### END ###################
			
			#reportSheet.write(6,3, str(sublist[7:8]).strip('[,],u,"'), style) #Create SMV
			#SMV
			c18 = str(sublist[7:8]).strip('[,],u') #CONVERT
			c18a = cell_type(c18)
			if c18a == '-':
				reportSheet.write(6,3, c18a, style) #Create SMV
			else:
							
				c18b = c18a/1000
				reportSheet.write(6,3, round(c18b), style) #Create SMV


		
		

	#TESTING TESTING#
	print "\n \n \n"
	
	#Find beginning value and ending value of bulk info data to display with spaces
	#At specific coordinate markers
	
	#Create new headers
	reportSheet.write(10,0, '  ', style)
	reportSheet.write(10,1, 'Position', style)
	reportSheet.write(10,2, 'Price', style)
	reportSheet.write(10,3, 'Change', style)
	reportSheet.write(10,4, 'Daily', style)
	reportSheet.write(10,5, 'Tag Level 1', style)

	#Begin collecting & Writing bulk data

	#Prep variables
	row = 12
	
	preS = 0
	postS = 1

	for eachSL in data[3:-1]:
		for col in xrange(6):
			
			if eachSL[0] == 'ASIAN PRODUCTS':
			

				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))

			if eachSL[0] == 'US SPEC SITS':

				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))
		
			
			if eachSL[0] == 'CREDIT (WALSH)':

				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))

			
			if eachSL[0] == 'L/SIND&TRAN (MORT)':

				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))

			if eachSL[0] == 'L/S_HEALTH (AMIEL)':

				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))

			
			if eachSL[0] == 'L/S_TECH (SOLOMON)':

				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))

			else:
				reportSheet.write(row,col, str(eachSL[preS:postS]).strip('[,],u,"'))
			
			
			
			preS += 1
			postS += 1
			
			col += 1
			

		
			if col == 6:

				row += 1
				preS = 0
				postS = 1
				break
				


		


			
	#Make files never the same when saved
	#Beginning on May 22nd 2012
	#datetime.datetime.today()
	date = ()
	date = str(datetime.datetime.today())
	month = date[5:7]
	day = date[8:10]
	year = date[0:4]

	fileNameToday = 'EOD_Report_'+month+'_'+day+'_'+year+'_'+'.xls'

	report.save('EOD_Report_'+month+'_'+day+'_'+year+'_'+'.xls')
	print "Finished!"

	########################################




