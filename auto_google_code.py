# -*- coding: utf-8 -*-
"""
Created on Tue Jul 18 16:50:29 2023

@author: sgvol
"""

# One hashtag - General notes
## Two hashtags - Code needs updating based on your Excel file and file paths 

from selenium import webdriver 				
import xlrd							     
import time



loc = ("C:/Users/sgvol/OneDrive/Documents/GitHub/Auto_Google_from_Excel/Example_Excel_Workbook.xls")  ## Insert path to Excel file. Must be .xls only    
wb = xlrd.open_workbook(loc)				           # Opens the workbook
sheet = wb.sheet_by_index(0)                           # Locates correct sheet within workbook
startrow = 0                                           # Creates variable 



# This block is to print values from the Excel file for only one row (prints within Python). This is so you check that you have the correct row and column numbers. 
# startrow is the number of the row that you want to search for from your Excel file 

def printfullname(startrow): 
    print(sheet.cell_value(startrow-1, 0)+ " " +sheet.cell_value(startrow-1, 1))    
    ## Columns to locate the full name. Change according to your file.  
def printphone(startrow):                              
	print(int(sheet.cell_value(startrow-1, 2)))                        # Make sure phone number is integer, not float
    ## Column to locate the phone number. Change according to your file
def printemail(startrow):
	print(sheet.cell_value(startrow-1, 3))             
    ## Column to locate the email address. Change according to your file
def printaddr(startrow):
	print(str(sheet.cell_value(startrow-1, 4))+ " " +str(sheet.cell_value(startrow-1, 5))+ " "+ str(sheet.cell_value(startrow-1, 6))+ " " + str(int(sheet.cell_value(startrow-1, 7))))
    ## Columns to locate the street address, city, state, and zip code. Change according to your file. 
    
    

# This block is the most generic. Allows you to specify what column to search with, starting on a specific row, and lets you choose how many Google searches it will go down the list
# Use the column and row number that you read in the Excel file (starting at 1, not 0)

def customsearch(column, starting_row, num_searches): 
    driver = webdriver.Chrome("C:/Users/sgvol/OneDrive/Desktop/ChromeDriver/chromedriver.exe")    ##  Insert address to chromedriver.exe
    main_window = driver.current_window_handle		
    tabnum = 1		
    
    for i in range(num_searches): 					
        driver.get("http://www.google.com/")			
        search_box = driver.find_element_by_name('q')			
        search_box.send_keys(str(sheet.cell_value(starting_row-1, column-1)))	
        time.sleep(3)				                                                               # pauses 3 seconds
        search_box.submit()
        driver.execute_script("window.open('https://www.google.com');")
        driver.switch_to.window(driver.window_handles[tabnum]) 
        starting_row = starting_row +1                              
        tabnum = tabnum + 1

    driver.switch_to.window(main_window)	



# Includes explanation how the code works
# Only need to input the row to start the search on. Will automatically perform 10 searches from the subsequent rows in the Excel file. 
    
def phone(startrow):					                  # startrow is the row number you want to start search in the Excel file 
	driver = webdriver.Chrome("C:/Users/sgvol/OneDrive/Desktop/ChromeDriver/chromedriver.exe")     # launches browswer   ##  Insert path to chromedriver.exe
	main_window = driver.current_window_handle		      # saves the 1st tab location. Will navigate back to 1st tab later
	tabnum = 1						                      # directs to next tab
	
	for i in range(10): 				                  # range is how many searches you want to run 
		driver.get("http://www.google.com/")
		search_box = driver.find_element_by_name('q')     # puts cursor into search bar 	
		search_box.send_keys(str(int(sheet.cell_value(startrow-1, 2))))	# begins search at the startrow. Uses the 3rd column. 
		time.sleep(3)	                                                # pauses 3 seconds, so Google allows the search 
		search_box.submit() 
		driver.execute_script("window.open('https://www.google.com');")
		driver.switch_to.window(driver.window_handles[tabnum])          # navigates to a new tab 

		startrow = startrow +1                                          # moves onto next row in Excel workbook
		tabnum = tabnum + 1                                             # allows new tabs 

	driver.switch_to.window(main_window)		                 	    # navigates Chrome back to 1st tab
    




# 3 More examples - Code for specific columns, so only the starting row needs to be defined    

def fullname(startrow):					    
	driver = webdriver.Chrome("C:/Users/sgvol/OneDrive/Desktop/ChromeDriver/chromedriver.exe")  ## Insert path to chromedriver.exe
	main_window = driver.current_window_handle		
	tabnum = 1						
	
	for i in range(10): 					                                                   ## range is how many searches you want to run 
		driver.get("http://www.google.com/")			
		search_box = driver.find_element_by_name('q')			
		search_box.send_keys(str(sheet.cell_value(startrow-1, 0))+" "+str(sheet.cell_value(startrow-1,1)))	## begins search at the startrow. Uses the 1st and 2nd columns 
		time.sleep(3)	                                                                       ## pauses 3 seconds
		search_box.submit()
		driver.execute_script("window.open('https://www.google.com');")
		driver.switch_to.window(driver.window_handles[tabnum])       
	
		startrow = startrow +1                                   
		tabnum = tabnum + 1 

	driver.switch_to.window(main_window)			
    
    
    
def email(startrow):					    
	driver = webdriver.Chrome("C:/Users/sgvol/OneDrive/Desktop/ChromeDriver/chromedriver.exe")  ## Insert path to chromedriver.exe
	main_window = driver.current_window_handle		
	tabnum = 1						
	
	for i in range(10): 					                                                   ## range is how many searches you want to run 
		driver.get("http://www.google.com/")			
		search_box = driver.find_element_by_name('q')			
		search_box.send_keys(str(sheet.cell_value(startrow-1, 3)))	## begins search at the startrow. Uses the 4th column 
		time.sleep(3)	                                            ## pauses 3 seconds
		search_box.submit()
		driver.execute_script("window.open('https://www.google.com');")
		driver.switch_to.window(driver.window_handles[tabnum])       
	
		startrow = startrow +1                                   
		tabnum = tabnum + 1 

	driver.switch_to.window(main_window)			                
    
    
    
def addr(startrow):					
	driver = webdriver.Chrome("C:/Users/sgvol/OneDrive/Desktop/ChromeDriver/chromedriver.exe")    ## Insert path to chromedriver.exe
	main_window = driver.current_window_handle		
	tabnum = 1						
	
	for i in range(10): 					                                                      ## range is how many searches you want to run 
		driver.get("http://www.google.com/")			
		search_box = driver.find_element_by_name('q')			
		search_box.send_keys(str(sheet.cell_value(startrow-1, 4))+" "+str(sheet.cell_value(startrow-1,5))+" "+str(sheet.cell_value(startrow-1,6))+' '+ str(int(sheet.cell_value(startrow-1, 7))))	
                                                                                                  ## begins search at the startrow. Uses 5th-8th columns 
		time.sleep(3)				                                                              ## pauses 3 seconds
		search_box.submit()
		driver.execute_script("window.open('https://www.google.com');")
		driver.switch_to.window(driver.window_handles[tabnum])        
		
		startrow = startrow +1                              
		tabnum = tabnum + 1

	driver.switch_to.window(main_window)			



    
    
    
    
