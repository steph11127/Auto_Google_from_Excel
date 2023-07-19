# Auto_Google_from_Excel
Python code that automatically performs Google web searches. Takes the information from an Excel file, and automatically Google searches down the column.  
Use cases: lists that need web searching, such as full names, addresses, phone numbers. 


## Dependencies 
- Excel file needs to be in .xls format only 
- Python needs packages selenium, xlrd, and time 
- Code utilizes ChromeDriver, which gets the Chrome web browswer that performs the automatic web searches. In my code, chromedriver.exe needs to be on the local machine and the path needs to be put into the code. Check out https://chromedriver.chromium.org/home to download ChromeDriver (about 12 megabytes). 

## Example_Excel_Workbook.xls
- columns: first_name, last_name, phone_number, email_address, street_address, city, state, zip_code
- rows: 10 examples, rows 2-11

## Python code
I use Python 3.9

The following adjustments need to be made, according to your file paths and Excel workbook: 
- Update path to Excel document 
	- line 17
- Update path to ChromeDriver.exe   
	- lines 46, 69, 94, 115, and 136
- Adjust columns of Excel document 
	- lines 28, 31, 34, 37, 76, 101, 122, and 143
- Adjust the number of searches to automatically complete in one run (default 10)
	- lines 73, 98, 119, 140
- Adjust the number of seconds between searches 
	- lines 54, 77, 102, 123, 145

The following functions were made: 
- printfullname(startrow) - prints the full name of the inputted row (prints in Python. Taken from columns 1 and 2 in Excel) 
- printphone(startrow) - prints the phone number of the inputted row (prints in Python. Taken from column 3 in Excel) 
- printemail(startrow) - prints the email address of the inputted row (prints in Python. Taken from column 4 in Excel) 
- printaddr(startrow) - prints the address of the inputted row (prints in Python. Taken from columns 5-8 in Excel) 
- customsearch(column, starting_row, num_searches)
- fullname(startrow)
- phone(startrow)
- email(startrow)
- addr(startrow)