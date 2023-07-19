# Auto_Google_from_Excel
Python code that automatically performs Google web searches. Takes the information from an Excel file, and automatically Google searches down the column.  

Use cases: lists that need web searching, such as names, addresses, phone numbers. 


## Dependencies 
- Excel file needs to be in .xls format only 
- Python needs packages selenium, xlrd, and time 
- Code utilizes ChromeDriver, which gets the Chrome web browswer that performs the automatic web searches. In my code, chromedriver.exe needs to be on the local machine and the path needs to be put into the code. Check out https://chromedriver.chromium.org/home to download ChromeDriver (about 12 megabytes). 

## Example_Excel_Workbook.xls
- This is an example Excel file that the Python code can run searches on
- columns: first_name, last_name, phone_number, email_address, street_address, city, state, zip_code
- rows: 10 examples, rows 2-11

## Python Code (auto_google_code.py)
I used Python 3.9

The following adjustments need to be made, according to your file paths and Excel file: 
- Update path to Excel document
	- line 17 
- Update path to ChromeDriver.exe   
	- lines 46, 69, 94, 115, and 136
- Adjust which sheet within Excel is being used
	- line 19 
- Adjust columns of Excel document 
	- lines 28, 31, 34, 37, 76, 101, 122, and 143
- Adjust the number of searches to automatically complete in one run (default 10 searches)
	- lines 73, 98, 119, 140
- Adjust the number of seconds between searches (default 3 seconds)
	- lines 54, 77, 102, 123, 145

The following functions are included: 
- printfullname(startrow) 
	- prints the full name of the inputted row (prints in Python, taken from columns 1 and 2 in Excel) 
- printphone(startrow) 
	- prints the phone number of the inputted row (prints in Python, taken from column 3 in Excel) 
- printemail(startrow) 
	- prints the email address of the inputted row (prints in Python, taken from column 4 in Excel) 
- printaddr(startrow) 
	- prints the address of the inputted row (prints in Python, taken from columns 5-8 in Excel) 

- fullname(startrow)
	- Automatically Google searches the contents of two Excel columns (first name and last name columns). Default searches through 10 rows, opening 10 tabs in Google Chrome. 
- phone(startrow)
	 - Automatically Google searches the contents of one Excel column (phone number). Default searches through 10 rows, opening 10 tabs in Google Chrome. 
- email(startrow)
	 - Automatically Google searches the contents of one Excel column (email address). Default searches through 10 rows, opening 10 tabs in Google Chrome. 
- addr(startrow) 
	- Automatically Google searches the contents of four Excel columns (street address, city, state, and zip code columns). Default searches through 10 rows, opening 10 tabs in Google Chrome. 

- customsearch(column, starting_row, num_searches)
	- The most customizable function. Allows you to set which column to search through, how many searches to run, and which row to start the searches on. 
	- For example, customsearch(10, 75, 20) allows you to run 20 automatic searches for the 10th column starting on row 75 and going through row 94. 