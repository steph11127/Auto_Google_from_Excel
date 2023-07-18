# Auto_Google_from_Excel
Python code that automatically performs Google web searches. Takes the information from an Excel file, and automatically Google searches down the column.  
Use cases: lists that need web searching, such as full names, addresses, phone numbers. 


## Dependencies 
- Excel file needs to be in .xls format only 
- Python needs packages selenium, xlrd, and time 
- Code utilizes ChromeDriver, which gets the Chrome web browswer that performs the automatic web searches. In my code, chromedriver.exe needs to be on the local machine and the path needs to be put into the code. Check out https://chromedriver.chromium.org/home to download ChromeDriver (about 12 megabytes). 

## Example_Excel_Workbook.xls
columns: first_name, last_name, phone_number, email_address, street_address, city, state, zip_code
rows: 10 examples, rows 2-11

## Python
I use Python 3.9
