# Property_Data_Analysis

This script is intented to go to automatically go to realtor.com search for a city you put in the city_state.txt and pull the addresses, prices, bedroom, and bathroom information. Use that info to determine mortgage (including average state property tax and state homeowners insurance), finds average city rents based on bed and bath size, calculated monthly and annual cash flow, and ROI. Then builds and inputs that into an excel spreadsheet.

Step 1: 

Open Powershell as an admin and run Install-Module -Name PSWriteExcel -RequiredVersion 0.1.1

Link to module for reference:
https://www.powershellgallery.com/packages/PSWriteExcel/0.1.1

Step 2: 
Open city_state.txt file and input desired city and state (use the examples for formatting) and save when done

******THINGS TO KNOW ABOUT THE SCRIPT******
It uses realtor.com to gather info based on SFH with 2+ beds/1+ baths between $10K and $250K. If you want to change that open the script in Powershell ISE. You can find the URL specifying the search criteria on line 37 and 77.

Starting on line 149 there is a mortgage calculator. Line 151 has $ask which is an asking price times 80% because that is what i use as my initial offer. Change that at your discretion. Line 152 has $down as the down payment which is 20% of $ask. Again, change that at your discretion. You will also need to change them on line 252 and 253.

Starting at line 208 is the monthly expense calculations. I use 5% for vacancy, 5% for repairs, 10% CapEx, and 10% for property managers. You can take out any of these but, if you do, make sure you take it out of 213, also.

Once it runs through the first loop (if you have multiple cities in the txt file) it will auto create an excel spreadsheet in the folder. You can open the spreadsheet when the progress bar saying "Taking a break before the next city" is ticking down. If you only have one city and this progress bar starts just ctrl+c the script.
