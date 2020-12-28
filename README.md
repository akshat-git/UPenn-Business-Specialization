# Business-Machines
Business Machines

First you must enable the google sheets and google drive api from the google cloud console, and take the credentials(OAuth2). There are videos about how to enable these and start off with them.
You must then put these files in your working directory, and put the names into the 

the drive_test_api file has the main program. run create.py to begin, and then pick up ids from the resulting spreadsheet and folder.
use these ids by putting them into the respective variable in the ids file
when all the ids have been placed, you can use the drive_test_api code which takes a file from the hist price folder and gets summary statistics.
if you want to add more stocks to the comparison, you can add them to the hist price folder, and add a symbol(the ticker) to the dictionary at the top.
also add more value to the for loop of function calls, as to iterate through the symbol you would like.

This python code uses the google sheets and google drive api, as well as json for styles and formatting. use the work days variable to specify how many open days of the market you want to consider into your comparison.
the data for the 10 given stocks were updated on 12/24/2020
