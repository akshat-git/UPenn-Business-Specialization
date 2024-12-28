# Business-Machines-for-app-and-commercial-use
# Uses-Google-Drive-&-Sheets-API-and-python-script-json

This git repo turns ideas and projects used in the UPenn Business Specialization from Coursera, and it uses google sheets API's to integrate it into a python script that updates a new google sheet for the user. This makes my code and the sheets I created to assist myself in the specialization more accessible and replicable on websites or commercial use.

[https://coursera.org/share/4ec21da8a22cf0d0fd5e021b79c9a7be](url)
- Certificate of completion of the 5-course specialization on the following topics:
- Course Certificates Completed
  - Decision-Making and Scenarios
  - Wharton Business and Financial Modeling Capstone
  - Introduction to Spreadsheets and Models
  - Fundamentals of Quantitative Modeling
  - Modeling Risk and Realities

Wharton's Business and Financial Modeling Specialization is designed to help you make informed business and financial decisions. These foundational courses will introduce you to spreadsheet models, modeling techniques, and common applications for investment analysis, company valuation, forecasting, and more. When you complete the Specialization, you'll be ready to use your own data to describe realities, build scenarios, and predict performance. 


Instructions for setup:

First you must enable the google sheets and google drive api from the google cloud console, and take the credentials(OAuth2). There are videos about how to enable these and start off with them.
You must then put these files in your working directory, and put the names into the ids file under the name credentials for each

the drive_test_api file has the main program. run create.py to begin, and then pick up ids from the resulting spreadsheet and folder.
use these ids by putting them into the respective variable in the ids file
when all the ids have been placed, you can use the drive_test_api code which takes a file from the hist price folder and gets summary statistics.
if you want to add more stocks to the comparison, you can add them to the hist price folder, and add a symbol(the ticker) to the dictionary at the top.
also add more value to the for loop of function calls, as to iterate through the symbol you would like.

This python code uses the google sheets and google drive api, as well as json for styles and formatting. use the work days variable to specify how many open days of the market you want to consider into your comparison.
the data for the 10 given stocks were updated on 12/24/2020
