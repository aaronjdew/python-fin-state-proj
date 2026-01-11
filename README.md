# Financial Statement Business Calendar Data Creation

This project updates an existing .xlsx (with formulas) used for financial statements for a class based fitness business.\
The .xlsx is used to forecast usage of the available 8 fitness studio rooms vs the available hours.\
This project utilises Pythons pandas and openpyxl to remove, create and insert 3 business years of calendar data for business financials.\
Then daily and monthly calendar data is produced for each available room in the business

## Key Considerations

All financial data sheets, P&L, Balance Sheet, Sales forecast etc.. have all been removed from the example dataset.\
This is purely the code written for creation of calendar dates, and the aggregated daily/monthly dates, split between rooms, for 3 business years.\
All other code and data for the other parts of the project... sales, expenses, cash flow code have been removed

## Python Packages

The following was used to create this project: \
(For a full list of installed packages and versions check the 'requirements.txt' file in this project)

- Python 3.14.2
- Main Python Packages
  - pandas
  - numpy
  - openpyxl
