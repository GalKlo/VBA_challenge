# Stocks Summary

This script loops through all the stocks for one year and outputs the following information in a form of two summary tables.

Table 1. Returns the information about each Ticker, including:
* Ticker name;
* Its yearly change (from the opening price at the beginning of a given year to the closing price at the end of that year);
* Percentage change (from the opening price at the beginning of a given year to the closing price at the end of that year);
* Total stock volume of the stock (sum of all the volume throught out a given year).

Table 2. Summarizes the tickers with:
* Greatest % increase;
* Greatest % decrease;
* Greatest Total Volume.

Script runs on multiple sheets in a workbook.


## Improvement opportunities

The placement of the two summary tables on a worksheet is hardcoded, in case the structure of the exported files changes the code will require dynamic positioning of the summary tables depending on the number of the last column in a raw data table.


### References 

Solution that allows to manipulate data format in the column <date> by changing it from text to numbers (lines 58-63 in the VBA script) was taken from: https://stackoverflow.com/questions/36771458/vba-convert-text-to-numbersort.