# Stock-analysis
Module 2 Respository on Stock Analysis

Added macro enabled spreadsheet with first stock analysis completed. The macros include multiple Nested Loops. The program must run through 3000 rows of data multiple times because of the nested loops. If the rows were to be increased with more stocks. the computing time might increase significantly due to the NESTED Loops. The Challenge 2 was to rewrite the code to improve efficiency. 

A second excel file labeled Challenge_2_green_stocks_final was created with completely rewritten code. See comments below under Challenge 2

PseudoCode For Modification of Macro to AllStockAnalysis

Sub AllStocksAnalysis()
   '1) Format the output sheet on All Stocks Analysis worksheet

   '2) Initialize array of all tickers
   '3a) Initialize variables for starting price and ending price

   '3b) Activate data worksheet
   '3c) Get the number of rows to loop over

   '4) Loop through tickers
   For i = 0 to 11
       ticker = tickers(i)
       '5) loop through rows in the data
       For j = 2 to RowCount
           '5a) Get total volume for current ticker
           '5b) get starting price for current ticker
           '5c) get ending price for current ticker

       Next j
       '6) Output data for current ticker

   Next i

End Sub
# DQAnalysis 
This Macro remains the same from the module lessons. It performs analsysis for a single stock. 

# ClearWorksheet
This Macro clears all cells in the worksheet. A run Button was added to initiate as well as a Run Button to run the AllStocksAnalysis. 

# Challenge 2 

In rewriting the code, I took the following approach:
* Created a New subroutine called "NewAllStocksAnalysis."

This new subroutine takes the following approach:
1) Obtain User Input on Year
2) Activate AllStocksAnalysis worksheet
3) Formats the headers and data type columns.
4) Dimensions all 4 arrays. 
5) For the Tickers Array, I did not preload it with stock symbols. Instead the Ticker Array is read in values from the data worksheets and checks to make sure there are no duplicates. 
6) RowCount is used to determine the number of loops to loop once through the entire spreadsheet of data. 
7) The loop started using the RowCount along with a itckerIndex variable to keep all results in sync. 
8) The conditional statements were revised to account for a single pass through the Loop. 
9) Conditionals for Ticker Change was added followed by immediate code to calculate totalVolume, startingPrice, and endingPrice - which are all stored in their respective arrays immediately. 
10) Two new columns were added to the Results worksheet for startingPrice and endingPrice to improve debugging of code and proof of valid results. 
11) The Return Column is calculated on the fly by pulling the data from the startingPrice and endingPrice arrays during the same Output Loop to print Tickers, and totalVolume. 
12) In this way, only a single pass through each of the 3013 rows has to be made - thus running more efficiently. 

After all values have been output and shown on the summary sheet AllStockAnalysis", conditional color formatting is applied to the Return Results. 
