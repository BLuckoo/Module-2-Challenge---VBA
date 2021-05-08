# **Module 2 Challenge - VBA**
## Using VBA to analyze stocks dataset to determine returns  - the pros and cons of refactoring of code is also discussed and illustrated in this challenge.

## **1. Overview of Project**
### This project is made up of three main parts:
### a) Using the 2017 and 2018 stock datasets provided to calculate the percentage returns of a list of stocks 
### b) Demonstrate the use of VBA code to format worksheets, to create Message Boxes, the use of Input Boxes and Display Boxes, use of arrays, loops and ifs statements in VBA code
### c) Discuss the pros and cons of refactoring in coding by using illustrations from this project
##
## **2. Analysis of Stocks for 2017 and 2018**
### Using VBA code, the 2017 & 2018 stocks dataset was analyzed in the following manner:
### a) Use of an Input Message Box to allow the client to select for which year he wants to do the analysis 
### b) The VBA code uses the following concepts to collect data, such as ticker symbol, starting price and ending price each stock ticker symbol 

####    Arrays - used to collect the data about each stock
####    Loops - used to iterate through every single row of the dataset 
####    Ifs statements - to validate and send specific data meeting certain conditions to the array for subsequent reporting


### **The Excel file VBA_Challenge is made up of a number of worksheets:**
###
### Sheets 2017 and 2018 contain the datasets for the years 2017 & 2018 which will be analyzed and reported upon using the VBA codes in subroutines:
####   **YearValueAnalysis** - VBA code prior to refactoring and displaying the results in worksheet *"All Stocks Analysis b4 Refactor"*
####   **AllStocksAnalysisRefactored** - VBA code following refactoring using indexed arrays and displaying the results in worksheet *"All Stocks Analysis"*

##
### **Output of "YearValueAnalysis"**
###     This subroutine uses arrays to keep the information about the stocks' ticker symbols and then using "loops" and "ifs" iterates through the data set to accumulate the             volumes traded and identify the starting and ending prices of each ticket symbol. A formula is then used to calculate the percentage return of each ticker symbol.
###
### The following are the reults for the 2017 Dataset prior to the refactoring exercise:
###
### 





