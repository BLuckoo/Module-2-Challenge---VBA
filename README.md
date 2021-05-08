# **Module 2 Challenge - VBA**
## Using VBA to analyze stocks dataset to determine returns  - the pros and cons of refactoring of code is also discussed and illustrated in this challenge.

## **1. Overview of Project**
### This project is made up of three main parts:
    a) Using the 2017 and 2018 stock datasets provided to calculate the percentage returns of a list of stocks 
    b) Demonstrate the use of VBA code to format worksheets, to create Message Boxes, the use of Input Boxes and Display Boxes, use of arrays, loops and ifs statements in VBA  code
    c) Discuss the pros and cons of refactoring in coding by using illustrations from this project
##
## **2. Analysis of Stocks for 2017 and 2018**
     
     Using VBA code, the 2017 & 2018 stocks datasets were analyzed in the following manner:
     a) Use of an Input Message Box to allow the client to select for which year he wants to do the analysis 
     b) The VBA code uses the following concepts to collect data, such as ticker symbol, starting price and ending price each stock ticker symbol 

        Arrays - used to collect the data about each stock
        Loops - used to iterate through every single row of the dataset 
        Ifs statements - to validate and send specific data meeting certain conditions to the array for subsequent reporting


### **The Excel file VBA_Challenge is made up of a number of worksheets:**
###
       Sheets 2017 and 2018 contain the datasets for the years 2017 & 2018 which will be analyzed and reported upon using the VBA codes in subroutines:
       **YearValueAnalysis** - VBA code prior to refactoring and displaying the results in worksheet *"All Stocks Analysis b4 Refactor"*
       **AllStocksAnalysisRefactored** - VBA code following refactoring using indexed arrays and displaying the results in worksheet *"All Stocks Analysis"*
##
##
## **Output of "YearValueAnalysis"**
##
This subroutine uses arrays to keep the information about the stocks' ticker symbols and then using "loops" and "ifs" iterates through the data set to accumulate the volumes traded and identify the starting and ending prices of each ticket symbol. A formula is then used to calculate the percentage return of each ticker symbol.
###
#### The following are the results for the 2017 Dataset prior to the refactoring exercise: 
#### https://github.com/BLuckoo/Module-2-Challenge---VBA/blob/main/VBA_Challenge_2017.PNG
####
     The results show that for the year 2017 most of the stocks had positive returns with the top performing stocks being DQ, SEDG and EWPH. 
     The TERP stock was the only non-performing stock with a negative return of 7.2%.
###
###
###  **The code for this subroutine took 1.046875 seconds to run.** 
### 
####  The following are the results for the 2018 Dataset prior to the refactoring exercise: 
####  https://github.com/BLuckoo/Module-2-Challenge---VBA/blob/main/VBA_Challenge_2018.PNG
      Unlike 2017, in 2018 most of the stocks had negative returns. Only ENPH and RUN had positive returns of 81.9% and 84% respectively.
###
###    **The code for this subroutine ran in 1.0625 seconds.**
###
####    The following file contains the code for this analysis: 
####    https://github.com/BLuckoo/Module-2-Challenge---VBA/blob/main/SUB%20YearValueAnalysis.txt
###
##
## **Output of "AllStocksAnalysisRefactored"**
##
    This subroutine uses the same coding concepts that were used in the subroutine "YearValueAnalysis" above except that the concept of indexing of the array is being introduced to help the code run a little bit faster.
##
### The results of the subroutine after refactoring are the same for both years, i.e. the returns for both 2017 and 2018, are exactly the same as with the prior subroutine "YearValueAnalysis". 
###  
####    The following files show the output for the year 2017 and the 2018 after the refactoring is done:
####    https://github.com/BLuckoo/Module-2-Challenge---VBA/blob/main/VBA_Challenge_2017_after_refactor.PNG
####    https://github.com/BLuckoo/Module-2-Challenge---VBA/blob/main/VBA_Challenge_2018_after_refactor.PNG
###
### For the 2017 dataset, the run time in seconds was 1.039063 compared to 1.046875 prior to refactoring.
###
### For the 2018 dataset, the run time in seconds was 1.054688 compared to 1.0625 prior to refactoring.
###
####    The following file contains the code for this analysis: 
####    https://github.com/BLuckoo/Module-2-Challenge---VBA/blob/main/SUB%20AllStocksAnalysisRefactored.txt
###
###
##  **3. The Pros and Cons of Refactoring**
### 
### Pros in general: 
        Making the code more efficient
        Helps in finding bugs
        Improvement in design of software
        Makes it easier to maintain over time
###
### Cons in general:
        Can be costly
        Time consuming








