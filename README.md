# VBA Challenge: Stock Analysis Refactor
Data Bootcamp Module 2, September 19, 2022

# Project Overview
This analysis is intended to provide information about stock performance on an annual basis, specifically the total daily volume of trading for the year and the annual return rate. The original data set consists of the performance data for 11 "green" stocks for 2017 and 2018, presented in an [Excel file](https://github.com/larabjork/stock-analysis/blob/main/VBA_Challenge.xlsm). This analysis uses Visual Basic for Applications (VBA) to automate data analysis, using VBA macros (also called subroutines).

Raw data provided for this assignment consists of:
* Ticker abbreviation
* Date
* Opening price
* High price
* Low price
* Closing Price
* Adjusted closing price
* Daily trading volume

By following instructions presented in the course curriculum, the analyis first focused on one specific stock (ticker = DQ) in 2018 and . With further instructions from the curriculum, the analysis was then modified to include all 12 stocks and to allow the user to indicate whether to include data from 2017 or 2018. 

This independent challenge assignment built on the work completed with the course curriculum. The purpose of the challenge assignment is to refactor the existing code so that:
* the same output can be achieved, regardless of data set size (e.g., more stocks in a single year and/or more years of data)
* run time can be improved

# Results
A file was provided [(challenge_starter_code.vbs)](https://github.com/larabjork/stock-analysis/blob/main/challenge_starter_code.vbs) that contained a new version of the yearValueAnalysis subroutine (Module7 within the Excel file) with comments indicating what new code to insert. Those comments (numbered 1a through 4) are retained in the refactored subroutine, which is called AllStocksAnalysisRefactored (Module8).

## Refactored Code
The initial portion of the code remais the same (setting up a timer, creating a box for the user to input a year for analysis, creating header rows, initializing an array of all stock tickers, and finding the number of total number of rows that contain data). Both scripts also end with the same process for ending the timer.

In the original code, **For** loops iterate over variables **i** and **j**. In the refactored code, more specific names are used for iteration variables.

As shown in the following screenshots, in the refactored code, differences from the original code include the following:

A. In step 1a, the variable **tickerIndex** is initialzed before loops are initiated (in contrast with **ticker = tickers(i)**, within the parent For loop in the original code).

B. In step 1b, arrays are used for the total trading volume, starting price, and ending price, rather than variables.

C. In step 2a, an independent **For** loop is used to set initial value of tickerVolumes to zero; in the original code, this is done at the start of the parent **For** loop.

D. In step 2b, looping over all the rows in the spreadsheet is its own parent **For** loop, rather than being nested within another **For** loop as in the original code.

E. In steps 3a, 3b, and 3c, a similar pattern is followed as far as checking the ticker abbreviation in the first column of data against the current value of the ticker for daily volumne, start price, and end price. However, the refactored code uses the array element **tickers(tickerIndex)** rather than the variable **ticker** to do so. 

F. Also in step 3c, **tickerIndex** is advanced, which is accomplished by iterating through the parent For loop in the original code.

G. In step 4, an independent **For** loop is used to populate the **All Stocks Analysis** worksheet with the contents of the arrays (**tickers**, **tickerVolumes**)and the calulation of annual return based on the contents of **tickerEndingPrices** and **tickerStartingPrices**. In the original code, the output step is contained within the parent **For** loop.

H. The refactored code also contains formatting within this script. In the original code, the formatting is part of another subroutine (**formatAllStockAnalysisTable**, Module4), which is not pictured.

Refactored code, items A through F
![screenshot of VBA script of refactored code, with indications A-F to match discussion above](https://github.com/larabjork/stock-analysis/blob/main/Resources/Refactored_Code_Part_One.png)

Original code, items A through F
![screenshot of VBA script of original code, with indications A-F to match discussion above](https://github.com/larabjork/stock-analysis/blob/main/Resources/Original_Code_Part_One.png)

Refactored code, items G and H
![screenshot of VBA script of refactored code, with indications G and H to match discussion above](https://github.com/larabjork/stock-analysis/blob/main/Resources/Refactored_Code_Part_Two.png)

Original code, items G and H
![screenshot of VBA script of original code, with indications G and H to match discussion above](https://github.com/larabjork/stock-analysis/blob/main/Resources/Original_Code_Part_Two.png)


## Comparison of Code Output and Performance
To assess the impact of the refactor, I compared:

* output of data running the old subroutine (yearValueAnalysis) versus the new one (AllStocksAnalysisRefactored), for the same year (2018), to ensure data quality
* time (in seconds) required to run each subroutine, to assess the effect on performance

### Code Output and Data Quality
As shown in the screenshots below, the two subroutines produced the same data. Comparable results, although not pictured, were achieved for the 2017 data.

The formatting differs to ensure that the images are in fact the results of different macros. 

Output for 2018 data after running yearValueAnalysis macro:
![screenshot of Excel worksheet, showing All Stocks (2018); headings have larger font and are grayish blue](https://github.com/larabjork/stock-analysis/blob/main/Resources/Original_Sub_Results_2018.png)

Output for 2018 data from AllStocksAnalysisRefactored macro:
![screenshot of Excel worksheet, showing All Stocks (2018); headings have smaller font and are black](https://github.com/larabjork/stock-analysis/blob/main/Resources/VBA_Challenge_Results_2018.png)

### Execution Time and Performance
The refactored subroutine provided identical information more than five times faster than the original subroutine, as shown in the screenshots below (0.25 seconds versus 1.296875 seconds). Comparable results, although not pictured, were also achieved for the 2017 data.

Message after running yearValueAnalysis (original) macro:
![screenshot of Excel alert, stating "This code ran in 1.296875 seconds for the year 2018"](https://github.com/larabjork/stock-analysis/blob/main/Resources/VBA_Challenge_2018_original.png)

Message after running  AllStocksAnalysisRefactored (refactored) macro:
![screenshot of Excel alert, stating "This code ran in 0.25 seconds for the year 2018"](https://github.com/larabjork/stock-analysis/blob/main/Resources/VBA_Challenge_2018_refactor.png)

# Summary
## General Advantages and Disadvantages of Refactoring Code
Refactoring code is generally a good idea as long as it does not change the output quality, offers faster performance, and/or makes the code easier to read. Refactoring can go too far, especially if code readability is not prioritized. For instance, code using single-letter variable names is shorter but not necessarily clearer.

Refactoring code can have unintended consequences. Without a way to check to see that the same output is maintained, errors could be introduced that might be visible to an end user. Other code could fail based on refactored code, and if there is not a suite of tests (ideally automated tests) available, the source of failure may be hard to track down.

## Specific Advantages and Disadvantages of Refactoring This VBA Script
The increased performance speed is a clear advantage of the refactored script. The code is also easier to read, because there are no nested loops.

The code currently does not have an automated way to check that the output of the two scripts are identical. That check has to be completed manually, which is a source of potential errors.

The refactored code is also based on knowing that 12 stocks will be analyzed. To be useful for a data set of a different size (presumably larger), further refactoring would be needed:
* The sizes of all arrays would have to hard coded to equal the new total number of stocks; or 
* Code that counted the number of unique values for stock ticker abbreviations would have to be added, with that count then used to reset the sizes of all arrays.

