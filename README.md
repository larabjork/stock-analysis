# VBA Challenge: Stock Analysis Refactor
Data Bootcamp Module 2, September 19, 2022

# Project Overview
This analysis is intended to provide information about stock performance on an annual basis, specifically the total daily volume of trading for the year and the annual return rate. The original data set consists of the performance data for 11 "green" stocks for 2017 and 2018, presented in an Excel file, available here (INSERT FILE PATH HERE). This analysis uses Visual Basic for Applications (VBA) to automate data analysis, using VBA macros (also called subroutines).

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
A file was provided [challenge_starter_code.vbs)](https://github.com/larabjork/stock-analysis/blob/main/challenge_starter_code.vbs) that contained a new version of the yearValueAnalysis subroutine (Module7) with comments indicating what new code to insert. Those comments (numbered 1a through 4) are retained in the refactored subroutine, which is called AllStocksAnalysisRefactored (Module8).

## Refactored Code

SUMMARY OF REFACTORING HERE

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
The refactored subroutine provided identical information more than five times faster than the original subroutine, as shown in the screenshots below (0.25 seconds versus 1.375 seconds). Comparable results, although not pictured, were also achieved for the 2017 data.

Message after running yearValueAnalysis (original) macro:
![screenshot of Excel alert, stating "This code ran in 1.375 seconds for the year 2018"](lhttps://github.com/larabjork/stock-analysis/blob/main/Resources/VBA_Challenge_2018_original.png)

Message after running  AllStocksAnalysisRefactored (refactored) macro:
![screenshot of Excel alert, stating "This code ran in 0.25 seconds for the year 2018"](https://github.com/larabjork/stock-analysis/blob/main/Resources/VBA_Challenge_2018_refactor.png)

# Summary
## General Advantages and Disadvantages of Refactoring Code

## Specific Advantages and Disadvantages of Refactoring This VBA Script
