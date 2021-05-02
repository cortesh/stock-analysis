# stock-analysis

## Overview

### The purpose of this project is to help develop an app to analyze 12 stocks across two years to identify trends with limited capabilities to select year and reset the results page once completed.  

### A second requirement of the project was to develop two methods (method A and B) to test speed and scaleability of the app.

## Results
### The first step was to test the performance of a particular stock (ticker "DQ") against other stocks in the portfolio.  This was done by calculating the Return as the **ending price / starting price - 1**.  When we looked at the results for 2018, it is clear that not only did DQ perform poorly for that year, but that all but two stocks ended the year down.  This data was displayed with a table that listed the stocks by *ticker symbol*, *Total Daily Volume* and *Returns*, conditionally formatted to turn the cell fill green when in positive territory and red when below.  This image is located in the Resources folder as per project requirements.  I have converted that table into a bar chart that I believe will make those relative results more impactful and immediate (see below):

![2018 Returns](https://github.com/cortesh/stock-analysis/blob/main/Resources/VBA_Challange_2018_returns.png)

### It is important to note that before concluding that "DQ" was not an outlier in its poor performance when compared to the other 9 stocks also in negative territory for this period.  What is more compelling still is that a comparison to the previous year ("2017") reveals that  all but one stock was in positive territory, with "DQ" ranked in 1st place! (see below)

![2017 Returns](https://github.com/cortesh/stock-analysis/blob/main/Resources/VBA_Challange_2017_returns.png)

### So this study is inconclusive with respect to the specific stock "DQ" in relation to other stocks once both years are considered, and that the pattern that emerges is that the greatest driver for success or failure seems to be the overall market conditions in each of the years of the study.

### Continuing to the second major requirement, the project sought to compare two methods of coding to determine processing speed as a proxy for scaleability (ability to handle larger datasets without performance loss).  The first method ("A") build a code with 2 nested for loops that read through all records in a given table (see below)


   
   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = 'ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = 'ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j






![Code Performance Comparison](https://github.com/cortesh/stock-analysis/blob/main/Resources/VBA_Challange_Code_Performance_Comparison.png)

## Summary
### para 1
### para 2
### para 3


