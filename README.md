# Stock Analysis and VBA Code Refactor

## Overview of Project

This project uses VBA to analyze the performance of twelve stocks during years 2017 and 2018. An input box allows a user to enter the year they wish to analyze and then, with a click of a button, the total daily volume and yearly return for each stock is presented and formatted.

### Purpose

The purpose of this challenge is to improve upon the code originally written during the module coursework.  The original macro utilized nested for loops and re-used the same variables for each of the twelve stocks.  The final results were accurate, however the solution may not scale for thousands of stocks.  

This challenge refactored the code to create an array of variables for each stock thus bypassing the need for a nested for loop.

Timers were added to measure elapsed time and code performance.


## Results

The analysis showed that the twelve stocks were considerably more successful in 2017 than 2018. Only one stock, TERP, had a negative return in 2017. In contrast, only two stocks, ENPH and RUN, had positive returns in 2018.  

The refactored code also showed performance gains.  It took roughly .30 seconds to run the original macro for either year 2017 or 2018.  In comparison, the timing results for the refactored VBA is almost 4 times faster at an average of .08 seconds.

The elapsed time of the original macro - roughly .30 seconds:

![Screenshot of timer - not refactored](/Resources/Original_Timing.png)

Compared the the elapsed time for the refactored VBA:

![Screenshot of timer - year 2017](/Resources/VBA_Challenge_2017.png)

![Screenshot of timer - year 2018](/Resources/VBA_Challenge_2018.png)

## Summary

The refactored code was noticeably faster than the original VBA macro.  A chief contributor to the elapsed time in the original macro is the nested for loop.

You can see in this code that the entire row count of the worksheet is looped over 12 times.

'''
   'Outer loop through tickers    For i = 0 To 11        ticker = tickers(i)        totalVolume = 0        Worksheets(yearValue).Activate                'Inner loop through each row for each ticker        For j = 2 To RowCount

            'get total volume for current ticker            If Cells(j, 1).Value = ticker Then                totalVolume = totalVolume + Cells(j, 8).Value            End If                        'get starting price for current ticker            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then                startingPrice = Cells(j, 6).Value            End If                        'get ending price for current ticker               If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then                endingPrice = Cells(j, 6).Value            End If                   Next j        
'''

In comparison, the refactored macro used an array to store the values for each ticker, thus allowing the loop to only run once through all the rows in the worksheet

If more stocks are added, the number of times that the rowCount is looped over would continue to increase in the original macro thus slowing elapsed time even more.  The refactored code would handle an increase in stocks over time with less performance degradation.

Refactoring code is beneficial in that it improves on the design, performance and readability of the software.  Refactoring can be risky however, if the code isn't well documented or if there are insufficient test cases and functional requirements. For this exercise there were no obvious disadvantages in refactoring the code because the script was well commented and the expected results were clear and understood.


