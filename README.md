# Stock Analysis and VBA Code Refactor

## Overview of Project

This project uses VBA to analyze the performance of twelve stocks during years 2017 and 2018. An input box allows a user to enter the year they wish to analyze and then, with a click of a button, the total daily volume and yearly return for each stock is presented and formatted.

### Purpose

The purpose of this challenge is to improve upon the code originally written during the module coursework.  The original macro utilized nested for loops and re-used the same variables for each of the twelve stocks.  The final results were accurate, however the solution may not scale for thousands of stocks.  

This challenge refactored the code to create an array of variables for each stock thus bypassing the need for a nested for loop.

Timers were added to measure elapsed time and code performance.


## Results

The analysis should that the twelve stocks were considerably more successful in 2017 than 2018. 

The refactored code also showed performance gains.  It took roughly .30 seconds to run the original macro for either year 2017 or 2018.  In comparison, the timing results for the refactored VBA is almost 4 times faster.

![Screenshot of timer - year 2017](/Resources/VBA_Challenge_2017.png)

![Screenshot of timer - year 2018](/Resources/VBA_Challenge_2018.png)



