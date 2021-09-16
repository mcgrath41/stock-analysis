# stock-analysis

## Overview of Project

Use VBA to automate a high-level analysis of the performance of a doezen stocks with code that can also run efficiently with larger datasets (i.e. entire stock market).

## Results

### 2017 vs. 2018 Performance

Using `for` loops and `if-then` statements, scripts were created to automatically calculate the Total Daily Volume ($) and Return (%) for the list of stocks, with an `InputBox` included to allow the user to input the desired year to measure. The results suggest that 2017 was a much better time to invest in these stocks overall; however, ENPH had positive returns in both years and RUN performed better in 2018.  

![](/Resources/VBA_Challenge_2017_Orig_Code.png)

![](/Resources/VBA_Challenge_2018_Orig_Code.png)

### Original vs. Refactored Script

The code was refactored to provide a more efficient script that is capable of running larger datasets with minimal loss of productivity. The two screen shots above were run with the original code with corresponding time stamps, which you can see were both around 1 second. The refactored code was used to generate the same tables for 2017 & 2018, and as you can see in the screen shots below, their time stamps are significantly improved.

![](/Resources/VBA_Challenge_2017.png)

![](/Resources/VBA_Challenge_2018.png)

## Summary

### What are the advantages or disadvantages of refactoring code?

Refactoring can be advantageous by making the code easier to understand and maintan (for future enhancement) and by removing bad smell (duplicate code, long method, etc.) to prevent future defects; however, the disadvantages of refactoring include cost, time, and potentially introducing bugs.

### How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original VBA script resulted in a more sophisticated code that improved the chances of future enhancement. The additional time required was minimal, and the matchcing results indicate that bugs were not introduced.
