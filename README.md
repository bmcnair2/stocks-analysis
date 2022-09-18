# Stock Analysis

## Overview of Analysis

This Excel workbook contains stock data from 2017 and 2018 on green stocks. The purpose of the analysis was to use VBA to analyze green stocks for the financial advisor Steve to present to his parents. Steve's parents are trying to diversify their portfolio with more green stocks.

## Results

The VBA analysis produced results for Total Daily Volume and Yearly Return for each year. The worksheet's formatting was edited to create better visualization and buttons were added for automatic analysis.

The analysis showed some slight difference between the factored and unfactored code.

![2017 Unfactored Code](../Module%202/resources/time%20before%20refactor%202017.png)

![2018 Unfactored Code](../Module%202/resources/time%20before%20refactor%202018.png)

![2017 Factored Code](../Module%202/resources/time%20after%20refactor%202017.png)

![2018 Factored Code](../Module%202/resources/time%20after%20refactor%202018.png)

![2017 Return Percentages](../Module%202/resources/2017%20analysis.png)

![2018 Return Percentages](../Module%202/resources/2018%20analysis.png)

As we can see, 2017 was a very good year for green stocks while 2018 was a down year. In 2017, only one stock ($TERP) was down on the year, while in 2018 only two stocks ($ENPH and $RUN) had positive performance. If Steve's parents were looking for a buying opportunity, 2018 was the year for buying the dips.

## Summary

The definition of refactoring is editing existing code to see if it will run faster. The editing is done in small steps in order for it to run correctly, and there are no major changes to the existing code. We see from this analysis that the original code actually ran faster than the refactored code. Some advantages to the refactored code are that it may be easier to read than the original version and it's considered safe since testing had to occur to make sure the code ran properly. Some disadvantages are since it took longer to run the refactored code it takes up more space on the computer and some parts of the code may be repeated.
