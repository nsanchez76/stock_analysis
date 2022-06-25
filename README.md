# VBA Challenge

## Overview of Project
An analysis was performed on a set of data containing information about 12 different stocks for the years 2017 and 2018. It was determined that only one stock from 2017 had a negative return. It was also calculated that only 2 stocks in 2018 had a positive return.

### Purpose
The purpose of this analysis was to provide the client with an analysis of the performance of 12 different stocks from the years 2017 and 2018.

## Results

Microsoft Excel was used to analyze the performance of 12 different stocks during the years 2017 and 2018. A macro was written to create a table that lists each stock's ticker symbol, its total daily volume and its annual return. Thousands of lines of data were analyzed by the macro. Multiple data arrays were created to hold the information that the macro calculated. For loops were used to systematically run through each row of input data, as well as run through each ticker symbol. If statements were used to determine the first row of a ticker's set of data. The starting price of that stock would then be extracted and saved in the "tickerStartingPrices" array. Another If statement was used to find the last row of data for each stock. The closing price for that stock would be extracted and saved in the "tickerEndingPrices" array. Another For loop was used to display the values in each array. Finally, the start time and end time of each run of the macro was calculated and displayed in a pop-up window.

The resulting table for 2017 is this:

![2017 Table](https://user-images.githubusercontent.com/106849689/175757872-9ed5d748-e649-416d-b41b-5b37de545d75.png)


The resulting table for 2018 is this:

![2018 Table](https://user-images.githubusercontent.com/106849689/175757886-3c92b7e9-5e16-4b0c-8c52-b1b21392ed04.png)


The resulting pop-up window for 2017 is this:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106849689/175757894-e7285d5c-e7fd-45ca-9ed8-d2de1ff9788c.png)


The resulting pop-up window for 2018 is this:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/106849689/175757896-588c1bb0-87e4-4c07-81ae-c277d7c2c2bf.png)


## Summary

### Advantages/Disadvantages of refactoring code
A big advantage of refactoring code is that the basic function of the code is already written. A disadvantage would be that the original code could have been written in a very sloppy way, or it could have been written in a different order than how someone else would have written it. This is where comments are extremely important. Comments provide anyone reading the code an explanation of what each bit of code is supposed to do. This way the reader can better follow the coder's logic.


### Advantages/Disadvantages of the original and refactored VBA script
An advantage of the original script was that it is very simple. A bid disadvantage is that is used a lot of "magic numbers", meaning that the coder went into the data and determined that certain numbers needed to be used in certain formulas, so those "magic number" were hard-coded. The refactored script does not have "magic numbers", which makes it more portable, and can be used on many different years of data of these 12 stocks. A disadvantage, is that is only works on this specific set of 12 stocks.
