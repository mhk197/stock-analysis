# VBA Stock Analysis

## Overview of Project

### Purpose
This project entailed refactoring a VBA Macro that automates key components of equity analysis. In particular, the Macro loops through Excel datasets that record a year of daily prices and volumes for several different stocks, calculates each stock's total daily volume and yearly return, and outputs the formatted results into a separate spreadsheet. The refactored code was designed to be faster and more flexible than the original code.

## Results

### Differences between the Original Code and the Refactored Code

#### Original Code
The original code implemented a nested For loop to iterate through the dataset. The outer loop iterated through an array of stock tickers, while the inner For loop iterated through all of the rows in the dataset and calculated the total daily volume, starting price, and ending price associated with the current ticker. The program held these values in single-value variables. After each iteration of the inner loop, the outer loop outputted the current ticker and its associated total daily volume and return on a designated output spreadsheet. Every time the outer loop iterated through the dataset, the total daily volume variable was reset to 0 and the starting price and ending price variables were updated with the values associated with the new ticker.

![Original_Code](https://user-images.githubusercontent.com/87445739/129258042-dcab29d4-93ef-4410-abcb-00a13033ba7d.png)

*The original code therefore looped through the dataset 12 times -- one for each ticker. Additionally, it switched between the spreadsheet holding stock data and the output spreadsheet with each loop.*

#### Refactored Code
The refactored code only implemented one For loop to iterate through the dataset. It held the tickers, total daily volumes, starting prices, and ending prices in separate arrays, and associated the corresponding values with the same index. A variable representing this index, tickerIndex, was initialized to 0 at the start of the loop. As the program progressed, it stored the total daily volumes, starting prices, and ending prices associated with the tickerIndex in their respective arrays. Finally, tickerIndex increased by 1 after the last row associated with the current ticker. The implementation of arrays instead of singular variables removed the need to iterate through the dataset for each ticker.

![Refactored_Code_Main](https://user-images.githubusercontent.com/87445739/129258999-9148a8d4-6d69-4214-85a5-5d455fe79254.png)

After collecting the relevant values from the dataset, the program then activated the separate output sheet and calculated and outputted the results of the analysis by iterating through each array with the same For loop.

![Refactored_Code_Formatting](https://user-images.githubusercontent.com/87445739/129258849-69486323-3195-4aa3-81e0-70f8e61f6703.png)

*The refactored code therefore looped through the dataset 1 time. It switched to the output spreadsheet only once, after this main loop.*

#### Additional Differences
The original program did not format the output sheet -- that task was left to another Macro. The refactored program includes the code from that Macro, so analysis and formatting can be accomplished by one program.

### Efficiency of the Refactored Code
While there is some variation in the runtime every time a Macro is run, the refactored code was clearly much faster than the original code. 

The refactored code ran in approximately 0.164 seconds for the 2017 dataset and approximately 0.168 seconds for the 2018 dataset. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/87445739/129258387-13370b10-76bc-49bc-849d-b23519bcdf4e.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/87445739/129258401-4e1cc6bc-89ab-4675-b47f-28bb9ba3d573.png)

In comparison, the original code ran in approximately 0.828 seconds for the 2017 dataset and 0.883 seconds for the 2018 dataset.
![Original_Code_Runtime_2017](https://user-images.githubusercontent.com/87445739/129258472-d5516ebb-941e-4a76-b353-da150ed4e777.png)
![Original_Code_Runtime_2018](https://user-images.githubusercontent.com/87445739/129258500-52215674-6442-4287-91f4-7c2bc3ca5995.png)

## Summary

### Refactoring 
