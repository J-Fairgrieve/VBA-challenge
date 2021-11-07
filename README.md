# VBA Challenge
### A homework assignment from the University of Birmingham Data Analytics Bootcamp (November 2021)

 - The aim of this project was to create a VBA script to analyse real stock market data
 - A test data file was used to develop the script before running on a larger file.
 - Loops, conditional formatting & worksheet functions are used in the script to produce the final tables

### Results
[The script](https://github.com/J-Fairgrieve/VBA-challenge/blob/main/WallStreetVBA.vb) automatically creates two summary tables on each worksheet of the file:

#### **Summary Table 1: Individual Ticker Metrics**
Summarises the Ticker's individual performance on the sheet, highlighting:
 - Yearly Change in Stock Price *(Final Close - First Open)*
 - % Change
 - Total Stock Volume

#### **Summary Table 2: Grouped Ticker Metrics**
Provides further analysis by providing the following metrics from the first summary table:
 - Greatest % Increase
 - Greatest % Decrease
 - Greatest Total Volume

Results are as follows:
![2016 Summary](https://raw.githubusercontent.com/J-Fairgrieve/VBA-challenge/main/2016%20Data.png)
