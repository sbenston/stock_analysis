# Analysis of Green Energy Stocks

## Overview of Project

This project uses Microsoft Excel with macros enabled to automatically populate a worksheet that reports the yearly perfomance of stocks within the clean energy sector. Automating the task of reporting simple analyses such as the yearly return of stocks showcases the power of macros written in VBA for Microsoft Excel; simply by inputting the year a user can instantly generate the most useful data to make judgments when investing. This valuable feature in Excel is combined with the concept of evalutating code performance to ensure that even when working with larger datasets, the macro will run quickly and reliably.

## Results

### Analysis and Results of Stock Performance

For analyzing yearly stock data two values were calculated for each stock ticker: the Total Daily Volume and the yearly Return. A macro was created that retrieved values from data sheets via iterating through each row on the sheet and using the ticker column pulled the needed data into variables created to hold the results. In order to find when the starting point of one stock's data began, a conditional statement was set to determine if the row above contained the same value for the ticker column; if not, it was determined to be the start.

```If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then```

Here the variable `tickers` refers to an array that contains all the known stock tickers, such as `tickers(2) = "DQ"`, with the `tickerIndex` being a variable that is incremented when we determine the end of a stock's data. The end point is determined when a row where the following row contains a different value in the stock ticker column is found in a similar conditional statement as above.

To determine the Total Daily Volume, the value in the Volume column is totalled on each iteration of the loop stored in an array that also utilizes the `tickerIndex` variable, meaning that `tickerVolumes(2)` refers to DQ's Total Daily Volume as it is related to the same index used for the tickers themselves.

The yearly Return uses the same conditional statments that determine the starting and ending point of a stock's data and sets `startingPrice(tickerIndex)` and `endingPrice(tickerIndex)` to the values in the Price column on the data sheet. The calculation is performed in a loop that also writes the values contained in the `tickers` and `tickerVolumes` to the analysis sheet:

```
For i = 0 To 11
  Worksheets("All Stocks Analysis").Activate

  Cells(4 + i, 1).Value = tickers(i)
  Cells(4 + i, 2).Value = tickerVolumes(i)
  Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Next i
```

This produces two tables:

| 2018 Stock Analysis | 2017 Stock Analysis |
| :---: | :---: |
| ![2018 Stock Analysis](/resources/VBA_2018_stocks.png) | ![2017 Stock Analysis](/resources/VBA_2017.stocks.png) |

In general most of the clean energy stocks used in this analysis did poorly in 2018, especially in comparison to 2017. Nearly all tickers dropped in total shares being traded and all but 2 had a negative Return, a stark contrast to 2017 where all but one had a positive Return. Two tickers stand above the rest: ENPH and RUN both grew in Total Daily Volume from years 2017 to 2018 and maintained a positive Return.

### Analysis of Code Performance

Originally this data was analyzed with a macro involving nested For loops. Only an array to keep track of the `tickers` was created, and the first loop iterated over each value in this array. A second loop nested in this array iterated over the rows in our data sheet that retrieved the desired values using three conditionals:

1) `If Cells(j, 1).Value = ticker Then` is used to identify all values in the tickers column that match a ticker in the array and sums the values in the Volume column to obtain `totalVolume`.

2) `If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then` to find the `startingPrice`.

3) `If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then` to find the `endingPrice`.

After all these values are determined, the nested loop ends after it reaches the last row in the spreadsheet. Returning to the outer loop, it activates the output analysis spreadsheet to write in the values obtained in the nested loop similar to the refactored macro described in the Stock Perfomance Analysis. After writing these values the index for the `tickers` array increments by 1, meaning this macro runs the nested For loop that iterates over each row in the data spreadsheet 12 times.

To compare the performance of these two macros a timer was set to start right after the user inputs the desired year to perform the analysis on. The results of those times for the year 2018 are as follows:

| Original 2018 Macro speed | Refactored 2018 Macro speed |
| :---: | :---: |
| ![Original 2018 Macro](/resources/VBA_Challenge_2018.png) | ![Refactored 2018 Macro](/resources/VBA_Challenge_2018_Refactored.png) |

## Summary

The highlight of this project is showcasing the usefulness in refactoring code. Refactoring code can improve the readability, fix poor code design and even increase performance. The only real drawback is that it that it means spending more time on a single project, especially when there is a lack of experience in spotting potential flaws in design.

In the case of this analysis, refactoring the original macro that used a nested loop allowed the macro to run nearly 10 times faster; the improvement in run time can be largely attributed to the refactored version only needing to loop through the dataset 1 time as opposed to 12 times. Additionally, the refactored version has improved readability by using the `tickerIndex` variable to more cohesively relate the values obtained from the data spreadsheet to the the tickers in the `tickers` array. The refactored script is far better to use on a larger dataset where increasing the number of rows could drastically the script's run time. One of the flaws comes from VBA requiring the use of explicitly declaring the number of elements for the array; to analyze more tickers, each array must be updated to match the new amount.
