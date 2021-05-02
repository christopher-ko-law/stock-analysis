# Stock Analysis

## Overview of the Project

### Purpose
The purpose of this analysis is to analyze an entire dataset of stocks within a worksheet in excel. Using a macro based on refactored code, the following data will be extracted to the "AllStocksAnalysisRefactored" worksheet:
1. Ticker
2. Total Daily Volume
3. Return

The results of this analysis will be compared to the original script to ensure that the results are identical. The execution times will be compared as well to determining if refactoring the original script was beneficial.


## Results

### Stock Performance

Below is a chart showing stock performance between 2017 and 2018.

![Stock Analysis](/Resources/stock-analysis-chart.png)

We see that stocks in 2018 generally posted a negative return then in 2017. Of the 12 stocks that were analyzed, only 2 - ENPH and RUN posted a positive return.

### Execution Times of the Original Vs. Refactored Script

The original script was created with a nested for loop that ran through the rows in the worksheet once per ticker.
```
For i = 0 To 11
    
        ticker = tickers(i)
        
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        
        For j = rowStart To RowEnd

            If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

            End If
    
    
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value
                
                End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value
        
            End If
        
        Next j
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
```

The refactored script removed the nested for loop and instead ran through the rows in the worksheet once for all tickers.
```
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            'set starting ticker prics if this is the first row
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) check if the current row is the last row with the selected ticker

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            ' set ending ticker priceif this is the last row.
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1

        End If
   
    Next i
 ```

Below are screenshots of the original execution times.

|     | Original | Refactored |
| --- | --- | --- | 
| 2017 | ![2017 Original](https://github.com/christopher-ko-law/stock-analysis/blob/main/Resources/AllStocksAnalysis-2017-Original.png) | ![2017 Refactored](/Resources/AllStocksAnalysis-2017-Refactored.png) |
| 2018 | ![2018 Original](/Resources/AllStocksAnalysis-2018-Original.png) |![2018 Original](/Resources/AllStocksAnalysis-2018-Refactored.png) |

The original and refactor scripts were run 3 times each for a spread of execution times. Below is the table of execution times.

| | 2017 Original | 2018 Original | 2017 Refactored | 2018 Refactored |
| --- | --- | --- | --- | --- |  
| Run 1 | 0.7968875	| 0.7617188	| 0.1523438	| 0.140625 |
| Run 2 | 0.7265625	| 0.7539062	| 0.137188 | 0.13675188 |
| Run 3 | 0.7226562	| 0.734375 | 0.140625 |	0.1445312 |
| Mean | 0.748702067 | 0.75	| 0.1433856	| 0.140636027 |

We see here that the refactored code runs around 5x faster then the original script.

## Summary

### Advantages and Disadvantages of Refactoring Code

**Advantages**
Refactoring means to restructure your existing code, while keeping the existing functionality. Typically we refactor for the following reasons:
* Keeping your code clean - Removing/optimizing duplicated/uneeded functions
* Improve performance - Generally when we refactor, we improve performance as the underlying code is reviewed and optimized.

**Disadvantages**
Refactoring needs to be done in small steps. When working with legacy code, care must be taken as to not unintentionally remove any functionalities and/or create any new bugs. This means that each refactored piece of code, must undergo rigorous testing before being released. 


### Summary of Refactoring the Original VBA Script
In light of the notes above, we see that the original VBA script read each line of the worksheet once per ticker. If there were 1000 tickers and 1000 lines, the nested loop would have ran 1 million times. This is very computationally heavy.

The refactored code instead, runs through the worksheet once. If there were 1000 tickers and 1000 lines, the loop would only have ran 1000 times.

This is the reason why the refactored code runs much faster.

On the other hand, we may have introduced a new issue that wouldn't have been caught by this particular worksheet. In the refactored code, we assume that the tickers in the spreadsheet are all sorted and in alphabetical order (I.E - our data is clean). However that may not be the case with other worksheets. The original code would be able to parse worksheets with ticker blocks out of alphabetical order, but the new refactored code would not.

In order to continue refactoring this code, we need to know what the expected inputs to the Macro are supposed to be.










