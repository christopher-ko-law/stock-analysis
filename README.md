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

Below are screenshots of the original execution times.

|     | Original | Refactored |
| --- | --- | --- | 
| 2017 | ![2017 Original](https://github.com/christopher-ko-law/stock-analysis/blob/main/Resources/AllStocksAnalysis-2017-Original.png) | ![2017 Refactored](/Resources/AllStocksAnalysis-2017-Refactored.png) |
| 2018 | ![2018 Original](/Resources/AllStocksAnalysis-2018-Original.png) |![2018 Original](/Resources/AllStocksAnalysis-2018-Refactored.png) |


## Summary

### Refactoring Code

### Pros and Cons of Refactoring the Original VBA Script