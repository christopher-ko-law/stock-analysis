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

The original and refactor scripts were run 3 times each for a spread of execution times. Below is the table of execution times.

| | 2017 Original | 2018 Original | 2017 Refactored | 2018 Refactored |
| --- | --- | --- | --- | --- |  
| Run 1 | 0.7968875	| 0.7617188	| 0.1523438	| 0.140625 |
| Run 2 | 0.7265625	| 0.7539062	| 0.137188 | 0.13675188 |
| Run 3 | 0.7226562	| 0.734375 | 0.140625 |	0.1445312 |
| Mean | 0.748702067 | 0.75	| 0.1433856	| 0.140636027 |

We see here that the refactored code runs nearly around 5x faster then the original script.

## Summary

### Refactoring Code

### Pros and Cons of Refactoring the Original VBA Script