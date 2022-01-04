# Stock Analysis

## Overview of Project

### Purpose

To refactor a VBA script used to analyze a dozen stocks so it runs more efficiently and can be applied the entire stock market. 

## Results
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92996865/148011190-fd3c4d57-6c95-4983-8283-4e07545f308d.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/92996865/148011195-4fe95a2d-9ff3-4a49-85eb-aaa3b29744f9.png)
Stocks performed overall a lot better in 2017 than in 2018

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring has its many advantages, such as: making code run faster, making it easier for the reader or user to understand, and generalizing it can be applied more broadly.

One disadvantage of refactoring code is the opportunity cost of recreating something that already works instead of working on something new. Chances are that new projects requiring similar structures will arise and the time to refactor will become apparent to save time on future projects.

### How these Pros and Cons Apply to refactoring the original VBA Script

The refactored code ran significantly faster than the previous code as it only looped through the data set once instead of 12 times for the current data set, or i times for a data set with i tickers. Since that data set gets longer as well, we can see how this inefficiency can compound on itself rapidly making it rather cumbersome to compute.

However, the difference in actual run time is negligible for a data set this small, therefore a refactor would discouraged if no more analysis was needed beyond these twelve stocks. 
