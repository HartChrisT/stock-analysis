# Stock Analysis

## Overview of Project

### Purpose

To refactor a VBA script used to analyze a dozen stocks so it runs more efficiently and can be applied the entire stock market. 

## Results
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92996865/148011190-fd3c4d57-6c95-4983-8283-4e07545f308d.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/92996865/148011195-4fe95a2d-9ff3-4a49-85eb-aaa3b29744f9.png)
Stocks performed a lot better overall in 2017 than in 2018. Eleven of the twelve stocks generated positive returns in 2017 as opposed to only two of the twelve stocks in 2018.

The refactored code ran significantly faster! Here are the run times for 2017 and 2018 before the refactor.
![2017 prefactor](https://user-images.githubusercontent.com/92996865/148012919-d2828e02-1c6f-4d98-8ed2-5f07a0e71033.png)
![2018 prefactor](https://user-images.githubusercontent.com/92996865/148012924-ce3041f4-e332-4838-9f9c-60c67e84fa7d.png)

It took approximately one tenth of a second for the refactored code to run as opposed to the .77 seconds the old code to run making the refactored code approximately 6.7 times faster when using the 2018 refactor time of .10 seconds.

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring has its many advantages, such as: making code run faster, making it easier for the reader or user to understand, and generalizing it can be applied more broadly.

One disadvantage of refactoring code is the opportunity cost of recreating something that already works instead of working on something new. Chances are that new projects requiring similar structures will arise and the time to refactor will become apparent to save time on future projects.

### How these Pros and Cons Apply to refactoring the original VBA Script

The refactored code ran significantly faster than the previous code as it only looped through the data set once instead of 12 times for the current data set, or i times for a data set with i tickers. Since that data set gets longer as well, we can see how this inefficiency can compound on itself rapidly making it rather cumbersome to compute.

However, the difference in actual run time is negligible for a data set this small, therefore a refactor would discouraged if no more analysis was needed beyond these twelve stocks. 
