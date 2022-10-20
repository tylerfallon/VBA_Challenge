# VBA_Challenge

## Overview and Purpose
The purpose of this project is to analyze the data of 12 stocks and optimize the VBA code that performs the analysis so that it runs quickly and efficiently. 

## Results

Stock performance for all stocks besides 1 (TERP) was higher in 2017 compared with 2018. This can be seen on our analysis images below. In 2017, returns for 11/12 stocks were a net positive, whereas in 2018, returns for 10/12 stocks was negative. 

### Original Analysis for 2017

![Original_Analysis_2017](https://github.com/tylerfallon/VBA_Challenge/blob/main/Resources/Original_2017.png?raw=true)

Execution time was 0.2617188 seconds.

### Refactored Analysis for 2017

![Refactored_Analysis_2017](https://github.com/tylerfallon/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

Execution time was 0.08203125 seconds.

### Original Analysis for 2018

![Original_Analysis_2018](https://github.com/tylerfallon/VBA_Challenge/blob/main/Resources/Original_2018.png?raw=true)

Execution time was 0.265625 seconds.

### Refactored Analysis for 2018

![Refactored_Analysis_2018](https://github.com/tylerfallon/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

Execution time was 0.07421875 seconds.

### Execution Time Comparison

The refactored analysis code executed more than 3x faster than the original analysis code. Both sets of scripts created the same results data, but the refactored code was far more efficient at the task. 

## Summary

### Advantages and Disadvantages to Refactoring Code

A major advantage to refactoring code is to improving the speed of execution. This could play a major role when scripts are used to analyze data on a large scale. 

Refactoring code can also increase the readability of code and may make it easier to maintain in many cases. Creating more efficient code is incredibly important in the workplace as well as code runtimes can get very long, which wastes time and reduces the productivity of the business. 

A disadvantage to refactoring code is that imprecise refactoring can result in new bugs or errors. While code may appear to produce the same results on a small set of data, it may perform differently on different or new datasets introduced in the future, which can skew data and cause issues. 

### Why refactor our original VBA script

In our case, refactoring the original VBA script was important. Not only did the new refactored script analyze and format the cells, but it also performed significantly faster than the original VBA script. We used our script to analyze the data for 12 stocks, but on a larger scale with hundreds or thousands of stocks, the runtime of the script could take minutes or even hours. In this case, ensuring that our code is optimized for speed and runtime, while still maintaining accuracy, would promote extendability of our code and improve productivity for stock analyzers. 

