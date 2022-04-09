# Stock Analysis With Excel VBA


## Overview of Project
**Pupose**

The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code. 

### The Data
The data that is presented includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.


## Results
### Analysis
Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file. 
![image](https://github.com/morriscomia/stock-analysis./blob/08ab142702045299594d133fa65e4eb017a91eda/Refactored%20code.png)

## Summary
### Pros and Cons of Refactoring Code
Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design and software improvement, debugging, and faster programming. It may also benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. However, we do not always have the luxury to refactor our code due to disadvantages. These disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code. 

### The Advantages of Refactoring Stock Analysis
The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time. The original analysis took approximately .21 second to run, whereas our new analysis only took about approximately 0.20 seconds to run. Attached below are the screenshots that indicate the run time for our old and new analysis.
![image](https://github.com/morriscomia/stock-analysis./blob/08ab142702045299594d133fa65e4eb017a91eda/Old%202017%20Run%20time.png)
![image](https://github.com/morriscomia/stock-analysis./blob/08ab142702045299594d133fa65e4eb017a91eda/2017%20run%20time.png)
