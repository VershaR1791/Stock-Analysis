# Stock-Analysis
Performing analysis on Stock data to find top performing companies by refactoring the code
## Overview of Project
Steve is a financial analyst who has started a new role in a finance company and would like to analyze stock data for different alternative energy companies. He would want to analyze the data to recommend to his parents on which company's stocks they should invest in based on their past performance. The analysis will include refactoring the exisiting code to make it more efficient. 
### Purpose
The purpose of this analysis is to automate the large excel data through the ease of a click of a button, for Steve to be able to review the return for all the stocks. The data available is for different alternative energy companies and corresponding information regarding the price and volume their stocks closed at. This analysis will -
- help summarize the return and total daily volume from the company stocks for years 2017 and 2018
- determine the time to execute the code before and after refactoring the code.

## Analysis
The refactoring code includes creating a tickerIndex to loop through all the list of companies repeatedly without specifying the name of the company. By creating the tickerIndex we are restricting the code to toggle between the sheets in the workbook multiple times. In the refactored code the code is optimized by storing the values within arrays and calling them back at the end to print the results.

The success of refactoring a code can be determined by the time it takes to execute the code. The following 2 snapshots are from the original code execution.

![Timer_2017](https://user-images.githubusercontent.com/84694664/125178609-c88d5300-e1b4-11eb-9b14-12a171223483.JPG)
![Timer_2018](https://user-images.githubusercontent.com/84694664/125178611-ca571680-e1b4-11eb-9f9d-00f37fd9b497.JPG)

The following 2 snapshots are from the refactored code which definitely is faster than the original code execution.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/84694664/125178613-cd520700-e1b4-11eb-901b-32ad2b07bbbb.JPG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/84694664/125178615-ce833400-e1b4-11eb-9b13-9a0fe8120cff.JPG)

## Results
The following table is the summary of the stock returns at the end of the year 2017. Almost all the companies had a positive return except *'TERP'*. The companies which did exceedingly well were - *'DQ', 'ENPH', 'SEDG'* and *'FSLR*.

![2017_Results](https://user-images.githubusercontent.com/84694664/125178600-b3b0bf80-e1b4-11eb-8df3-a6a1d8796fe1.png)

The following table is the summary of the stock returns at the end of the year 2018. This seems to have been a difficult year for almost all companies since their returns were negative. The only 2 companies which did reasonably well were *'ENPH'* and *'RUN'*. It's to be noted that *'RUN'* had an increased stock return inspite of a lowered return for other companies.

![2018_Results](https://user-images.githubusercontent.com/84694664/125178603-b7444680-e1b4-11eb-99af-54b6bf9ed523.png)

Overall **'ENPH'** performed well in both the years and Steve should recommend his parents to invest in their stocks. In order to diversify his parents funds he can also ask his parents to invest in low risk stocks which overall have a positive return such as *'DQ', 'SEDG', 'FSLR'* and *'RUN'*.

### Advantages or disadvantages of refactoring code
- The most obvious advantage of refactoring code is that it makes it more efficient by making the code run faster and use less memory. It also improves the logic of the code.
- Refactoring also makes the code easier to understand and helps find bugs in the code.
- The disdavantages includes that one can not add any features or functionalities mainly due to budget or time contraints in a real world.
- There could also be risk in refactoring a code when the application is too big.

### Application of pros and cons apply to refactoring the original VBA script
