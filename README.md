# Stock-Analysis-Module-2
## Overview of Project: 
The objective of the project is to make analysis of the stock dataset and to provide information on transaction volumes and stock returns of 11 stocks in 2017 and 2018. 


## Results:
### 1. Comparison of Stock Performance Between 2017 and 2018

In 2017, except TERP stock's return, all other 10 stocks' returns are positive. The top 3 return in 2017 are those of stock DQ, ENPH, and FSLR with the rate of 199.4%, 129.5% and 101.3%, respectively. On the contrary, except ENPH and RUN with the rate of return, 81.9% and 84%, respectively, all other 9 stocks' returns are negative, which means that these stocks' prices decrease in 2018.  As for total daily transaction volume, the total daily volumes of stock DQ, ENPH, HASI, RUn, SEDG, TERP and VSLR present an increasing trend from 2017 to 2018 while the total daily volumes of AY, CSIQ, FSLR, JKS and SPWR present a decreasing trend. In 2017, the 3 largest daily transaction volumes are those of SPWR, FSLR, and CSIQ with the amount of 782,187,000; 684,181,400; and 310,592,800; repsectively. As for 2018, the 3 largest daily transaction volumes are those of ENPH, SPWR, and RUN, with the volumes of 607,473,500; 538,024,300; and 502,757,100; respectively. Generally speaking, comparing the differences of daily volumes and the return in 2017 and 2018, the increase in total daily volume contributes to the decrease in price and the decrease in rate of return. 

![2017-challenge-result](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/2017-challenge-result%20.png)

2017-challenge-result


![2018-challenge-result](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/2018-challenge-result.png)

2018-challenge-result

### 2. Comparison of the Execution Times of the Original Script and the Refactored Script 


![2017-exercise](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/2017-exercise.png)

2017-exercise


![VBA_Challenge_2017](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/VBA_Challenge_2017.Png)

VBA_Challenge_2017


![2018-exercise](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/2018-exercise.png)

2018-exercise


![VBA_Challenge_2018](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/VBA_Challenge_2018.Png)

VBA_Challenge_2018

![code](https://github.com/irisyidi/Stock-Analysis-Module-2/blob/main/code.png)

Code Screenshot

From the screenshots of the pop-up messages showing the elapsed run time for the original script and the refactored script, the elapsed run time for refactored scripts is less than that for the original script. The original code ran in 0.2890625 seconds for the year 2017 while the refactored code ran 0.265625 seconds for the year 2017. The orignal code ran in 0.2734375 seconds for year 2018 while the refactored code ran for 0.2617188 for year 2018. 
The decline of the run time attributes to the improvement of the code by taking fewer steps and using less memory and improving the logic of the code. By using the tickerIndex ,the logic of the code has been improved and less memory is used in the process. As for the reason why less time is consumed when running the code for 2018 than that for 2017, it has connection with workload of calculation. 

## Summary:
### 1. What are the advantages or disadvantages of refactoring code?

#### Advantages: 
The refactoring is important process to make the code more efficient by taking fewer steps, using less memory or improving the logic of the code to make it easier for future users to read. the refactoring process facilitates the check of strategy to using better methods. Meanwhile, refactoring the code makes the code more reusable, clear and simple, which will improve the working efficiency to work with the similar problems. 

#### Disadvantages: 
The disadvantages is that the refactoring code does no add new function and the user has no idea of how long it will cost him to complete this process. Besides, there is potential risk if the refactored code is applied. It may waste much more time to solve the problems than to compile new code due to complexity of the refactored code. 

### 2. How do these pros and cons apply to refactoring the original VBA script?
When refactoring the orignal VBA script, the application of the tickerIndex to increase the tickerVolume, tickerStartingPrice and tickerEndingPrice makes the logic more clear and easier. Separating the loop to output the ticker, total daily volume and the return with that to calculation of those takes fewer steps and uses less memory, which saves the time to run the code. 
