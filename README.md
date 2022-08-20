# `Stock Analysis using VBA.`


## `Overview of Project.`<br/>
We have prepared the workbook for Steve to analyze the green energy stock market for year 2017 and 2018.
At the click of a button, he can analyze an entire dataset. 
Now, to do a little more research for his parents, he wants us to expand the dataset to include the entire stock market over the last few years.<br/>
## `Purpose of Project.`<br/>
The purpose of this project is to take our existing VBA code which was run on a handful of stocks and make it be able to handle 
thousands of stocks as quickly as possible.<br/>
## `Analysis and Results.`<br/>
Looking at the results of the code for the different years show very different results in terms of returns. 
During year 2017, most of the picked stocks were doing well and having positive returns except ticker TERP, which was down 7.2%. 
The biggest return was ticker DQ, and its return was 199.4%.<br/>
#### Table 1: 2017 Stock Analysis.<br/>
![](https://github.com/Valeriia161/stock-analysis/blob/main/VBA_Challenge_2017.png.png)


In 2018 stock market was not in a good shape. Only two tickers, ENPH and RUN had positive return of 81.9% and 84%. 
Other 10 tickers suffered with negative returns; the biggest drop was DQ, fell 62.6%. <br/>

#### Table 2: 2018 Stock Analysis.<br/>
![](https://github.com/Valeriia161/stock-analysis/blob/main/VBA_Challenge_2018.png.png)


Elapsed run time for the refactored code is faster than in original code.<br/>
The 2017 program managed to run in 0.078125 seconds (~original code = 0.6796875 seconds~).<br/>
#### Table 3: 2017 code execution times.<br/>
![](https://github.com/Valeriia161/stock-analysis/blob/main/2017%20code%20execution%20times_old.png.png)  ![](https://github.com/Valeriia161/stock-analysis/blob/main/2017%20code%20execution%20times_new.png.png)



The 2018 program managed to run in 0.0625 seconds (~original code = 0.6992188 seconds~).<br/>
#### Table 4: 2018 code execution times.<br/>
![](https://github.com/Valeriia161/stock-analysis/blob/main/2018%20code%20execution%20times_old.png.png)  ![](https://github.com/Valeriia161/stock-analysis/blob/main/2018%20code%20execution%20times_new.png.png)


## `Summary.`<br/>
#### 1.	What are the advantages or disadvantages of refactoring code? <br/>
*There are many advantages of refactoring the code:* <br/>
- making code more useful; <br/>
- making code more efficiently and quicker. <br/>

*Theare are also disadvantages of refactoring the code:* <br/>
-	Making code more complicated; <br/>
-	Making code run into bugs. <br/>

#### 2.	How do these pros and cons apply to refactoring the original VBA script?<br/>
Refactoring the original code has allowed for nested loops to be removed leaving more flexible code that executes more efficiently and allows more stock tickers to be evaluated in the future. The downside of refactoring the original VBA script was the time spent on understanding creation of the index variable and correctly applying it to the looping logic.

