# stock-analysis
## Overview
  This analysis uses Excel and Visual Basic for Applications (VBA) to analyze stock performance for some companies to help a financer recommend the right investing options for his first clients
### Background
The data used in this analysis covers the stock performance for 12 green energy companies over the years 2017 & 2018.
### Purpose
The purpose of this analysis is to provide a clear view of the performance of each of the twelve green energy companies stocks over the two years, showing the total volume & the return of each of them. This should help the financer decide which stock options are good for his clients to invest in, as well as analyze the performance of one particular company they prefer (DAQO New Energy Corp)
Another purpose is to refactor the analysis sheet to perform the required analysis in the least amount of time

## Results
### Stock Analysis results
After analyzing the stock performacne of companies over the years 2017 & 2018, we can notice that almost all green energy companies (with the exception of RUN) had a much better performance in 2017 than in 2018. It is also obvious that the preferred comapny DQ lost more than 60% of its value in 2018, which means that it is not a good option for investment.


<img width="235" alt="2017" src="https://user-images.githubusercontent.com/79733383/111095640-46499680-8514-11eb-87e4-30ab3709b543.PNG">

<img width="233" alt="2018" src="https://user-images.githubusercontent.com/79733383/111095670-56fa0c80-8514-11eb-8f8e-d4cc3df5b4b5.PNG">

Looking over the data shown in the above figures, it is obvious that the only two companies that kept a positive performance over both years are ENPH & RUN, both comapnies had an increased overall volume & a positive return when all other companies either had a lower volume or a negative return or both.
The data also shows that the stocks of RUN had an increased return, from 5.5% to 84%, and the volume is almost doubled. And that the stocks of ENPH had a smaller return in 2018 than in 2017, 81.9% & 129.5% respectively, while the volume increased by more than double.
Based on this analysis, it would be recommended to narrow the investments in green energy companies between ENPH & RUN, with a bigger portion in RUN as it shows a positive increase. It would also be beneficial to investigate the reason of the poor performance of all the companies in 2018 before deciding to invest.
### Refactoring results
After analyzing the data using a nested For function, with variables for both tickers and rows, as shown in this link
https://github.com/Lamismn/stock-analysis/blob/main/original%20code.PNG
We decided to try and refactor the code to try and reduce the time & processing required to go through the data, to do this we used a tickerIndex to go through tickers while using a single For function to go through the rows, as shown in the following link
https://github.com/Lamismn/stock-analysis/blob/main/refactored%20code.PNG
After testing the time elapsed in each scenario, we realized that the refactoring reduced the elapsed time dramatically as shown in the following figures

<img width="275" alt="2017 original" src="https://user-images.githubusercontent.com/79733383/111097042-210a5780-8517-11eb-98cf-6a6d444da007.PNG">
<img width="297" alt="2017 refactored" src="https://user-images.githubusercontent.com/79733383/111097051-28316580-8517-11eb-8f84-d8d28b48955e.PNG">
<img width="296" alt="2018 original" src="https://user-images.githubusercontent.com/79733383/111097065-2ff10a00-8517-11eb-916b-b9334d4404fd.PNG">
<img width="302" alt="2018 refactored" src="https://user-images.githubusercontent.com/79733383/111097075-34b5be00-8517-11eb-983d-3c21fc64b36c.PNG">

## Summary
After running the code both originally and after refactoring, we can conclude that refactoring a code has a great advantage whne it comes to saving the processing power and the elapsed time runnung a specific code. But we can also see that refactoring a code is a very time consuming process on its own. This means that while it can enhance the code performance, it may overall consume more time if this code is not supposed to be used for different scenarios. If we are writing a code for a one time use, refactoring it may take twice as much time to finish the task, which is not efficient. Whereas if we are writing a code that will be used multiple times under different variables, it makes more sense to spent more time refactoring it rather than changing it everytime we use it.

For our analysis, we can see that after we refactored the code, it makes it easier for the financer to use this code for other stocks in the future, and he will only have to change the ticker string in the beginig of the code & some minor changes within the code. If, however, he wants to use this code for this analysis only, refactoring it would not be very beneficial.
