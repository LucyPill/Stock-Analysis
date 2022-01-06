# Green Stocks Analysis for the years 2017 and 2018
## Project Overview:
  Steve’s parents want to invest on green stocks and asked Steve to do an analysis on green stocks. Steve asked for help and using VBA in excel, I was able to create a macro that helps Steve find the stock’s total volume and the annual return for each stock for the years 2017 and 2018.

## Purpose of the Analysis:
  The purpose of this project was to be able to use VBA to create a macro to analyze different stocks. I was able to create a macro that contained nested loops to run the analysis for 12 different stocks. However, I learned that there was a more efficient way to do the same analysis by refactoring my code, so the end goal is to use the original macro and refactor it to accomplish the same analysis in a more efficient manner.

## Results and Disccusion
  The first thing, I needed to do was to think and figure out a way of how I can accomplish my task of refactoring my code. In order to do this: I created a variable called tickerIndex, created four arrays: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The ticker array was then used to determine the ticker of a stock. Then, the ticker array was correlated to the other three arrays by using the tickerIndex variable. Setting the tickerIndex variable was helpful because allows me to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to individual ticker before looping through the data. Creating arrays and creating the tickerIndex variable was a more efficient way than using the nesting for loop that we used earlier before we refactored our code. Now our code runs faster. I was successful in refactoring the code to make it more efficient.

  We can see that the first time we ran the code the running time for 2017 was 0.5625 seconds and when we ran the refactored code the running time was in 0.1953125 seconds. We can conclude that the refactored code runs 3X faster than the original code for the year of 2017.

### Original Code for year 2017
![VBA_Challenge_Original_2017.png](https://github.com/LucyPill/Stock-Analysis-/blob/main/VBA_Chellenge_Original_2017.png)

### Refactored Code for year 2017
![VBA_Challenge_Refatored_2017.png](https://github.com/LucyPill/Stock-Analysis-/blob/main/VBA_Challenge_Refactored_2017.png)


  Similarly, for the year of 2018 we got in 0.578125 seconds for the original code and 0.1992188 seconds for the refactored code. Again, the refactored code runs 3X faster than the original code. This difference in time can be meaningful when running code in a complex set of data because time is critically important in many aspects of businesses and sometimes seconds can be the determining factor in closing a deal.

### Original Code for year 2018 
![VBA_Challenge_Original_2018.png](https://github.com/LucyPill/Stock-Analysis-/blob/main/VBA_Challenge_Original_2018.png)

### Refactored Code for year 2018
![VBA_Challenge_Refactored_2018.png](https://github.com/LucyPill/Stock-Analysis-/blob/main/VBA_Challenge_Refactored_2018.png)

  Finally, we can compare the stocks’ performance for 2017 and 2018. For the year of 2017 the stocks performed very well and for many 2017 was an epic year. We can see that 2018 was the opposite and this is in accordance with what happened in 2018 where the stock market did not do well. It has been said that 2018 was one of the worse years for stocks. Nonetheless, Steve now has a code that he can use to run the analysis and use it to make an educated decision about what to tell his parents.

## Summary:
  Looking back at the way I worked through this challenge to refactor the code, I can see advantages and disadvantages. One of the advantages is that you are not starting from scratch, and we are taking a piece of work that we already have and making it more efficient. Although, this is an advantage it can also be a disadvantage because if we are not careful enough, we can end up messing up our code and now the code that we had to begin with is no longer useful. This outcome can be detrimental to our work. That is why we need to take precautions before starting the refactoring of the original code. What I always do is that anytime I have to do work I make a copy of the file and work on the copy just in case something happens I can always revert or go back to my starting point.
  It is hard for me to think ad compare advantages and disadvantages of refactoring code in VBA because I have never coded or refactored a code before. This is the first time I am doing it, but I can think of few things that were very helpful such as minimizing the amount of work and making the code more efficient like performing complex tasks in few and simple steps. Users can use the developer tools making it easier to input and output data with a much simpler interaction than doing it in the excel sheet. We can simplify a repetitive task using the developer tool in VBA.

