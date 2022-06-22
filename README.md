# Stock Analysis
---
## Overview of Project 
- The purpose of this analysis is to further knowledge on the various ways to use VBA to code outcomes with the same conclusions in order to compare run time of the code and to potentially simplify the code itself. In this case, we refactored the code to loop through the data in less steps than in our previous work and to decrease the amount of time it took to run the macro.
---
## Results 
---
### Run Times for Original Code 
![Original_Code_2017](/Original_Code_2017.png)
![Original_Code_2018](/Original_Code_2018.png)
---
### Run Times for Refactored Code
![VBA_Challenge_2017](/VBA_Challenge_2017.png)
![VBA_Challenge_2018](/VBA_Challenge_2018.png)
---
### Stock Performance 
- For the same tickers, stock performance saw an overall trend of decrease in the Return in comparison to 2017 to 2018. In 2017, only 1 stock had a negative return. In 2018, only 2 stocks had positive returns. The TERP stock actually saw an increase in return from 2017 with a return of -7.2% in 2017 and a return of -5.0% in 2018. While still negative, this does reflect an increase. In addition, the ENPH stock, while positive for both years, actually saw a decrease in the return from 129.5% in 2017 down to 81.9% in 2018. The RUN stock saw a dramatic shift in stock return with the value shooting up from 5.5% in 2017 to 84.0% in 2018. 
### Execution Time Comparison
- As you can see in the images above, the code ran significantly faster (by about one whole decimal place) from the Original Code performance to the Refactored Code performance. 2017 analysis saw a jump from 0.69 seconds to 0.078 seconds, and 2018 analysis saw a jump from 0.63 seconds to 0.082 seconds. 
---
## Summary
-  I would say that the advantage to refactoring code would be this faster run speed. We can see through this challenge that while the refactored code follows a similar pattern to the original code, the run time did get better after refactoring. However, I would say that a disadvantage of refactoring code is the change for creating the opportunity for new bugs or errors to occur. In this example, while we did see the change in run time, I did notice that I ran into some more problems that came with deeper analysis and understanding of VBA. In addition, the difference in time, while noticeable after printing out the run time and comparing the images side by side, was not sufficient enough that I could tell the difference in my own observations without being told the change in value. So in this instance, with this specific example, it could be argued that the change in time might not have been necessary for the more substantial time it took to put into changing the code (aside from of course the teaching exercise!) 
