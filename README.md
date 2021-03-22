# Stock Analysis

## Overview of Project
Expanding workbook for client, Steve, with refactored VBA code to analyze twelve green energy stocks performance over 2017-2018 


### Stock Results
Looking at the overall green energy market performance (using the twelve stocks Steve had selected for analysis) we saw a higher average return in 2017, +67.5% compared to -8.5% in 2018. There were two stocks that performed well both years: Enphase Energy, Inc (ENPH) and Sunrun, Inc (RUN), with two additional stocks that averaged more positive returns than the others, DAQO New Energy Corp (DQ) and Solar Edge Technologies, Inc (SEDG) Ranking average Total Daily Volume of shares traded from lowest to highest would be DQ with the least number of shares traded, SEDG, RUN, and then ENPH (having almost tripled from '17-'18). 

### Refactoring Results
Concerning the performance of the VBA script when refactored, the execution time showed a small but significant decrease compared to the original script. 

Refactored Results for 2017:
<img width="817" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/78561980/111930775-afd32300-8a87-11eb-9120-a1ab607b2ef9.png">

Original Script Results for 2017:
<img width="816" alt="2017_Stock_Analysis" src="https://user-images.githubusercontent.com/78561980/111930788-b366aa00-8a87-11eb-93c6-5335dea28c39.png">

Refactored Results for 2018:
<img width="818" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/78561980/111930854-de50fe00-8a87-11eb-9458-b268f68677d6.png">

Original Script Results from 2018:
<img width="816" alt="2018_Stock_Analysis" src="https://user-images.githubusercontent.com/78561980/111930879-ea3cc000-8a87-11eb-9d56-549a83edc43b.png">


# Summary
The stock analysis we performed successfully identified four companies in the green energy sector with positive returns, and offers us the framework to analyze additional stocks in the future. We were able to reduce the execution time by 82% by utilizing the refactored VBA script, which will allow us to utilize this code on larger datasets using more stocks and a longer timeline for comparision. 


### Refactoring: Cost/Benefit Analysis  
Our goal with refactoring code is to use an existing, working code, and restructuring it with improved efficiency, decreased execution time, better readability, and less room for errors while maintaining the code's function. The advantages for this are more obvious on a larger scale, so depending on how often the code will be run, whether different people (besides the original coder) will be executing the code, and whether the code will be reused or adapted, etc. A refactored code will save time depending on scale of use. The disadvantages of refactoring code include time spent reworking a funtional code, which may not be worth the cost of refactoring if the life of the code will be short, since the original code fuctioned as intended.


### VBA: Comparing scripts
This VBA script refactored will prove worth the effort with the reduction in execution time for Steve. The original was effective for the smaller dataset, simple to read, and simple to write with the nest for loop. The main reasoning behind refactoring this particular code was to allow us to increase the size of the dataset, which wouldn't have been ideal using the original, having a loop repeat through the data multiple times. The refactored code allows us a single loop through the data, and more agile framework as the coding is layered versus being interwoven. Refactoring in VBA is challenging and there was a lot of troubleshooting and testing pieces of code. VBA is not the most clear when it comes to communicating errors and debugging, which can cause a lot of wasted time during the refactoring process. However, VBA does make it easy to essentionally shut off parts of the code by adding ' at the beginning of each line, so nothing is lost in testing code. It is also quite simple to retain the original code in a seperate Module in VBA, and edits can be made during the refactoring process that don't jeopardize the funtionality of the original. 

