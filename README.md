# Stock_Analysis

## Overview of Project
Steve wants to do more stock research for his parents but would like for it to be quick and efficient. This required the following things.

1. Refactoring the Module 2 solution code to loop through all the data one time in order to collect the information.
2. Determine whether refactoring the code successfully made the VBA script run faster.
3. Present the exact amount of time it took for the VBA script to run through the stocks.

## Resources
Data source: VBA_Challenge.xlxs
Software: Microsoft Excel 2019, Microsoft Visual Basic for Applications

## Results
Using both for loops and nested for loops the VBA code was able to run through the entire data set. After refactoring the VBA code, I successfully reduced the process time of the analysis. I achieved this by reducing the number of itirations by using arrays as well as setting up an index that was being used by the four arrays.

<img width="215" alt="2017_election_analysis" src="https://user-images.githubusercontent.com/85372441/124336713-a8321880-db64-11eb-982b-7a868694466a.png">

<img width="211" alt="2018_election_analysis" src="https://user-images.githubusercontent.com/85372441/124336717-ac5e3600-db64-11eb-8cd9-f97e955eb6e7.png">

## Summary 
The advantage of refactoring code is that it reduces the processing time and simplifies the code.
The disadvantage of refactoring code is that it takes some extra time and knowledge to set up an index with arrays.

The advantage of the original VBA script was that it was not as complex and did not require learning how to index. 
However the refactored VBA script had a few key advantages:
- Less processing time
- The ability to handle larger data sets
- Uses less memory
- Improving the logic of the code to make it easier for future users to read.

The disadvantage of the original VBA script is that its not efficient for larger data sets such as the entire stock market.
There were no disadvantages, in my opinion, with the refactored VBA script.
