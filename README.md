# stock-analysis

## Overview of Project 
A workbook was prepared for Steve that analyzed 12 different stocks. He was very happy with the end result but wanted to do more research and expand his dataset to include the entire stock market over the last couple years. In order to do this efficiently, we need to refactor the code. 

### Purpose
The purpose of this project is to take our original code and refactor it so the code can loop through all the data one time and collect the same information it did before. We also want to determine if refactoring the code made the VBA script run faster. 

## Results 
The first thing I had to do was create a ticker index and set it to zero before going through all the data. I also had to create three different outputs in order to analyze our data - volume, starting price and ending price. 

![image](https://user-images.githubusercontent.com/117782103/204052808-3f9e5d74-c5a8-4e7a-b5a9-c16403510d9d.png)

A for loop was then created to set all the volumes to zero. Another for loop was created to loop over all of the rows in the spreadsheet. Inside that loop we also increased the ticker volume and added it for the current stock ticker. 

![image](https://user-images.githubusercontent.com/117782103/204052996-b13260b1-4e28-4720-9c69-55f8d0cab76c.png)

Conditionals were then created to determine if the current row was the first or last row of the ticker index. This allows us to be able to assign the current starting and ending ticker prices. If the ticker index did not match the previous rows ticker, we increased the ticker index so we could analyze all tickers. 

![image](https://user-images.githubusercontent.com/117782103/204053142-b93dde2b-2bd2-4621-8a94-3253fe927c93.png)

Then we had to loop through our arrays of tickers, volume, starting and ending price so we could find the output data which included the ticker, total daily volume and return. 

![image](https://user-images.githubusercontent.com/117782103/204053240-ef523aea-2677-4b0d-91f2-9de4ecc13d27.png)

Lastly, the data was formatted to make it easy to read and a pop-up message was created to show the run time for the code. It was determined that by refactoring the code, the run time significantly decreased. In the original script, the run time was 1.078 and 1.074 Once it was refactored, the run time was cut down to 0.152 and 0.164 seconds. Below shows the screenshots of the 2017 original and refactored code. 

![image](https://user-images.githubusercontent.com/117782103/204053554-352e4a9e-11cf-4c06-be8e-c6b6408ca280.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/117782103/204053569-6c50d2c0-e5e4-4a53-95f8-6dc2bb07c957.png)

## Summary

### Advantages and Disadvantages of Refactoring Code 
Advantages of refactoring code include improvement in the efficiency, format and design of code. 
Disadvantages include the time and effort it takes to refactor. You also have to have the time to test all of your new code. Refactoring can lead to additional errors and bugs that need to be fixed. 

### Refactoring the Original VBA Script 
The advantages and disadvantages of refactoring code in general also seemed to be the same pros and cons that were encountered during refactoring this code. The biggest advantage was code run time. We were able to significantly decrease the code run time which made the code much more efficient. 

The disadvantage was the time it took to refactor. I encountered a few errors when testing so had to spend more time to debug the data to ensure we would receive the correct results. 

Overall, it was still advantageous to re-factor this code so Steve can analyze a larger set of data efficiently. 
