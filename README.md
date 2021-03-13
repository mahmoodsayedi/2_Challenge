# Purpose
In this project and analyisis, I edit, or refactored, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataset. Then, I will determine whether refactoring this code successfully made the VBA script run faster. Finally, I just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
# Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
## 1. The tickerIndex is set equal to zero before looping over the rows.
Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the different arrays on VBA Code: the tickers array and the the output arrays created on next requierement. ![image](https://user-images.githubusercontent.com/79608027/111041679-62abdd00-8407-11eb-8bd2-7ea922117eb6.png) 

## 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type. ![image](https://user-images.githubusercontent.com/79608027/111041746-c6cea100-8407-11eb-86a3-97e885ecd336.png)

## 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex ![image](https://user-images.githubusercontent.com/79608027/111041801-11e8b400-8408-11eb-84c9-760f0f96d70e.png)
## 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. ![image](https://user-images.githubusercontent.com/79608027/111041859-5e33f400-8408-11eb-8dfa-157326f873c8.png)

## 5. There are comments to explain the purpose of the code.
Adding Comments is requiered, as a Best Practices for Writing Super Readable Code such ![image](https://user-images.githubusercontent.com/79608027/111041919-ab17ca80-8408-11eb-8569-b80010ae7711.png)
## 7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module
Finally, when run the stock analysis, to confirm that the stock analysis outputs for 2017 and 2018 are the same as dataset example provided  In adition, in my resources folder and below you can see the final Stock Analysis Results named, Final VBA Analysis 2017 and 2018 save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. ![image](https://user-images.githubusercontent.com/79608027/111042064-680a2700-8409-11eb-9675-b51384a54b30.png) ![image](https://user-images.githubusercontent.com/79608027/111042075-7b1cf700-8409-11eb-8723-6f8b540a5751.png)
# SUMMARY: Our Statement:
## Deliverable with detail analysis:
### 1. What are the advantages or disadvantages of refactoring code?

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

##### Disadvantages:

###### A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
###### A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called   from the other functions.
####### A complex unstructured code is usually best to split in several functions.
####### Refactoring process can affect the testing outcomes.

##### Advantages:

Logical errors easily appear in well structure code that contains nested conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.
### 2. How do these pros and cons apply to refactoring the original VBA script?

Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. Now, let's think about something, What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand? If yes then definitely you didn’t pay attention to improve your code or to restructure your code.

We need to consider the code refactoring process as cleaning up the orderly house. Unnecessary clutter in a home can create a chaotic and stressful environment. - The same goes for written code.

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.
