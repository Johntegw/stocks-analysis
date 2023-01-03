# Title: 
    # Stock analysis_Challenge


# Decription
 #Stock Analysis on expanded years


# Overview of Project
The object of this project is to help Steve analyze stockes over a period of time by making it easy and faster to gather the information.

# The purpose and background are well defined 
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, we will edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then determine whether refactoring your code successfully made the VBA script run faster. 

# Results
Before the refactoring of the already established code, there were run time errors that didn't allow the code to run properly. (see exhibit A) So looking at the code and making a few changes, I was able to get the code to run and process the analysis of the stockes between two years. (see screen shots)  For YR 2017 the process time went from having errors to .09375 seconds. The stock volumes and returns were as follows: 
            The stock with the lowest volumes in 2017 ended up having the highest return. That was ticker DQ with 35,796,200 in vol. and return was 199.4%
            The ticker with the lowest return was ticker TERP at -7.2%
![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/119356389/210444635-681d846e-2287-411f-bb7f-99f1c7149b98.png)
![Stock Ticker Prices 2017](https://user-images.githubusercontent.com/119356389/210444690-657e8fc5-233e-4d5e-be4f-fa45dfcac4f7.png)



The processing time for year 2018 was .109375; stock volumes and returns for YR 2018 were as follows: 
            The stock with the highest return was RUN at 84% while volumes were only 502,757,100; and the lowest returns were from stock DQ at 162.6% with 107,873,900 in volumes 
![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/119356389/210444809-7f607cf0-c77f-49dc-812f-03ae7cc29582.png)
![Stock Ticker Prices 2018](https://user-images.githubusercontent.com/119356389/210444855-87ac8a69-39eb-4d84-bc67-2bc9581a5013.png)



Based on these data sets for both 2017 and 2018 years, is seems that the returns were not predicated on the volumes of the stock.  Following is the refactor code I was able to use to get the outcomes for the challenge:
    '1a) Create a ticker Index
            Dim tickerIndex As Integer
            tickerIndex = 0
            

    '1b) Create three output arrays
    
     Dim tickerStartingPrices(12), tickerEndingPrices(12) As Single
     Dim tickerVolumes(12) As Long
     
     
            
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    Dim i As Integer
    
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrice = 0
        tickerEndigPrice = 0
        
      
        
    Next i
     
    
             '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
                     '3a) Increase volume for current ticker
     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
     
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
         
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
                       

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
              tickerIndex = tickerIndex + 1
            End If
            

    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
         Cells(4 + i, 1).Value = tickers(i)
         Cells(4 + i, 2).Value = tickerVolumes(i)
         Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i


# Summary
# There is a detailed statement on the advantages and disadvantages of refactoring code in general.

While the challenge was a good way to learn how others would code for a specific outcome, it was hard to step into the code to try to come up with what will work for you and your understanding. Debugging the code helped with helping you find the where the errors were and how to solve them.  Advantages to this is that you already have a foundation in which to build upon. The disadvantage is filling  in the gaps and finding what will help the code run and easier to be read by other. 

# There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script.
The advantage of using the original code is that you don't have to necessarily start from scratch. It gives you a foundations in which to start with re imagining a cleaner code.The disadvantage is making sure that the code works with your ideas and that you have all the right information. 
