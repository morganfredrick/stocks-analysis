# Stock Analysis VBA in Excel 

## Purpose
### Project Overview
The purpose of the project was to compare stocks analysis for years 2017 and 2018, presenting them in a way that an informed decision on which stocks would be worth investing in. Our goal with the adjustment of the refractored code was to see if we could make run times faster than the initial, similar, code used.

The data presented two charts with various information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the highest and lowest price, along with the volume of the stock, and the opening, closing, and adjusted closing price. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.

## Results
### Analysis
I began by importing the refactored code and editing it to update ticker volumes, ticker ending prices and beginning prices. I then activated the accurate worksheet in order to gather the correct information and manipulated the ticker to attempt to make the code run faster with the refactored code, than the original code, as you can see below: 

 '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
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
    
## Summary
### Pros and Cons of Refactoring Code
As a pro, refactoring can make code cleaner and more organized, by design and software improvements, debugging, and sometimes faster programming. It also benefits other users who view our coding because it becomes easier to read, as we have made it more concise. 

However, cons where we don’t have the option to refactor our own code may cause issues such as applications that are too large or the improper tests for the existing codes, which may ultimately cause risks if we try to refactor the code.

Our refactored code did not run faster. As you can see below, the refactored code for the 2017 analysis made the macro run slightly slower, at 0.25 seconds:

![VBA_green_stocks_2017](https://user-images.githubusercontent.com/104293158/171071658-acbfb535-106e-40fd-91ca-8de3790dc5b1.png)

Than the original code, at 0.0234375 seconds:

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/104293158/171071662-c7f13766-c808-4d0f-945e-6ebaa45e7cae.PNG)


The refactoring of the macro also caused the 2018 code to run slightly slower, at 0.1953125 seconds:

![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/104293158/171085292-4ae84c2e-0552-4c47-85e3-479ea7e29ae4.PNG)

Than the original, at 0.015625 seconds:

![VBA_green_stocks_2018](https://user-images.githubusercontent.com/104293158/171085261-3a4d2490-ba05-44f7-b7d4-10a1ebd106e4.png)


