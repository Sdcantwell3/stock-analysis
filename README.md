# Green Stock Analysis
## Overview
### Purpose
  The purpose of this project is to look at functioning VBA code and attempt to refactor it into a more efficient form.  The code is built to analyse stock data.  In the model we looked at 12 different Green stocks annual performance for 2017 and 2108.
## Results
### Analysis
  I built the refactored code on the structure of the original analysis keeping the code for the input box and headers. I also used the original code for the ticker array. Below are the step by step changes I took to refactor the orginal code.
       
       '1a) Create a ticker Index
        tickerIndex = 0
       
        '1b) Create three output arrays      
        Dim tickerVolumes(12) As Long   
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
        ''2a) Create a for loop to initialize the tickerVolumes to zero.  
        For i = 0 To 11
             tickerVolumes(i) = 0   
        Next i
        
        ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
         
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
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
## Summary
### Advantages and Disadvantages of Refactoring
  Refactoring code on a smaller subset of data that parallels the structure of the larger data set gives us the opportunity to efficiently analyse abnormally large data sets, that would be out of reach with less efficient code.  If taken too far you can easily loose functionality. It is also very important be certain that the data you are modeling on is close enough to the bigger data set to comfortable that the code will perform without error.
### The Refactored Stock Analysis
  The adjustments we made to the code definitely sped up the run time for the analysis. The run time was approximately 5% faster for the refactored code. Below are to screen grabs showing the run times for each year using the refactored code.
  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/104606589/169217534-c222994f-7852-4586-86c3-124ea5a0dbcf.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/104606589/169217544-0dbe0685-f448-4ca3-ae3a-29c9c05c3d63.png)

  

