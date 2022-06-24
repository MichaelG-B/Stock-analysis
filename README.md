# Stock Analysis with Excel VBA

## Overview of Project

- Link to project file
- https://github.com/MichaelG-B/Stock-analysis/blob/bd67d3f1078e1bb428f11d2a49fb341d00c3b4ed/VBA_Challenge.xlsm

### Purpose

- The purpose of this project was to refractor a previous Excel VBA code that we created to collect stock information on specific green energy stocks over multiple years in order to provide information to a client about how best they can diversify their investments. Refactoring will allow this code to applied with more flexibility and efficiency to better serve the clients interests as the develop over the future. 

### The Data

-The Data highlights 12 different stock options for 2017 and 2018. This data includes the name of the stock, the total daily volume of each stock as well as its return over the course of the year. This allows us to quickly and concisely view the performance of each stock over each year as well as determine which stocks did well for each year with goal of allowing us to determine smart diversification efforts based on the results of the data we analyzed.

## Results

- Here is our VBA code we developed to conduct our analysis of the stocks.

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
            tickerIndex = 0
    

    '1b) Create three output arrays
            Dim tickerVolumes(12) As Long
            Dim tickerStartingPrices(12) As Single
            Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
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
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerIndex = tickerIndex + 1
            
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

```

- As well as screen shots of our results for 2017 and 2018.

![Analysis 2017 Image](https://github.com/MichaelG-B/Stock-analysis/blob/bd67d3f1078e1bb428f11d2a49fb341d00c3b4ed/Analysis_2017.png)
![Analysis 2018 Image](https://github.com/MichaelG-B/Stock-analysis/blob/bd67d3f1078e1bb428f11d2a49fb341d00c3b4ed/Analysis_2018.png)

### Analysis

- Our analysis showed that for 2017 numerous stocks benefitted from a succesful year but the three prominent ones were DQ, SEDG and ENPH, which all provided over a 100% return rate for that year. Our Analysis showed that 2018 was not nearly as successful as the previous year with many of the stocks showing negative returns, with the exception of ENPH and RUN which roughly returned the same rate at 80%. Bases on our analysis a decent strategy would be to hold stock in ENPH, SEDG, and RUN to minimize volatility and maintain sustainable returns.


## Summary

### Pros and Cons of Refactoring Code

- In general some pros of refactoring code are an increase in efficiency and an increase in usability. Our code becomes more efficient due to it being written in a format that is usually more concise and streamlined by really focusing on the goal it is going to be used for. Our Code becomes more usable due to it being written in a manner that is easier to understand for someone who may not have all the background information about why it was written in the first place, the goal is for anyone to be able to understand its purpose generally by merely glancing over it. Some cons of refactoring code include obstacles that may prevent it from being possible, such as a certain code being to specific for a given task or dataset limiting its ability to become more generalized. Another obstacle may be in the ability to test the code given for instance time constraints in the real world where time is of the essence.

### The Advantages of Refactoring Stock Analysis 

-The main advantage of refactoring this specific code that we created for our client in order to analyze different green stocks to help them determine how best they should diversify their investment portfolio, is the drastic decrease in run time from .777 to .144 seconds in 2017 and .683 to .140 seconds in 2018. This means our new refactored code will be able to take on larger amounts of data and establish a result in a quicker time, which is invaluable in a real life setting where investments have the abilitiy to be affected by the hour.

![VBA Challenge 2017 Image](https://github.com/MichaelG-B/Stock-analysis/blob/bd67d3f1078e1bb428f11d2a49fb341d00c3b4ed/VBA_Challenge_2017.png)
![VBA Challenge 2018 Image](https://github.com/MichaelG-B/Stock-analysis/blob/bd67d3f1078e1bb428f11d2a49fb341d00c3b4ed/VBA_Challenge_2018.png)
