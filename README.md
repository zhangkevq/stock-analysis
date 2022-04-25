# Stock Analysis
  This project seeks to demonstrate the effects of efficient coding and show their benefits using VBA. Images will be provided which will show the results of the differences between the original script we ended up with after performing analysis on the data.

## Data
  This project used stock information in two sheets of an Excel file. The data has information on 12 different stocks, including their ticker value, issue dates, opening/closing and adjusted closing price, highest/lowest price, and volume of the stock. The _All Stock Analysis_ function's purpose is to get the ticker, the ticker's total daily volume, and the return-percentage on each stock. The aim of refactoring the code is to reduce the time it takes for the program to perform the task. Instead of a nested for loop, the code was changed to use one loop. It cleans up some of the unnecessary processes in the code, making it run more efficiently, as well.

## Refactoring
The original stocks analysis code can be found here: [Green Stocks](https://github.com/zhangkevq/stock-analysis/blob/main/green_stocks.xlsm)  
```
Sub allAnalysis()
    'format output on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Dim startTime As Single
    Dim endTime As Single
    
    'prompt user for year
    yearValue = InputBox("What year would you like to run the analysis on?")
    'start of timer
    startTime = Timer
    
    'run analysis for ANY year
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'initialize array of all tickers
    Dim tickers(11) As String
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
    
    'initialize variables for starting price and ending price
    Dim startPrice As Single
    Dim endPrice As Single

    'activate DATA worksheet
    Sheets(yearValue).Activate
    
    'Get number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop through tickers
    For i = 0 To 11
        'a line of code here will run 12 times
        ticker = tickers(i)
        totalVolume = 0
        'loop through rows in data
        Sheets(yearValue).Activate
        For j = 2 To RowCount
            'find total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            'find starting price for current ticker
            If Cells(j - 1, 1) <> ticker And Cells(j, 1).Value = ticker Then
                startPrice = Cells(j, 6).Value
            End If
            'find ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endPrice = Cells(j, 6).Value
            End If
        Next j
    'output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endPrice / startPrice - 1
    Next i
    
    'end of timer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub
```
When running the VBA script, the run-time of the program is recorded. It takes a consistent ~1.68 seconds for both sheets.  
![2017 Analysis](https://github.com/zhangkevq/stock-analysis/blob/main/runtime_2017_analysis1.png) ![2018 Analysis](https://github.com/zhangkevq/stock-analysis/blob/main/runtime_2018_analysis1.png)  

The refactored code will have header formatting and color-coding included, and those were also in the original stock analysis code. Below is the refactored analysis code:

#### Refactored All Stocks Analysis Code
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
    Dim tickerIndex as Single
    tickerIndex = 0

    '1b) Create three output arrays   
    dim tickerVolumes(12) as Long
    dim tickerStartingPrices(12) as Single
    dim tickerEndingPrices(12) as Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    for i = 0 to 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    next i

    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i,1).Value = tickers(tickerIndex) AND Cells(i-1,1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i,6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i,1).Value = tickers(tickerIndex) AND Cells(i+1,1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i,6).Value

        '3d Increase the tickerIndex. 
                tickerIndex = tickerIndex + 1
        'End If
            End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i)/tickerStartingPrices(i) - 1
        
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
[Refactored Code](https://github.com/zhangkevq/stock-analysis/blob/main/VBA_Challenge.vbs)  

## Conclusions

### Pros and Cons

  Using this new code, we can test the run speed and compare to the above speed. The runtimes between the two sets of data are different now:  
![2017 refactored](https://github.com/zhangkevq/stock-analysis/blob/main/runtime_2017_analysis2.png) ![2018 refactored](https://github.com/zhangkevq/stock-analysis/blob/main/runtime_2018_analysis2.png)  

As can be seen, there is a significant time save when using this new refactored program, which is a big advantage. However, this analysis may not always be possible. There maybe times where data is unavailable to pull from the table and the program can get stuck in the refactored version but not in the original version. There is also an inconsistency between the run time between 2017 and 2018, which means that in certain situations, there could be a wider run-time variance. For example, with an extremely large data set, it could take longer than the original function for some sheets, and faster for other sheets which will lead to user frustration.
