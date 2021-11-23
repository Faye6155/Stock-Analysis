# Stock-Analysis
Module 2-VBA-Stock Analysis
Stock Analysis

Overview
This project was to refacter the MS VBA code to collect stock information for 2017 and 2018.The data collected would help the user to determine if the stocks are good to invest. The goal of the project is to increase the efficiency of the orginal code from the module.

Analysis
After applying the code to collect data from the origin data workshhets, the 12 different stocks,ticker value and return on each stock would be pointed out in the worksheet “All Stocks Analysis"
The timer setting is to show the time used to run the code. It would indicate how efficient this project works.

The code applied is following:

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
    
    For r = 0 To 11
    tickerVolumes(r) = 0
    tickerStartingPrices(r) = 0
    tickerEndingPrices(r) = 0
    Next r
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For e = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(e, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(e, 1).Value = tickers(tickerIndex) And Cells(e - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(e, 6).Value
            End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row鈥檚 ticker doesn鈥檛 match, increase the tickerIndex.
        'If  Then
            
             If Cells(e, 1).Value = tickers(tickerIndex) And Cells(e + 1, 1).Value <> tickers(tickerIndex) Then
             tickerEndingPrices(tickerIndex) = Cells(e, 6).Value
             End If

            '3d Increase the tickerIndex.
             
            If Cells(e, 1).Value = tickers(tickerIndex) And Cells(e + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If

        'End If
    
    Next e
    
       Worksheets("All Stocks Analysis").Activate
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
         Cells(4 + i, 1).Value = tickers(i)
         Cells(4 + i, 2).Value = tickerVolumes(i)
         Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
      
    'Formatting
 
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


Summary

Advantages of refactoring code
- The code would be more organized. It is easy and clear for people to review the code and help the software improvement and debugging. 
- By comparing with the original code, it is easy to add the missed code.

Disadvantages of refactoring code
- Refactoring code might cause the miscoding for the variables. Same output worksheet used might have the overlapped info from the original code.

Advantages of refactored VBA Script
The screenshot shows the code ran both 2018 and 2017 data under 0.25 second. Very little time used to get the outcomes. The origin code takes around 1 second to run the result. This VBA Script shows the efficiency approved.
