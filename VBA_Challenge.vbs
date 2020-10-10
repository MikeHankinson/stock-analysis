Attribute VB_Name = "Module3"
Sub AllStocksAnalysisRefactored()
    
   '********Variable Declarations*********
    
    Dim startTime As Single
    Dim endTime  As Single
    YearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
    Dim tickers(12) As String        'Initialize array of all tickers (required array 1) -- all 4 arrays total (15 pts)
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
               
        Dim tickerVolumes(12) As Long                   '1b) Create three output arrays (required array 2)
        
        Dim tickerStartingPrices(12) As Single          '(required array 3)
        Dim tickerEndingPrices(12) As Single            '(required array 4)
        tickerIndex = 0                                 '1a) Create a ticker Index and set equal to 0 before itterating. (5 points)
 
    
   '*******Set Title and Row header in "All Stocks Analysis" tab*******
   
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + YearValue + ")"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

   
   
    '*******For Loops to search data and populate analysis*******
    
        Worksheets(YearValue).Activate                      'Activate data worksheet
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row     'Get the number of rows to loop over
    
        For n = 0 To 11                                     '2a) Create a for loop to initialize the tickerVolumes to zero.
            tickerVolumes(n) = 0
        Next n
        
        For i = 0 To 11                                    'Ticker Loop
            
            ticker = tickers(i)
            
      
            

            For j = 2 To RowCount                           ''2b) Loop over all the rows in the spreadsheet.
            
            
                If Cells(j, 1).Value = ticker Then                                                           '3a) Increase volume for current ticker. Ticker index Usage(5 points)
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                End If
                
                
                If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then  '3b) Check if the current row is the first row with the selected tickerIndex. Ticker index Usage(5 points)
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                End If
                
                If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then   '3c) check if the current row is the last row with the selected ticker. Ticker index Usage(5 points)
                    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                End If
            
            
               If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then  '3d Increase the tickerIndex.  Ticker index Usage
                    tickerIndex = tickerIndex + 1
                End If
        
    
        
           
            Next j
    
   Next i
    
        Worksheets("All Stocks Analysis").Activate          '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        
        For i = 0 To 11
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = (tickerEndingPrices(i) - tickerStartingPrices(i)) / tickerStartingPrices(i)
        Next i
  
    
    '*******Text Formatting*******
    
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (YearValue)

    
    

End Sub


Sub test()

Worksheets("test").Activate

Dim tickers(3) As String        'Initialize array of all tickers
        tickers(0) = "a"
        tickers(1) = "b"
        tickers(2) = "c"
        
        tickerIndex = 0                                 '1a) Create a ticker Index and set equal to 0 before itterating.
Dim tickerVolumes(3) As Long


        For n = 0 To 2
            tickerVolumes(n) = 0
        Next n
        
            
Dim tickerStartingPrices(3) As Single

Dim tickerEndingPrices(3) As Single
        
    
    For i = 0 To 2                                    'Ticker Loop
            ticker = tickers(i)
                                      '2a) Create a for loop to initialize the tickerVolumes to zero.
            

            For j = 2 To 10                           ''2b) Loop over all the rows in the spreadsheet.
                
                
                If Cells(j, 1).Value = ticker Then                                      '3a) Increase volume for current ticker.
                    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                End If
                
                
                If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then  '3b) Check if the current row is the first row with the selected tickerIndex. Defines starting price for ith ticker
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                End If
            
                If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then  '3b) Check if the current row is the first row with the selected tickerIndex. Defines starting price for ith ticker
                    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                End If
            
            
            
            
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then  ''2a)If the next row's ticker doesn't match, increase the tickerIndex.
                    tickerIndex = tickerIndex + 1
                End If
            
            Next j
            
            
        Worksheets("test").Activate
        Cells(15 + i, 1).Value = ticker
        Cells(15 + i, 2).Value = totalVolume
        'Cells(15 + i, 3).Value = ((endingPrice - startingPrice) / startingPrice)
            
            
    Next i
    
Cells(15, 8).Value = tickerVolumes(0)
Cells(16, 8).Value = tickerVolumes(1)
Cells(17, 8).Value = tickerVolumes(2)
Cells(18, 8).Value = tickerStartingPrices(0)
Cells(19, 8).Value = tickerStartingPrices(1)
Cells(20, 8).Value = tickerStartingPrices(2)
Cells(21, 8).Value = tickerEndingPrices(0)
Cells(22, 8).Value = tickerEndingPrices(1)
Cells(23, 8).Value = tickerEndingPrices(2)



MsgBox (tickers(0) & " " & tickerIndex & " " & tickerVolumes(0))
End Sub
