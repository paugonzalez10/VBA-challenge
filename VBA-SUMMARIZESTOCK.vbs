Attribute VB_Name = "Module1"
Sub SummarizeStocks()

    '--- Define variables ---
    Dim Ticker As String
    Dim StockOpen, StockClose, YearlyChange, PercentChange As Double
    Dim i, LastRow, OutputRow As Double
    Dim Volume, TotalVolume As Double
    
    Dim IncreaseTicker, DecreaseTicker, VolumeTicker As String
    Dim GreatestIncrease, GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    Dim ws As Worksheet
    
    '--- Loop through each worksheet ---
    For Each ws In Worksheets
      
        '--- Insert summary headers into worksheet ---
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        '--- Find the last row in the worksheet and assign to LastRow ---
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        '--- Assign initial values to variables ---
        Ticker = ws.Range("A2").Value
        StockOpen = ws.Cells(2, 3).Value
        
        StockClose = 0
        PercentChange = 0
        YearlyChange = 0
        TotalVolume = 0
        
        IncreaseTicker = " "
        DecreaseTicker = " "
        VolumeTicker = " "
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
    
        '--- Assign first output row for the summary table to be on row 2 ---
        OutputRow = 2
    
        '--- For loop to go through data from row 2 to the last row  ---
        For i = 2 To LastRow
                   
            '--- Conditional statement to check whether the next row contains the same ticker value i.e. the current ticker is continuing ---
            If Ticker = ws.Cells(i + 1, 1).Value Then
            
                '--- Assign a value to volume from worksheet ---
                Volume = ws.Cells(i, 7).Value
                
                '--- Calculate total stock volume ---
                TotalVolume = TotalVolume + Volume
            
            '--- Else statement when the next row's ticker value is NOT the same i.e. the current ticker has ended ---
            Else
        
                '--- Assign volume and closing price from this row (end of year value) from worksheet ---
                StockClose = ws.Cells(i, 6).Value
                Volume = ws.Cells(i, 7).Value
        
                '--- Calculate yearly change ---
                YearlyChange = StockClose - StockOpen
                
                '--- Check whether the stock open price is zero (for division) ---
                If StockOpen <> 0 Then
                
                    '--- Calculate percent change ---
                    PercentChange = (YearlyChange / StockOpen)
                    
                Else
                    MsgBox ("Percent change error: There is an opening price of 0! Please check row" + Str(i) + " in ticker: " + Ticker)
                    PercentChange = 0       'Set a default value if it stock open price is 0 to override error
                
                End If
                
                '--- Calculate total stock volume ---
                TotalVolume = TotalVolume + Volume
            
                '--- Insert calculated values into summary table on worksheet ---
                ws.Cells(OutputRow, 9).Value = Ticker
                ws.Cells(OutputRow, 10).Value = YearlyChange
                ws.Cells(OutputRow, 11).Value = PercentChange
                ws.Cells(OutputRow, 12).Value = TotalVolume
                        
                '--- if YearlyChange is positive, change the cell color to green ---
                If YearlyChange >= 0 Then
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = 4
                
                '--- if YearlyChange is negative, change cell color to red ---
                Else
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = 3
                    
                '--- end if statements for conditional formatting ---
                End If
          
                '--- Nested If statements to check greatest % increase, decrease, and volume ---
                If PercentChange > GreatestIncrease Then
                    IncreaseTicker = Ticker
                    GreatestIncrease = PercentChange
            
                ElseIf PercentChange < GreatestDecrease Then
                    DecreaseTicker = Ticker
                    GreatestDecrease = PercentChange
                
                End If
            
                If TotalVolume > GreatestVolume Then
                    VolumeTicker = Ticker
                    GreatestVolume = TotalVolume
                    
                End If
            
                '--- Reassign values for the next ticker (new ticker value, new open price and reset total volume) ---
                Ticker = ws.Cells(i + 1, 1).Value
                StockOpen = ws.Cells(i + 1, 3).Value
                StockClose = 0
                TotalVolume = 0
                
                '--- Update OutputRow so that info for new ticker is entered onto next row in the summary table ---
                OutputRow = OutputRow + 1
            
                '--- Enter calculated greatest values onto worksheet ---
                ws.Range("P2").Value = IncreaseTicker
                ws.Range("Q2").Value = GreatestIncrease
                ws.Range("P3").Value = DecreaseTicker
                ws.Range("Q3").Value = GreatestDecrease
                ws.Range("P4").Value = VolumeTicker
                ws.Range("Q4").Value = GreatestVolume
                                   
                '--- Format percent values to percentages ---
                ws.Columns("K").NumberFormat = "0.00%"
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("Q3").NumberFormat = "0.00%"
                
                '--- Format yearly change to currency ---
                ws.Columns("J").NumberFormat = "$#,##0.00"
                
                '--- Format cells to autofit contents --
                ws.Columns("I:Q").AutoFit
                
            '--- End conditional statements ---
            End If
        
        '--- Continue for loop to next row until the last row ---
        Next i
    
    '--- Continue to next worksheet ---
    Next ws
    
End Sub
