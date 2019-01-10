Sub StockAnalysisHard()
    
    'Declare variable name of ws as a worksheet type
    Dim ws As Worksheet
    
    
    'START OF LOOP THROUGH EACH WORKSHEET IN WORKBOOK
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    
        'Set intial variables that hold running data to be displayed in display table when ready
        'Ticker Symbol, Ticker Volume, YearOpenPrice, YearClosePrice
        Dim TickerSymbol As String
        Dim TickerVolume As Double
        TickerVolume = 0
        Dim YearOpenPrice As Double
        YearOpenPrice = Cells(2, 3).Value
        Dim YearClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
       'Set Headers for summary table
       Range("I1").Value = "Ticker"
       Range("J1").Value = "Yearly Change"
       Range("K1").Value = "Percent Change"
       Range("L1").Value = "Volume"
        
        'Tracker for the location for each ticker in the summary table
        Dim SummaryTableRow As Double
        SummaryTableRow = 2
        
        'Find last row on each sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Start of loop through each row
        For i = 2 To LastRow
            
            'Comparing current row to next row
            'Checks if we are still within the same ticker symbol
            'if next row is different than current row (which means we reached new ticker, then...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'Lock in the current ticker symbol
                'Lock in current closing price (which means end of year close price)
                TickerSymbol = Cells(i, 1).Value
                YearClosePrice = Cells(i, 6).Value
                
                'Add up current ticker symbol volumes
                TickerVolume = TickerVolume + Cells(i, 7).Value
                
                'Calculate yearly change in price from open to close
                YearlyChange = YearClosePrice - YearOpenPrice
                
                'Calculate yearly percent change of price
                If (YearOpenPrice = 0 And YearClosePrice = 0) Then
                
                PercentChange = 0
                
                ElseIf (YearOpenPrice = 0 And YearClosePrice > 0) Then
                
                PercentChange = YearlyChange / 100
                
                Else
                PercentChange = YearlyChange / YearOpenPrice
                
                End If
                
                
                'Print the current ticker symbol to summary table
                'Print the yearly change to the summary table
                Range("I" & SummaryTableRow).Value = TickerSymbol
                Range("J" & SummaryTableRow).Value = YearlyChange
                'Format year change green or red
                If Range("J" & SummaryTableRow).Value >= 0 Then
                
                    Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                Else
                    Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    
                End If
                
                'Print the percent change to the summary table
                Range("K" & SummaryTableRow).Value = PercentChange
                'Format percent change as a percentage
                Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                'Print the curent ticker symbol volume total to summary table
                Range("L" & SummaryTableRow).Value = TickerVolume
                
                'Add one row to the summary table
                SummaryTableRow = SummaryTableRow + 1
                
                'Reset the ticker volume counter
                'Reset the year open price
                TickerVolume = 0
                YearOpenPrice = Cells(i + 1, 3).Value
                
                'If the ticker symbols in comparisson match
                'which means we are still within same ticker symbol
                Else
                
                    'Holder to keep running total of volume for the same ticker symbols
                    TickerVolume = TickerVolume + Cells(i, 7).Value
                    
                    
            
            'End of ticker check
            End If
        
        'End of row loop and starting of next row loop
        Next i
        
        
        'START OF ANALYZING SUMMARY TABLE TO FIND GREATEST INCREASE, DECREASE, VOLUME
        
       'Set variables for greatest table
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        Dim SummaryLastRow As Double
        Dim RangePercent As Range
        Dim RangeVolume As Range
    
        'find last row of summary table
        SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set Greatest Table Headers
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Set range of yearly percents to check
        Set RangePercent = Range("K2:K" & SummaryLastRow)
        Set RangeVolume = Range("L2:L" & SummaryLastRow)
            
        'Built in function to find min and max of a range
        GreatestDecrease = WorksheetFunction.Min(RangePercent)
        GreatestVolume = WorksheetFunction.Max(RangeVolume)
        GreatestIncrease = WorksheetFunction.Max(RangePercent)
        
        'Start of summary table loop through each row on table
        For j = 2 To SummaryLastRow
        
            'If current row percent change is indeed the max, then do the following
            If Cells(j, 11).Value = GreatestIncrease Then
            
                'Print current row value to greatest table
                Range("P2").Value = Cells(j, 9).Value
                Range("Q2").Value = Cells(j, 11).Value
                Range("Q2").NumberFormat = "0.00%"
                
                'If current row percent change is indeed the min, then do the following
                ElseIf Cells(j, 11).Value = GreatestDecrease Then
                
                    'Print current fow value to greatest table
                    Range("P3").Value = Cells(j, 9).Value
                    Range("Q3").Value = Cells(j, 11).Value
                    Range("Q3").NumberFormat = "0.00%"
                    
                'If current row volume is indeed he max, then do the following
                ElseIf Cells(j, 12).Value = GreatestVolume Then
                
                    'Print the current row value to greatest table
                    Range("P4").Value = Cells(j, 9).Value
                    Range("Q4").Value = Cells(j, 12).Value
            
            End If
            
        'End of first row analysis and onto next row
        Next j
        
    'END OF WORKSHEET LOOP and moving onto next worksheet loop
    Next ws
    
End Sub