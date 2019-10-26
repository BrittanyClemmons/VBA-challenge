Sub VBAStocks_Clemmmons()

    'Worksheet variables
    Dim WS_Count As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count
    '
    'Find the last row; Used "B" because "A" did not catch all rows
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    'Summary Table
    Dim Summary_Table_Row As Integer
    
    'Stocks variables
    Dim StockName As String
    Dim YearEnd As Double
    Dim YearStart As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    
    'Total Stock Volume variable
    Dim TSV As Double
    TSV = 0
    
    '--------------------------------------------------------------------------
    
    
    'Worksheet loop
    For i = 1 To WS_Count
    
        'Activate current worksheet
        Worksheets(i).Activate
        
        'Populate headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 14).Value = "Year End"
        Cells(1, 15).Value = "Year Start"
        
        'Set Summary Table to 2 each worksheet loop
        Summary_Table_Row = 2
        
        'Stocks loop
        
        If i = 7 Then
        
            'Sheet P has 261 rows per stock
            For k = 2 To LastRow
            
                If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
            
                    'Populate Ticker
                    StockName = Cells(k, 1).Value
                    Range("I" & Summary_Table_Row).Value = StockName
                    
                    'Populate Yearly Change
                    YearEnd = Cells(k, 6).Value
                    YearStart = Cells(k - 260, 3).Value
                    YearChange = YearEnd - YearStart
                    Range("N" & Summary_Table_Row).Value = YearEnd
                    Range("O" & Summary_Table_Row).Value = YearStart
                    Range("J" & Summary_Table_Row).Value = YearChange
                    
                    'Populate Percent Change
                    PercentChange = YearChange / YearStart
                    Range("K" & Summary_Table_Row).Value = PercentChange
                    
                    'Check Percent Change for Condition
                    If PercentChange >= 0 Then
                        Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                    Else
                        Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    'Populate Total Stock Volume
                    TSV = TSV + Cells(k, 7).Value
                    Range("L" & Summary_Table_Row).Value = TSV
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    TSV = 0
                    
                Else
                    TSV = TSV + Cells(k, 7).Value
                
                End If
            Next k
        
        Else
        
            'Sheets A through F have 262 rows per stock
            For j = 2 To LastRow
            
                If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
            
                    'Populate Ticker
                    StockName = Cells(j, 1).Value
                    Range("I" & Summary_Table_Row).Value = StockName
                    
                    'Populate Yearly Change
                    YearEnd = Cells(j, 6).Value
                    YearStart = Cells(j - 261, 3).Value
                    YearChange = YearEnd - YearStart
                    Range("N" & Summary_Table_Row).Value = YearEnd
                    Range("O" & Summary_Table_Row).Value = YearStart
                    Range("J" & Summary_Table_Row).Value = YearChange
                    
                    'Populate Percent Change
                    PercentChange = YearChange / YearStart
                    Range("K" & Summary_Table_Row).Value = PercentChange
                    
                    'Check Percent Change for Condition
                    If PercentChange >= 0 Then
                        Range("K" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                    Else
                        Range("K" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    'Populate Total Stock Volume
                    TSV = TSV + Cells(j, 7).Value
                    Range("L" & Summary_Table_Row).Value = TSV
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    TSV = 0
                    
                Else
                    TSV = TSV + Cells(j, 7).Value
                    
                End If
            Next j
        End If
        
        'Worksheet identification
        MsgBox ActiveWorkbook.Worksheets(i).Name
    
    Next i

    'Returns to first worksheet
    Worksheets("A").Activate

End Sub