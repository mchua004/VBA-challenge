Attribute VB_Name = "Module1"
Sub stonks()
    'Create variables for worksheets
    Dim S1 As Worksheet
    Set S1 = Worksheets("2016")
    Dim S2 As Worksheet
    Set S2 = Worksheets("2015")
    Dim S3 As Worksheet
    Set S3 = Worksheets("2014")
    
    'Create variables for ticker symbol, yearly price change, yearly percent change, and total stock volume
    Dim ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
        
    Dim StockVolume As LongLong
    StockVolume = 0
    
    Dim OP As Double
    OP = 0
    
    Dim CP As Double
    
    Dim i As Long
    
    Dim j As Long
    j = 2
    
    Dim LastRow As Long
    
    S1.Cells(1, 9).Value = "Ticker"
    S1.Cells(1, 10).Value = "Yearly Change"
    S1.Cells(1, 11).Value = "Percent Change"
    S1.Cells(1, 12).Value = "Total Stock Volume"
    
    S2.Cells(1, 9).Value = "Ticker"
    S2.Cells(1, 10).Value = "Yearly Change"
    S2.Cells(1, 11).Value = "Percent Change"
    S2.Cells(1, 12).Value = "Total Stock Volume"
    
    S3.Cells(1, 9).Value = "Ticker"
    S3.Cells(1, 10).Value = "Yearly Change"
    S3.Cells(1, 11).Value = "Percent Change"
    S3.Cells(1, 12).Value = "Total Stock Volume"
    
    'Define the last row of the sheet
    LastRow = S1.Cells(S1.Rows.Count, "A").End(xlUp).Row
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        If OP = 0 Then
            
            OP = S1.Cells(i, 3).Value
            
        End If
        
        If S1.Cells(i, 1).Value <> S1.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = S1.Cells(i, 1).Value
            
            'Print the ticker symbol
            S1.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = S1.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            S1.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            S1.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + S1.Cells(i, 7).Value
            
            S1.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + S1.Cells(i, 7).Value
                                    
        End If
        
        If S1.Cells(j, 10).Value > 0 Then
        
            S1.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf S1.Cells(j, 10).Value < 0 Then
            
            S1.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = S2.Cells(S2.Rows.Count, "B").End(xlUp).Row
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        
        If OP = 0 Then
            
            OP = S2.Cells(i, 3).Value
            
        End If
        
        If S2.Cells(i, 1).Value <> S2.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = S2.Cells(i, 1).Value
            
            'Print the ticker symbol
            S2.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = S2.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            S2.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            S2.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + S2.Cells(i, 7).Value
            
            S2.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + S2.Cells(i, 7).Value
                                    
        End If
        
        If S2.Cells(j, 10).Value > 0 Then
        
            S2.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf S2.Cells(j, 10).Value < 0 Then
            
            S2.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = S3.Cells(S3.Rows.Count, "C").End(xlUp).Row
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        
        If OP = 0 Then
            
            OP = S3.Cells(i, 3).Value
            
        End If
        
        If S3.Cells(i, 1).Value <> S3.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = S3.Cells(i, 1).Value
            
            'Print the ticker symbol
            S3.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = S3.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            S3.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            If YearlyChange = 0 And OP = 0 Then
            
                PercentChange = 0
                
            Else
                
                PercentChange = YearlyChange / OP
            
            End If
            
            'Print the percent change
            S3.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + S3.Cells(i, 7).Value
            
            S3.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + S3.Cells(i, 7).Value
                                    
        End If
        
        If S3.Cells(j, 10).Value > 0 Then
        
            S3.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf S3.Cells(j, 10).Value < 0 Then
            
            S3.Cells(j, 10).Interior.ColorIndex = 3
        
        Else
                    
            S3.Cells(j, 10).Interior.ColorIndex = 2
            
        End If
        
    Next i
    
End Sub
