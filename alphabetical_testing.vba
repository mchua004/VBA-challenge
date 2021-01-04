Attribute VB_Name = "Module1"
Sub stonks()
    'Create variables for worksheets
    Dim A As Worksheet
    Set A = Worksheets("A")
    Dim B As Worksheet
    Set B = Worksheets("B")
    Dim C As Worksheet
    Set C = Worksheets("C")
    Dim D As Worksheet
    Set D = Worksheets("D")
    Dim E As Worksheet
    Set E = Worksheets("E")
    Dim F As Worksheet
    Set F = Worksheets("F")
    Dim P As Worksheet
    Set P = Worksheets("P")
    
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
    
    'Define the last row of the sheet
    LastRow = A.Range("A" & Rows.Count).End(xlUp).Row
    
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        If OP = 0 Then
            
            OP = A.Cells(i, 3).Value
            
        End If
        
        If A.Cells(i, 1).Value <> A.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = A.Cells(i, 1).Value
            
            'Print the ticker symbol
            A.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = A.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            A.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            A.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + A.Cells(i, 7).Value
            
            A.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + A.Cells(i, 7).Value
                                    
        End If
        
        If A.Cells(j, 10).Value > 0 Then
        
            A.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf A.Cells(j, 10).Value < 0 Then
            
            A.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = B.Range("B" & Rows.Count).End(xlUp).Row
    
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        
        If OP = 0 Then
            
            OP = B.Cells(i, 3).Value
            
        End If
        
        If B.Cells(i, 1).Value <> B.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = B.Cells(i, 1).Value
            
            'Print the ticker symbol
            B.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = B.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            B.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            B.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + B.Cells(i, 7).Value
            
            B.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + B.Cells(i, 7).Value
                                    
        End If
        
        If B.Cells(j, 10).Value > 0 Then
        
            B.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf B.Cells(j, 10).Value < 0 Then
            
            B.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = C.Range("C" & Rows.Count).End(xlUp).Row
    
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        
        If OP = 0 Then
            
            OP = C.Cells(i, 3).Value
            
        End If
        
        If C.Cells(i, 1).Value <> C.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = C.Cells(i, 1).Value
            
            'Print the ticker symbol
            C.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = C.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            C.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            C.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + C.Cells(i, 7).Value
            
            C.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + C.Cells(i, 7).Value
                                    
        End If
        
        If C.Cells(j, 10).Value > 0 Then
        
            C.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf C.Cells(j, 10).Value < 0 Then
            
            C.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = D.Range("D" & Rows.Count).End(xlUp).Row
    
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        
        If OP = 0 Then
            
            OP = D.Cells(i, 3).Value
            
        End If
        
        If D.Cells(i, 1).Value <> D.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = D.Cells(i, 1).Value
            
            'Print the ticker symbol
            D.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = D.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            D.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            D.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + D.Cells(i, 7).Value
            
            D.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + D.Cells(i, 7).Value
                                    
        End If
        
        If D.Cells(j, 10).Value > 0 Then
        
            D.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf D.Cells(j, 10).Value < 0 Then
            
            D.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = E.Range("E" & Rows.Count).End(xlUp).Row
    
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        If OP = 0 Then
            
            OP = E.Cells(i, 3).Value
            
        End If
        
        If E.Cells(i, 1).Value <> E.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = E.Cells(i, 1).Value
            
            'Print the ticker symbol
            E.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = E.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            E.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            E.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + E.Cells(i, 7).Value
            
            E.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + E.Cells(i, 7).Value
                                    
        End If
        
        If E.Cells(j, 10).Value > 0 Then
        
            E.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf E.Cells(j, 10).Value < 0 Then
            
            E.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = F.Range("F" & Rows.Count).End(xlUp).Row
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        If OP = 0 Then
            
            OP = F.Cells(i, 3).Value
            
        End If
        
        If F.Cells(i, 1).Value <> F.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = F.Cells(i, 1).Value
            
            'Print the ticker symbol
            F.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = F.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            F.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            F.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + F.Cells(i, 7).Value
            
            F.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + F.Cells(i, 7).Value
                                    
        End If
        
        If F.Cells(j, 10).Value > 0 Then
        
            F.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf F.Cells(j, 10).Value < 0 Then
            
            F.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
    j = 2
    
    'Define the last row of the sheet
    LastRow = P.Range("P" & Rows.Count).End(xlUp).Row
    
    'Create a for loop from 1 to the last row
    For i = 2 To LastRow
        
        If OP = 0 Then
            
            OP = P.Cells(i, 3).Value
            
        End If
        
        If P.Cells(i, 1).Value <> P.Cells(i + 1, 1).Value Then
            
            'Set the ticker symbol
            ticker = P.Cells(i, 1).Value
            
            'Print the ticker symbol
            P.Cells(j, 9).Value = ticker
            
            'Assign closing price
            CP = P.Cells(i, 6)
            
            'Calculate the yearly change
            YearlyChange = OP - CP
            
            'Print the yearly change
            P.Cells(j, 10).Value = YearlyChange
            
            'Calculate the percent change
            PercentChange = YearlyChange / OP
            
            'Print the percent change
            P.Cells(j, 11).Value = FormatPercent(PercentChange, 2, , vbTrue)
            
            StockVolume = StockVolume + P.Cells(i, 7).Value
            
            P.Cells(j, 12).Value = StockVolume
                        
            j = j + 1
            
            OP = 0
            
            StockVolume = 0
        
        Else
            
            StockVolume = StockVolume + P.Cells(i, 7).Value
                                    
        End If
        
        If P.Cells(j, 10).Value > 0 Then
        
            P.Cells(j, 10).Interior.ColorIndex = 4
        
        ElseIf P.Cells(j, 10).Value < 0 Then
            
            P.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
End Sub
