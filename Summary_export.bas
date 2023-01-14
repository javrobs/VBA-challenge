Attribute VB_Name = "Module1"
Sub summary():

For Each ws In Sheets
    
    ws.Activate
    
    '---------------------------- Table 1 ----------------------------'
        
    'Declare variables
    Dim pointer As Integer
    Dim opening As Double
    Dim closing As Double
    Dim volume As LongLong
    
    'Initialize variables
    pointer = 2
    'Initialize opening and volume to first numeric value. See readme.md for more information.
    opening = Range("C2").Value
    volume = Range("G2").Value
    
    'Write headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    'Populate table
    For i = 3 To Cells(Rows.Count, 1).End(xlUp).Row + 1
        
        If Cells(i, 1) = Cells(i - 1, 1) Then 'If current row is equal to last row
            
            volume = volume + Cells(i, 7).Value
                    
        Else 'If current row is not equal to last row
            
            'Obtain closing value from last row
            closing = Cells(i - 1, 6).Value
            'Write values in table
            Cells(pointer, 9).Value = Cells(i - 1, 1).Value
            Cells(pointer, 10).Value = closing - opening
            Cells(pointer, 11).Value = (closing - opening) / opening
            Cells(pointer, 12).Value = volume
            'Formatting
            Cells(pointer, 10).NumberFormat = "0.00"
            Cells(pointer, 11).NumberFormat = "0.00%"
            Cells(pointer, 12).NumberFormat = "0"
            'Coloring yearly change
            If Cells(pointer, 10).Value >= 0 Then
                Cells(pointer, 10).Interior.ColorIndex = 4
            Else
                Cells(pointer, 10).Interior.ColorIndex = 3
            End If
            'Overwrite values and increase pointer
            pointer = pointer + 1
            opening = Cells(i, 3).Value
            volume = Cells(i, 7).Value
                    
        End If
        
    Next i
        
    'Autofit table 1
    Range("I1:L" & Cells(Rows.Count, 9).End(xlUp).Row).Columns.AutoFit
        
    '---------------------------- Table 2 ----------------------------'
        
    'Declare variables
    Dim incTicker As String
    Dim incValue As Double
    Dim decTicker As String
    Dim decValue As Double
    Dim volTicker As String
    Dim volValue As Double
    
    'Initialize variables
    incValue = 0
    decValue = 0
    volValue = 0
    
    'Write headers
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    'Autofit headers
    Range("O2:O4").Columns.AutoFit
   
    
    'Obtain values
    For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row
        
        If Cells(i, 11).Value > incValue Then
            incValue = Cells(i, 11).Value
            incTicker = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value < decValue Then
            decValue = Cells(i, 11).Value
            decTicker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > volValue Then
            volValue = Cells(i, 12).Value
            volTicker = Cells(i, 9).Value
        End If

    Next i
    
    'Write values
    Range("P2").Value = incTicker
    Range("Q2").Value = incValue
    Range("P3").Value = decTicker
    Range("Q3").Value = decValue
    Range("P4").Value = volTicker
    Range("Q4").Value = volValue
    'Formatting
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").NumberFormat = "0.00E+0"
    
Next ws

End Sub
