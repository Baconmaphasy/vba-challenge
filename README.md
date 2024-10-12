# vba-challenge
Sub calculatePercentageChange()
    Dim out_row As Long
    Dim r As Long
    Dim lastRow As Long
    Dim previousValue As Double
    Dim currentValue As Double
    Dim percentageChange As Double
    
    out_row = 2
    lastRow = 99301
    
    ' Check if the cell is numeric before assigning
    If IsNumeric(Cells(2, 1).Value) Then
        previousValue = CDbl(Cells(2, 1).Value)
    Else
        MsgBox "The value in cell A2 is not numeric."
        Exit Sub
    End If

    For r = 2 To lastRow
        currentValue = Cells(r, 1).Value
        
        If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
            Cells(out_row, "I").Value = Cells(r, 1).Value
            out_row = out_row + 1
            
            If previousValue <> 0 Then
                percentageChange = (currentValue - previousValue) / previousValue * 100
            Else
                percentageChange = 0
            End If
            
            Cells(out_row, "K").Value = percentageChange
            out_row = out_row + 1
            previousValue = currentValue
        End If
    Next r
End Sub
