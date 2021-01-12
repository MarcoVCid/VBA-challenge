Attribute VB_Name = "Module1"
Sub hw()
    
    Dim t As Double
    Dim tsv As Double
    Dim op As Double
    Dim cp As Double
    
    Application.ScreenUpdating = False
    
    For Each ws In Worksheets
    
        t = 2
        tsv = 0
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
    
        For i = 2 To ws.Range("A1", ws.Range("A1").End(xlDown)).Count
            If ws.Cells(i, 2).Value = ws.Range("B2").Value Then
                op = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 2).Value = ws.Range("B2").End(xlDown).Value Then
                cp = ws.Cells(i, 6).Value
            End If
            
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                tsv = tsv + ws.Cells(i, 7).Value
            Else
                tsv = tsv + ws.Cells(i, 7).Value
                ws.Cells(t, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(t, 10).Value = cp - op
                If ws.Cells(t, 10).Value < 0 Then
                    ws.Cells(t, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(t, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(t, 11).Value = cp / op - 1
                ws.Columns("K").NumberFormat = "0.0000%"
                ws.Cells(t, 12).Value = tsv
                tsv = 0
                t = t + 1
            End If
        Next i
    Next
    
    Call bonus
    
    Application.ScreenUpdating = True
    
End Sub
Sub bonus()
    
    For Each ws In Worksheets
    
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest total volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2", ws.Range("K2").End(xlDown)))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2", ws.Range("K2").End(xlDown)))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2", ws.Range("L2").End(xlDown)))
        ws.Range("Q2:Q3").NumberFormat = "0.0000%"
    
        ws.Range("P2").Value = WorksheetFunction.XLookup(ws.Range("Q2").Value, ws.Range("K2", ws.Range("K2").End(xlDown)), ws.Range("I2", ws.Range("I2").End(xlDown)))
        ws.Range("P3").Value = WorksheetFunction.XLookup(ws.Range("Q3").Value, ws.Range("K2", ws.Range("K2").End(xlDown)), ws.Range("I2", ws.Range("I2").End(xlDown)))
        ws.Range("P4").Value = WorksheetFunction.XLookup(ws.Range("Q4").Value, ws.Range("L2", ws.Range("L2").End(xlDown)), ws.Range("I2", ws.Range("I2").End(xlDown)))
    
        ws.Columns("O:Q").AutoFit
    Next
    
End Sub
Sub reset()
    
    For Each ws In Worksheets
        ws.Columns("I:Q").Delete
    Next
    
End Sub
