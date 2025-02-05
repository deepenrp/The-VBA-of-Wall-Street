Sub Stocks()

For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    
    Dim StockOpen As Double
    Dim StockClose As Double
    
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Volume = 0

Dim Summary_Table_Row As Double
Summary_Table_Row = 2

For i = 2 To lastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value

ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Range("L" & Summary_Table_Row).Value = Volume

Volume = 0

StockClose = ws.Cells(i, 6)

    If StockOpen = 0 Then
    YearlyChange = 0
    PercentChange = 0
    Else:
    YearlyChange = StockClose - StockOpen
    PercentChange = (StockClose - StockOpen) / StockOpen
    End If
    
ws.Range("J" & Summary_Table_Row).Value = YearlyChange
ws.Range("K" & Summary_Table_Row).Value = PercentChange
ws.Range("K" & Summary_Table_Row).Style = "Percent"
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

Summary_Table_Row = Summary_Table_Row + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    StockOpen = ws.Cells(i, 3)
    

Else: Volume = Volume + ws.Cells(i, 7).Value

End If


    Next i


For i = 2 To lastrow

If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4

ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
End If

    Next i
    
Next ws


End Sub


