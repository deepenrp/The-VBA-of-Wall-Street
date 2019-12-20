Sub Stocks()

For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Volume"

    Dim Ticker As String
    Dim Volume As Double
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Volume = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Ticker = ws.Cells(i, 1).Value
    Volume = Volume + ws.Cells(i, 7).Value

ws.Range("I" & Summary_Table_Row).Value = Ticker
ws.Range("J" & Summary_Table_Row).Value = Volume

Volume = 0

Summary_Table_Row = Summary_Table_Row + 1


Else: Volume = Volume + ws.Cells(i, 7).Value


End If

Next i

Next ws


End Sub
