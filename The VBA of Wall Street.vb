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


For r = 2 To lastrow

    If ws.Range("J" & r).Value > 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & r).Value < 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 3
        
    End If

    Next r
    

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

For a = 2 To lastrow


    If ws.Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(a, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(a, 9).Value
    End If

    Next a

For b = 2 To lastrow
    
    If ws.Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(b, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(b, 9).Value
    End If
    
   Next b

For c = 2 To lastrow
    
    If ws.Cells(c, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(c, 12).Value
        ws.Range("Q4").Value = GreatestVolume
        ws.Range("P4").Value = ws.Cells(c, 9).Value
    End If
  
    Next c
 
ws.Columns("A:Q").AutoFit
    
Next ws


End Sub




