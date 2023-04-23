Sub StockSummary()
 
 For Each ws In Worksheets
 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim Total As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    

    yearlyChange = 0
    percentChange = 0
    Total = 0
    openPrice = 0
    closePrice = 0
    Dim StockSummary_Row As Integer
    StockSummary_Row = 2
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim YearRow As Long
    YearRow = 2
 

    For i = 2 To lastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(YearRow, 3).Value
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            percentChange = (yearlyChange / openingPrice)
            Total = Total + ws.Cells(i, 7).Value

            ws.Range("I" & StockSummary_Row).Value = ticker
            ws.Range("J" & StockSummary_Row).Value = yearlyChange
            ws.Range("K" & StockSummary_Row).Value = percentChange
            ws.Range("K" & StockSummary_Row).NumberFormat = "0.00%"
            ws.Range("L" & StockSummary_Row).Value = Total

            StockSummary_Row = StockSummary_Row + 1
            YearRow = i + 1
            
            yearlyChange = 0
            percentChange = 0
            Total = 0
        Else
            
            Total = Total + ws.Cells(i, 7).Value
            
        End If

    Next i
    
        Dim SummaryLastRow As Long
        SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To SummaryLastRow
        If ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 0
        End If
    Next i
    
    For i = 2 To SummaryLastRow
        If ws.Cells(i, 11).Value < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 11).Value > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 11).Interior.ColorIndex = 0
        End If
    Next i
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    Dim GreatestInc As Double
    Dim GreatestDec As Double
    Dim GreastestVol As Double
    Dim TickerInc As String
    Dim TickerDec As String
    Dim TickerVol As String
    
    GreatestInc = 0
    GreatestDec = 0
    GreatestVol = 0
    
    For i = 2 To SummaryLastRow
        If ws.Cells(i, 11).Value > GreatestInc Then
        GreatestInc = ws.Cells(i, 11).Value
        TickerInc = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 11).Value < GreatestDec Then
        GreatestDec = ws.Cells(i, 11).Value
        TickerDec = ws.Cells(i, 9).Value
        
        ElseIf ws.Cells(i, 12).Value > GreatestVol Then
        GreatestVol = ws.Cells(i, 12).Value
        TickerVol = ws.Cells(i, 9).Value
        
        End If
        
    Next i
    
    ws.Cells(2, 16).Value = TickerInc
    ws.Cells(3, 16).Value = TickerDec
    ws.Cells(4, 16).Value = TickerVol
    ws.Cells(2, 17).Value = GreatestInc
    ws.Cells(3, 17).Value = GreatestDec
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = GreatestVol
  
  Next ws
        
End Sub

