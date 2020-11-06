Sub StockCount()

'variables'
Dim Ticker As String
Dim YearlyChange As Double
YearlyChange = 0
Dim PercentChange As Double
PercentChange = 0
Dim TotalStock As Double
Dim ClosePrice As Double
ClosePrice = 0
Dim OpenPrice As Double
OpenPrice = 0
Dim Greatest As Double
Dim GreatestTicker As String
Dim GreatestTotal As Double
Dim TotalTicker As String
Dim GreatestDecrease As Double
Dim DecreaseTicker As String
Dim LastRow As Double

'loop through all worksheets'
For Each ws In Worksheets

'variable to hold total per ticker name'
TotalStock = 0

'keep track of ticker name location'
Dim j As Integer
j = 2

'find last row of each worksheet'
LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'initialize open price'
OpenPrice = ws.Cells(2, 3).Value

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TotalStock = TotalStock + ws.Cells(i, 7).Value
        
        Ticker = ws.Cells(i, 1).Value
        ClosePrice = ws.Cells(i, 6).Value
        YearlyChange = ClosePrice - OpenPrice
        
            PercentChange = YearlyChange / OpenPrice
               
        'color formatting, red for negative and green for positive'
        If YearlyChange > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
           
        'initialize variables'
        Greatest = 0
        GreatestTotal = 0
        GreatestDecrease = 0
        
        'find greatest increase, decrease and total volume values'
        If PercentChange > Greatest Then
            Greatest = PercentChange
            GreatestTicker = Ticker
        End If
           
        If PercentChange < GreatestDecrease Then
            GreatestDecrease = PercentChange
            DecreaseTicker = Ticker
        End If
           
        If GreatestTotal < TotalStock Then
            GreatestTotal = TotalStock
            TotalTicker = Ticker
        End If
            

            'variables for ticker, yearlychange, percentchange, totalstock'

            ws.Cells(j, 9).Value = Ticker
            ws.Cells(j, 10).Value = YearlyChange
            ws.Cells(j, 11).Value = PercentChange
            ws.Cells(j, 12).Value = TotalStock
        
            
            'reset to zero'
            TotalStock = 0
            
            'add 1 to row count'
            j = j + 1
    Else
        'add to ticker name total stock'
        TotalStock = TotalStock + ws.Cells(i, 7).Value

    End If
    

Next i

'greatest increase, decrease and total volume'
ws.Range("P2").Value = Greatest
ws.Range("Q2").Value = GreatestTicker
ws.Range("P3").Value = GreatestDecrease
ws.Range("Q3").Value = DecreaseTicker
ws.Range("P4").Value = GreatestTotal
ws.Range("Q4").Value = TotalTicker

'set column and row headers'
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'format all sheets'
    With ws
    .Columns("K:K").NumberFormat = "0.00%"
    .Range("P2").NumberFormat = "0.00%"
    .Range("P3").NumberFormat = "0.00%"
    
    End With

Next ws

End Sub