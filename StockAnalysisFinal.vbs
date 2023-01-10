Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Storing paramenters as variables
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double

Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim lastrow As Long
Dim Sheet As Worksheet

' Looping through all sheets
For Each Sheet In Worksheets

On Error Resume Next
' Determining the Last Row
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

If Err.Number > 0 Then
  MsgBox (Cells(Rows.Count, "A").Row)
End If

'  Making Header Tabs
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

' Tracking row for the ticker summary
Dim TickerCounterSummary As Double
TickerCounterSummary = 2

' Saving open and close ticker values
OpenCloseValues = 2
TotalStockVolume = 0


For s = 2 To lastrow
    If Sheet.Cells(s + 1, 1).Value <> Sheet.Cells(s, 1) Then
        TotalStockVolume = TotalStockVolume + Cells(s, 7).Value
        Ticker = Sheet.Cells(s, 1).Value
    
    Sheet.Range("I" & TickerCounterSummary).Value = Ticker
    Sheet.Range("L" & TickerCounterSummary).Value = TotalStockVolume
    
    TotalStockVolume = 0
    ClosingPrice = Sheet.Cells(s, 6)
    
    ' setting the parameters for calculting the percent change
        If OpeningPrice = 0 Then
            YearlyChange = 0
            PercentChange = 0
        Else:
            YearlyChange = ClosingPrice - OpeningPrice
            PercentChange = ((ClosingPrice - OpeningPrice) / Opening) * 100
        End If
    
            Sheet.Range("J" & TickerCounterSummary).Value = YearlyChange
            Sheet.Range("K" & TickerCounterSummary).Value = PercentChange
            Sheet.Range("K" & TickerCounterSummary).Style = "Percent"
            Sheet.Range("J" & TickerCounterSummary).NumberFormat = "0.00%"
    
            TickerCounterSummary = TickerCounterSummary + 1
    
    ElseIf Sheet.Cells(s - 1, 1).Value <> Sheet.Cells(s, 1) Then
        OpeningPrice = Sheet.Cells(s, 3)
        
    Else: TotalStockVolume = TotalStockVolume + Sheet.Cells(s, 7).Value

    End If

    Next s
    
For c = 2 To lastrow
    If Sheet.Range("J" & c).Value > 0 Then
        Sheet.Range("J" & c).Interior.ColorIndex = 4
        
    ElseIf Sheet.Range("J" & c).Value < 0 Then
        Sheet.Range("J" & c).Interior.ColorIndex = 3
        
    End If
    
    Next c
    
' Setting Hard Challenge Label Tabs

Sheet.Range("P1").Value = "Ticker"
Sheet.Range("Q1").Value = "Value"

Sheet.Range("O2").Value = "Greatest % Increase"
Sheet.Range("O3").Value = "Greatest % Decrease"
Sheet.Range("O4").Value = "Greatest Total Volume"

' Storing Hard Challege parameters as variables
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Double

' Setting Parameters to zero
GreatestIncrease = 0
GreatestDecrease = 0
GreatestTotalVolume = 0
    
' Conditionals

For x = 2 To lastrow
    If Sheet.Cells(x, 11).Value > GreatestIncrease Then
        GreatestIncrease = Sheet.Cells(x, 11).Value
        Sheet.Range("Q2").Value = GreatestIncrease
        Sheet.Range("Q2").Value = "Percent"
        Sheet.Range("Q2").NumberFormat = "0.00%"
        Sheet.Range("P2").Value = Sheet.Cells(x, 9).Value
        
    End If
    
    Next x
    
    
For y = 2 To lastrow
    If Sheet.Cells(y, 11).Value < GreatestDecrease Then
        GreatestDecrease = Sheet.Cells(y, 11).Value
        Sheet.Range("Q3").Value = GreatestDecrease
        Sheet.Range("Q3").Style = "Percent"
        Sheet.Range("Q3").NumberFormat = "0.00%"
        Sheet.Range("P3").Value = Sheet.Cells(y, 9).Value
    
    End If
    
    Next y


For v = 2 To lastrow
    If Sheet.Cells(v, 12).Value > GreatestTotalVolume Then
        GreatestTotalVolume = Sheet.Cells(v, 12).Value
        Sheet.Range("Q4").Value = GreatestTotalVolume
        Sheet.Range("P4").Value = Sheet.Cells(v, 9).Value
    
    End If
    
    Next v
        
        
Next Sheet
    
    
End Sub
