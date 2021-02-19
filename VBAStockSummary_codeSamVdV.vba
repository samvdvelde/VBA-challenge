Sub Stocksummary_for_each_year()

For Each ws In Worksheets



'Set row labels'

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"


Dim Tick As Long
Tick = 2
Dim i As Long
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row





'Format percent range as percent'

ws.Range("K:K").NumberFormat = "0.00%"



'Insert tickers in summary table'


For i = 1 To LastRow

    'Fill ticker column'

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        ws.Cells(Tick, 9).Value = ws.Cells(i + 1, 1).Value
        
          Tick = Tick + 1
          
        End If
        
        
Next i

'Calculate year changes and add to summary table'

Dim SumRow As Long
SumRow = 2
Dim TradeDay As Integer
TradeDay = 0
Dim YearChange As Double
YearChange = 0
Dim PercentChange As Double
PercentChange = 0


For i = 2 To LastRow

    

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value and ws.Cells(i + 1, 3).Value <> 0 Then

        
    'Define year change'
        
        YearChange = ws.Cells(i, 6).Value - ws.Cells((i - TradeDay), 3).Value
        
   'Define percent change'
   
        PercentChange = YearChange / (ws.Cells((i - TradeDay), 3).Value)

        
    'Define stock volume'
    
        Volume = Volume + ws.Cells(i, 7).Value
        
        
        
        
    'Print year change to summary table'
    
        ws.Cells(SumRow, 10).Value = YearChange
        
    'Print percent change to summary table'
    
        ws.Cells(SumRow, 11).Value = PercentChange
        
    'Print percent volume to summary table'
    
        ws.Cells(SumRow, 12).Value = Volume
        
        
        
    'Reset variables'
    
        TradeDay = 0
        YearChange = 0
        PercentChange = 0
        Volume = 0
        
        
    'Skip to next row in the summary table'
        
        SumRow = SumRow + 1
        
        
   'If ticker value does not change, do this'
        
      Else
      
      TradeDay = TradeDay + 1
      
      Volume = Volume + ws.Cells(i, 7).Value
      
      
        
        End If
        
 Next i



'Format yearly change cells'

Dim LastRowChange As Long
LastRowChange = ws.Cells(Rows.Count, 10).End(xlUp).Row


For i = 2 To LastRowChange

    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 2
        
    End If
    
Next i




 'set column/row labels of MinMax table'

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"


'Format max and min pct cells as pct'

ws.Range("Q2:Q3").NumberFormat = "0.00%"


'Find values in summary table'

Dim pctrng As Range
Dim volrng As Range
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
Set pctrng = ws.Range("K:K")
Set volrng = ws.Range("L:L")
Dim MinPct As Double
Dim MaxPct As Double
Dim MaxVol As Double


MaxPct = Application.WorksheetFunction.Max(pctrng)
MinPct = Application.WorksheetFunction.Min(pctrng)
MaxVol = Application.WorksheetFunction.Max(volrng)


'Add max pct, min pct and max vol to MinMax table'

ws.Cells(2, 17).Value = MaxPct
ws.Cells(3, 17).Value = MinPct
ws.Cells(4, 17).Value = MaxVol


'Add tickers to MinMax table'

For i = 2 To LastRow

    If ws.Cells(i, 11).Value = MaxPct Then
    ws.Cells(2, 16).Value = ws.Cells(i, 11).Offset(0, -2).Value
    
    ElseIf ws.Cells(i, 11).Value = MinPct Then
    ws.Cells(3, 16).Value = ws.Cells(i, 11).Offset(0, -2).Value
    
    ElseIf ws.Cells(i, 12).Value = MaxVol Then
    ws.Cells(4, 16).Value = ws.Cells(i, 12).Offset(0, -3).Value
    
    End If
    
Next i


Next ws




End Sub