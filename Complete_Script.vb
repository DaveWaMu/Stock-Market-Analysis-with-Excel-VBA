Sub VBA_homework():

Dim lastrow As Long
Dim next_row As Double
Dim ticker_value As String
Dim ticker_column As Double
Dim open_price As Double
Dim close_price As Double
Dim price_change_value As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_total_volume As Double
Dim greatest_increase_ticker As String
Dim greatest_decrease_ticker As String
Dim greatest_total_volume_ticker As String

Set ws = Sheets(1)

For Each ws In Worksheets

ws.Activate

ticker_column = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
next_row = 2
total_stock_volume = 0
    
'Table set-up
Cells(1, "I").Value = "Ticker"
Cells(1, "P").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"
Cells(1, "Q").Value = "Value"
Cells(2, "O").Value = "Greatest % Increase"
Cells(3, "O").Value = "Greatest % Decrease"
Cells(4, "O").Value = "Greatest Total Volume"

'Loop through ticker info
For I = 2 To lastrow
    If Cells(I + 1, "A").Value = Cells(I, "A").Value Then
        ticker_column = ticker_column + Cells(I, "A").Count
        total_stock_volume = total_stock_volume + Cells(I, 7).Value
    ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        ticker_column = ticker_column + 1
        ticker_value = Cells(I, 1).Value
        Range("I" & next_row).Value = ticker_value
        open_price = Cells(I - ticker_column + 1, 3).Value
        close_price = Cells(I, 6).Value
        price_change_value = close_price - open_price
            If open_price = 0 Then
            percent_change = 0
            Else: percent_change = price_change_value / open_price
            End If
        Range("K" & next_row).Value = percent_change
        Range("J" & next_row).Value = price_change_value
        total_stock_volume = total_stock_volume + Cells(I, 7).Value
        Range("L" & next_row).Value = total_stock_volume
        next_row = next_row + 1
        total_stock_volume = 0
        ticker_column = 0
    End If
Next I

'Color Column
lastrow = Cells(Rows.Count, "K").End(xlUp).Row

For I = 2 To lastrow
    'Color Index
    If Cells(I, "J").Value < 0 Then
    Cells(I, "J").Interior.ColorIndex = 3
    Else: Cells(I, "J").Interior.ColorIndex = 4
    End If
Next I

'Bonus
'Greatest percent increase
greatest_increase = WorksheetFunction.Max(Range("K2:K" & lastrow))
Cells(2, "Q") = greatest_increase
greatest_increase_ticker = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
Cells(2, "P") = Cells(greatest_increase_ticker + 1, "I")

'Greastest percent decrease
greatest_decrease = WorksheetFunction.Min(Range("K2:K" & lastrow))
Cells(3, "Q") = greatest_decrease
greatest_decrease_ticker = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
Cells(3, "P") = Cells(greatest_decrease_ticker + 1, "I")

'Greatest total volume
greatest_total_volume = WorksheetFunction.Max(Range("L2:L" & lastrow))
Cells(4, "Q") = greatest_total_volume
greatest_total_volume_ticker = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
Cells(4, "P") = Cells(greatest_total_volume_ticker + 1, "I")

'Clean up
Range("K:K").NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"
Range("I291:L" & lastrow).Delete
Range("I291:L" & lastrow).Interior.ColorIndex = 0
ws.Cells.EntireColumn.AutoFit

Next ws

End Sub




