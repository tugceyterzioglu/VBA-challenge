Sub Stock_Loop()

Dim ws As Worksheet
Dim ticker_label As String
Dim ticker_volume As Double
ticker_volume = 0

Dim last_row As Long
Dim summary_row As Integer
summary_row = 2
Dim lastrow_summary_table As Long

Dim open_price As Double
Dim close_price As Double
Dim quarterly_change As Double
Dim percent_change As Double

Dim ticker_greatest_increase As String
Dim ticker_greatest_decrease As String
Dim ticker_greatest_volume As String

Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_total_volume As Double

greatest_percent_increase = 0
greatest_percent_decrease = 0
greatest_total_volume = 0

For Each ws In ThisWorkbook.Worksheets
If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then

summary_row = 2
ticker_volume = 0
greatest_percent_increase = 0
greatest_percent_decrease = 0
greatest_total_volume = 0

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

open_price = ws.Cells(2, 3).Value

For I = 2 To last_row
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
ticker_label = ws.Cells(I, 1).Value
ticker_volume = ticker_volume + ws.Cells(I, 7).Value

ws.Range("I" & summary_row).Value = ticker_label
ws.Range("L" & summary_row).Value = ticker_volume

close_price = ws.Cells(I, 6).Value
quarterly_change = (close_price - open_price)
ws.Range("J" & summary_row).Value = quarterly_change

If (open_price) = 0 Then
percent_change = 0

Else
percent_change = quarterly_change / open_price
End If

ws.Range("K" & summary_row).Value = percent_change
ws.Range("K" & summary_row).NumberFormat = "0.00%"

If percent_change > greatest_percent_increase Then
greatest_percent_increase = percent_change
ticker_greatest_increase = ticker_label
End If

If percent_change < greatest_percent_decrease Then
greatest_percent_decrease = percent_change
ticker_greatest_decrease = ticker_label
End If

If ticker_volume > greatest_total_volume Then
greatest_total_volume = ticker_volume
ticker_greatest_volume = ticker_label
End If


summary_row = summary_row + 1
ticker_volume = 0
open_price = ws.Cells(I + 1, 3).Value

Else
ticker_volume = ticker_volume + ws.Cells(I, 7).Value

End If

Next I

lastrow_summary_table = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row

For I = 2 To lastrow_summary_table
If ws.Cells(I, 10).Value > 0 Then
ws.Cells(I, 10).Interior.ColorIndex = 10
Else
ws.Cells(I, 10).Interior.ColorIndex = 3
End If
Next I

ws.Range("P2").Value = ticker_greatest_increase
ws.Range("Q2").Value = Format(greatest_percent_increase, "0.00%")
ws.Range("P3").Value = ticker_greatest_decrease
ws.Range("Q3").Value = Format(greatest_percent_decrease, "0.00%")
ws.Range("P4").Value = ticker_greatest_volume
ws.Range("Q4").Value = greatest_total_volume

End If

Next ws

End Sub