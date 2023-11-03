Sub VBA()

'Loop through each sheet
For Each ws In Worksheets

'Declare the variables
Dim Ticker As String
Dim Row As Integer

Dim greatest_ticker As String
Dim greatest_increase As Double
Dim lowest_ticker As String
Dim greatest_decrease As Double
Dim total_ticker As String
Dim greatest_total As Double

'Input the Header Text for ws
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

Row = 2

'Testing a smaller sample
For i = 2 To 100000


open_stockprice = ws.Cells(i, 3).Value
close_stockprice = ws.Cells(i, 6).Value
volume = ws.Cells(i, 7).Value

'If the next row's ticker does not equal the current row, then input the previous ticker, add If
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

c = close_stockprice - open_stockprice
o = open_stockprice
a = volume

yearly_change = c
denom = denom + o
sum_volume = sum_volume + a

percent_change = (yearly_change / denom) * 100

ws.Cells(Row, 9).Value = Ticker
ws.Cells(Row, 10).Value = yearly_change
ws.Cells(Row, 11).Value = percent_change
ws.Cells(Row, 12).Value = sum_volume

Row = Row + 1

yearly_change = 0
denom = 0
percent_change = 0
sum_volume = 0

Else
Ticker = ws.Cells(i, 1).Value
c = close_stockprice - open_stockprice
yearly_change = yearly_change + c
a = volume
sum_volume = sum_volume + a

o = open_stockprice
denom = denom + o

End If
Next i

'Greatest Percent Increase Calculation
greatest_increase = 0
For e = 2 To 400
    If ws.Cells(e, 11).Value > greatest_increase Then
        greatest_ticker = ws.Cells(e, 9).Value
        greatest_increase = ws.Cells(e, 11).Value
    End If
Next e

ws.Cells(2, 15).Value = greatest_ticker
ws.Cells(2, 16).Value = greatest_increase

'Greatest Percent Decrease
greatest_decrease = 0
For f = 2 To 400
    If ws.Cells(f, 11).Value < greatest_decrease Then
        lowest_ticker = ws.Cells(f, 9).Value
        greatest_decrease = ws.Cells(f, 11).Value
    End If
Next f

ws.Cells(3, 15).Value = lowest_ticker
ws.Cells(3, 16).Value = greatest_decrease


'Greatest Total Volume
greatest_total = 0
For g = 2 To 402
    If ws.Cells(g, 12).Value > greatest_total Then
        total_ticker = ws.Cells(g, 9).Value
        greatest_total = ws.Cells(g, 12).Value
    End If
Next g

ws.Cells(4, 15).Value = total_ticker
ws.Cells(4, 16).Value = greatest_total


'Background Color Coding Yearly Change
For i = 2 To lastrow_summary_table
If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = vbRed
    Else
        ws.Cells(1, 10).Interior.ColorIndex = vbGreen
    End If
    Next i

Next ws

End Sub
