Attribute VB_Name = "Module1"
Sub Stockloop():

'Function to allow loop to run through eash ws
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

'Equation to find LastRow
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Define all variables before the starting the loop
Dim Opening_Balance As Double
Opening_Balance = Cells(2, 3).Value
Dim Closing_Balance As Double
Dim Yearly_Change As Double
Dim Ticker_Symbol As String
Dim Percent As Double
Dim StockV As Double
StockV = 0
Dim StartRow As Double
StartRow = 2


'Create summary table to store data
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Start loop
Dim i As Long
For i = 2 To LastRow

'Equation to check ticker symbols and pull differing ones
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_Symbol = Cells(i, 1).Value
Cells(StartRow, 9).Value = Ticker_Symbol

'Adding conditions and Initializing the Close Price
Closing_Balance = Cells(i, 6).Value

'Calculation for Year change
Yearly_Change = Closing_Balance - Opening_Balance
Cells(StartRow, 10).Value = Yearly_Change

'Percent Change
If (Opening_Balance = 0 And Closing_Balance = 0) Then
Percent = 0
ElseIf (Opening_Balance = 0 And Closing_Balance <> 0) Then
Percent = 1
'Storing percent in the summary table
Else
Percent = Yearly_Change / Opening_Balance
Cells(StartRow, 11).Value = Percent
Cells(StartRow, 11).NumberFormat = "0.00%"
End If

'Calculating for Total Column
StockV = StockV + Cells(i, 7).Value
Cells(StartRow, 12).Value = StockV
StartRow = StartRow + 1

'Reseting Variables to calculate if the symbols are the same
Opening_Balance = Cells(i + 1, 3)
StockV = 0

'If they are NOT differing symbols
Else
StockV = StockV + Cells(i, 7).Value
End If
Next i

'Finding Yearly Change last row in order to assign colors
LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
'Second loop
Dim j As Integer
For j = 2 To LastRow2
'Color index
 If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
'Positive change
Cells(j, 10).Interior.ColorIndex = 4
'Negative Change
ElseIf Cells(j, 10).Value < 0 Then
Cells(j, 10).Interior.ColorIndex = 3

End If
Next j
Next ws
End Sub

