Sub Moderate()
'For this I modified the "Easy" routine to added the extra report

'Write headers for report data
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'-----------------------
'Define variables needed
'-----------------------

'RowHeader is the number of rows on the sheet above the actual data. Common practice.
RowHeader = 1
'variable TickerPlace keeps each output result on a different line. Starts with 2
TickerPlace = 1 + RowHeader
'VolumeYear is the variable that holds the volume totals for each stock. Long number.
Dim VolumeYear As Double
VolumeYear = 0
'Extract the maximum row so I don't have to know the total number of rows in the sheet
'"- RowHeader" at the end to remove the empty row at the bottom
Dim MaxRow As Long
MaxRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row - RowHeader
'StockOpen is the variable that holds the opening price for the year.
Dim StockOpen As Double
'StockClose is the variable that holds the closing price for the year.
Dim StockClose As Double


'The initial value is set for the sheet, then changed inside the loop through ticker change
StockOpen = Cells(RowHeader + 1, 3).Value


'----------------------------------------------------------------------------------------------
'Loop through the rows, picking up the different stock tickers and gathering the report numbers
'----------------------------------------------------------------------------------------------

For i = 1 To MaxRow

'Compare tickers from current row and next, to detect the last row for current ticker
'This test will only be true for the last row of every ticker
If Cells(i + RowHeader, 1).Value <> Cells(i + RowHeader + 1, 1).Value Then
    'Add the last volume to the total so far
    VolumeYear = VolumeYear + Cells(i + RowHeader, 7)
    'Copy the current ticker and total volume to the report
    Cells(TickerPlace, 9) = Cells(i + RowHeader, 1)
    Cells(TickerPlace, 12).Value = VolumeYear
    'Capture the closing price
    StockClose = Cells(i + RowHeader, 6).Value
        'Calculate the open and close difference and percent change
        '----------------------------------------------------------
        Cells(TickerPlace, 10).Value = StockClose - StockOpen
        'Color flag positve and negative price changes
        If Cells(TickerPlace, 10).Value < 0 Then
            Cells(TickerPlace, 10).Interior.ColorIndex = 3
            ElseIf Cells(TickerPlace, 10).Value > 0 Then
            Cells(TickerPlace, 10).Interior.ColorIndex = 4
        End If
        'Calculate the percent change and post to report
        'Using If statement to prevent division by zero.
        If StockOpen = 0 Then
            Cells(TickerPlace, 11).Value = 0
            Cells(TickerPlace, 11).Interior.ColorIndex = 3
            Cells(TickerPlace, 13).Value = "Open =" & StockOpen & ", Close =" & StockClose
            Cells(TickerPlace, 13).Interior.ColorIndex = 3
        Else
            Cells(TickerPlace, 11).Value = ((StockClose - StockOpen) / StockOpen)
            Cells(TickerPlace, 11).NumberFormat = "0.00%"
        End If
    'Move one row down on the report to write the next ticker results
    TickerPlace = TickerPlace + 1
    'Reset the volume so the next ticker volume can be calculated
    VolumeYear = 0
    'Reset the StockOpen variable to capture the initial price of the next row (new ticker)
    StockOpen = Cells(i + RowHeader + 1, 3).Value
    
'All but the last row of each ticker will go through this Else and add the volumes in them
Else
    VolumeYear = VolumeYear + Cells(i + RowHeader, 7)
End If

Next i

'Adjusting the report columns to fit the data width
Columns("I:M").Select
Selection.EntireColumn.AutoFit

End Sub