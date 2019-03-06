Sub Easy()
'Write headers for report data
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Stock Volume"

'Define variables needed
'-----------------------

'variable TickerPlace keeps each output result on a different line. Starts with 2
TickerPlace = 2
'RowHeader is the number of rows on the sheet above the actual data. Common practice.
RowHeader = 1
'VolumeYear is the variable that holds the volume totals for each stock. Long number.
Dim VolumeYear As Double
VolumeYear = 0
'Extract the maximum row so I don't have to know the total number of rows in the sheet
'"- 1" at the end to remove the header row and use the loop from 1
Dim MaxRow As Long
MaxRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row - 1

'Loop through the rows, picking up the different stock tickers and adding up volumes
'-----------------------------------------------------------------------------------

For I = 1 To MaxRow
'Test if the current row is the last for a particular ticker by comparing to the ticker on the next row
'This test will only be true for the last row of every ticker
If Cells(I + RowHeader, 1).Value <> Cells(I + RowHeader + 1, 1).Value Then
    'Add the last volume to the total so far
    VolumeYear = VolumeYear + Cells(I + RowHeader, 7)
    'Copy the current ticker and total volume to the report
    Cells(TickerPlace, 9) = Cells(I + RowHeader, 1)
    Cells(TickerPlace, 10).Value = VolumeYear
    'Move one row down on the report to write the next ticker results
    TickerPlace = TickerPlace + 1
    'Reset the volume so the next ticker volume can be calculated
    VolumeYear = 0
'All but the last row of each ticker will go through this Else and add the volumes in them
Else
    VolumeYear = VolumeYear + Cells(I + RowHeader, 7)
End If
Next I
End Sub