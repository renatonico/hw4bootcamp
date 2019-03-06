Sub challenge()

'==========================================================================
'= This macro is run on the whole workbook (Multiple_year_stock_data.xlsx)=
'= There are 2 parts (from the "hard" work):                              =
'= Part 1 runs through the original data and creates a report with a row  =
'=        for each ticker with the stock's total volume and yearly change =
'= Part 2 runs through the report from part 1 and reports the stocks with =
'=        the greatest variations in price and the greatest volume traded =
'==========================================================================

'--------------------------------------------------------------------------
'This For Each loop takes each spreadsheet, starting with 2016 as it is the
'first one, so the Cells and Range references in the versions of the code
'where the user is in the active sheet, must be converted to ws.Cells.
'All variables used to count remained the same, as they are reset at the
'beginning of the analysis loops. The autofit functions had to be adapted too.
'--------------------------------------------------------------------------

For Each ws In Worksheets

'       ===========================================
'       = PART 1 (adapted from easy and moderate) =
'       ===========================================

    '------------------------------
    'Write headers for report data
    '------------------------------

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

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
    MaxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row - RowHeader
    'StockOpen is the variable that holds the opening price for the year.
    Dim StockOpen As Double
    'The initial value is set for the sheet, then changed inside the loop through ticker change
    StockOpen = ws.Cells(RowHeader + 1, 3).Value
    'StockClose is the variable that holds the closing price for the year.
    Dim StockClose As Double

'----------------------------------------------------------------------------------------------
'Loop through the rows, picking up the different stock tickers and gathering the report numbers
'----------------------------------------------------------------------------------------------

    For i = 1 To MaxRow

    'Compare tickers from current row and next, to detect the last row for current ticker
    'This test will only be true for the last row of every ticker
    If ws.Cells(i + RowHeader, 1).Value <> ws.Cells(i + RowHeader + 1, 1).Value Then
        'Add the last volume to the total so far
        VolumeYear = VolumeYear + ws.Cells(i + RowHeader, 7)
        'Copy the current ticker and total volume to the report
        ws.Cells(TickerPlace, 9) = ws.Cells(i + RowHeader, 1)
        ws.Cells(TickerPlace, 12).Value = VolumeYear
        'Capture the closing price
        StockClose = ws.Cells(i + RowHeader, 6).Value
        'Calculate the open and close difference and percent change
        '----------------------------------------------------------
        ws.Cells(TickerPlace, 10).Value = StockClose - StockOpen
        'Color flag positve and negative price changes
        If ws.Cells(TickerPlace, 10).Value < 0 Then
            ws.Cells(TickerPlace, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(TickerPlace, 10).Value > 0 Then
            ws.Cells(TickerPlace, 10).Interior.ColorIndex = 4
        End If
        'Calculate the percent change and post to report
        'Using If statement to prevent division by zero
        If StockOpen = 0 Then
            ws.Cells(TickerPlace, 11).Value = 0
            ws.Cells(TickerPlace, 11).Interior.ColorIndex = 3
            ws.Cells(TickerPlace, 13).Value = "Open =" & StockOpen & ", Close =" & StockClose
            ws.Cells(TickerPlace, 13).Interior.ColorIndex = 3
        Else
            ws.Cells(TickerPlace, 11).Value = ((StockClose - StockOpen) / StockOpen)
            ws.Cells(TickerPlace, 11).NumberFormat = "0.00%"
        End If
    'Move one row down on the report to write the next ticker results
    TickerPlace = TickerPlace + 1
    'Reset the volume so the next ticker volume can be calculated
    VolumeYear = 0
    'Reset the StockOpen variable to capture the initial price of the next row (new ticker)
    StockOpen = ws.Cells(i + RowHeader + 1, 3).Value
    
    'All but the last row of each ticker will go through this Else and add the volumes in them
    Else
        VolumeYear = VolumeYear + ws.Cells(i + RowHeader, 7)
    End If

    Next i


    '       ============================================
    '       =       PART 2 ( adapted from hard)        =
    '       ============================================

    '-------------------------------------------------
    'Find and report the greatest % changes and volume
    '-------------------------------------------------

    'Write headers for report data
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    'Set variables
    '-------------

    'Variables that hold the values needed to report.
    Dim GreatIncTicker As String
    Dim GreatDecTicker As String
    Dim GreatVolTicker As String

    'Initial values set to first row of data
    Dim GreatInc As Double
    GreatInc = ws.Cells(2, 11).Value
    Dim GreatDec As Double
    GreatDec = ws.Cells(2, 11).Value
    Dim GreatVol As Double
    GreatVol = ws.Cells(2, 12).Value

    'Set the maximum row to search for data
    'Note that the variable TickerPlace was incremented by 1 after reaching the last row, so I need to subtract it
    RepMaxRow = TickerPlace - 1

    'Loop through the source report to search for the data
    '-----------------------------------------------------

    For i = 3 To RepMaxRow

    If ws.Cells(i, 11).Value > GreatInc Then
        GreatInc = ws.Cells(i, 11).Value
        GreatIncTicker = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 11).Value < GreatDec Then
        GreatDec = ws.Cells(i, 11).Value
        GreatDecTicker = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 12).Value > GreatVol Then
        GreatVol = ws.Cells(i, 12).Value
        GreatVolTicker = ws.Cells(i, 9).Value
    End If

    Next i

    'Write the results to the table
    ws.Cells(2, 16).Value = GreatIncTicker
    ws.Cells(3, 16).Value = GreatDecTicker
    ws.Cells(4, 16).Value = GreatVolTicker
    ws.Cells(2, 17).Value = GreatInc
    ws.Cells(3, 17).Value = GreatDec
    ws.Cells(4, 17).Value = GreatVol

    'Formatting the percentages
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Adjusting the columns to fit the data width
    ws.Columns.AutoFit

Next ws

End Sub