Attribute VB_Name = "ProcessQuarterlyTickerData"
' Set row index constants for readability.
Const headingsRow As Long = 1
Const greatestIncreaseRow As Long = 2
Const greatestDecreaseRow As Long = 3
Const greatestVolumeRow As Long = 4

' Set column index constants for readability.
Const tickerCol As Long = 1
Const openCol As Long = 3
Const closeCol As Long = 6
Const volumeCol As Long = 7
Const outTickerCol As Long = 9
Const outChangeCol As Long = 10
Const outPercentageCol As Long = 11
Const outVolumeCol As Long = 12
Const summaryCategoryCol As Long = 15
Const summaryTickerCol As Long = 16
Const summaryValueCol As Long = 17

' Set color value constants for readability.
Const greenColor As Long = 65280
Const redColor As Long = 255

Sub ProcessAllQuarterlyWorksheets()
    ' Process all worksheets in the workbook and then go back to the first worksheet.
    
    Dim ws As Worksheet
    
    ' For each worksheet in the workbook, activate the worksheet and process it,
    For Each ws In Worksheets
        ws.Activate
        ProcessSingleQuarterlyWorksheet
    Next ws
    
    ' then go back to the first sheet (`Worksheets` is 1 indexed).
    Worksheets(1).Activate
End Sub

Sub ProcessSingleQuarterlyWorksheet()
    ' Iterate through the worksheet's source data, compute the quarterly changes for each
    ' ticker.  Keep track of the first open price and the last close price as well as the
    ' total volume of shares traded for each ticker.  Using the first open and last close
    ' compute the quarterly price change, and the percentage change compared to the first
    ' open.  Keep track of the ticker and value for the greatest % increase and decrease,
    ' and the total volume traded.
    '
    ' Fill columns I to L with the quarterly change, percent change (relative to the first
    ' open price), and the total volume of shares traded.
    '
    ' Fill columns O to Q with summary information including the summary category, ticker,
    ' and value for the following categories: Greatest % Increase, Greatest % Decrease,
    ' and Greatest Total Volume.
    '
    ' Then, format the sheet properly, including auto-fitting the output and summary, and
    ' adding conditional formatting to the output Quarterly Change (column J from J2
    ' downwards) so that cells are filled with green color when positive, red color when
    ' negative, and no fill color if zero.
    '
    ' Finally, go back to the top left cell (A1).
    
    ' Clear output area, and go to top left cell.
    Columns("H:Q").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Clear
    
    ' Insert Quarterly Summary Heading.
    Cells(headingsRow, outTickerCol).value = "Ticker"
    Cells(headingsRow, outChangeCol).value = "Quarterly Change"
    Cells(headingsRow, outPercentageCol).value = "Percent Change"
    Cells(headingsRow, outVolumeCol).value = "Total Stock Volume"
    
    ' Determine source data start, end, out, and summary rows.
    Dim startRow As Long
    Dim endRow As Long
    Dim outRow As Long
    Dim summaryRow As Long
    
    startRow = 2
    endRow = Cells(Rows.Count, 1).End(xlUp).row
    outRow = 2
    summaryRow = 2
    
    ' Iterate over all source data rows.
    Dim thisRow As Long
    Dim nextRow As Long
    
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    
    Dim nextTicker As String
    Dim isTickerBoundary As Boolean
    Dim tickerStartRow As Long
    Dim tickerEndRow As Long
    
    ' Declare output variables.
    Dim change As Double
    Dim percentage As Double
    
    ' Declare summary variables.
    Dim greatestIncreaseTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestVolumeTicker As String
    Dim greatestVolume As Double
    
    ' Initialize the ticker start and end rows for the first ticker,
    tickerStartRow = startRow
    tickerEndRow = endRow
    
    ' then the volume accumulator to 0,
    volume = 0
    
    ' and then the summary variables: ticker to empty string, value to 0.
    greatestIncreaseTicker = ""
    greatestIncrease = 0
    greatestDecreaseTicker = ""
    greatestDecrease = 0
    greatestVolumeTicker = ""
    greatestVolume = 0
    
    ' Iterate over all source data rows.
    For thisRow = startRow To endRow
        ' Compute the next row.
        nextRow = thisRow + 1
        
        ' Get this and next row's tickers, and this row's volume.
        ticker = Cells(thisRow, tickerCol).value
        nextTicker = Cells(nextRow, tickerCol).value
        volume = volume + Cells(thisRow, volumeCol).value
        
        ' Check if this is the last row for this ticker.
        isTickerBoundary = (thisRow = endRow) Or (ticker <> nextTicker)
        
        ' If we're at a ticker boundary, we have some stuff to do.
        If isTickerBoundary Then
            ' Set the ticker end row.
            tickerEndRow = thisRow
            
            ' Get this ticker's open and close price.
            openPrice = Cells(tickerStartRow, openCol).value
            closePrice = Cells(tickerEndRow, closeCol).value
            
            ' Compute the change and percentage.
            change = closePrice - openPrice
            percentage = change / openPrice
            
            ' "Output" the ticker, change, percentage, and volume for this ticker.
            Cells(outRow, outTickerCol).value = ticker
            Cells(outRow, outChangeCol).value = change
            Cells(outRow, outPercentageCol).value = percentage
            Cells(outRow, outVolumeCol).value = volume
            
            ' Update summaries if applicable.
            
            ' First the greatest increase,
            If percentage > greatestIncrease Then
                greatestIncreaseTicker = ticker
                greatestIncrease = percentage
            End If
            
            ' then the greatest decrease,
            If percentage < greatestDecrease Then
                greatestDecreaseTicker = ticker
                greatestDecrease = percentage
            End If
            
            ' and lastly the greatest total volume.
            If volume > greatestVolume Then
                greatestVolumeTicker = ticker
                greatestVolume = volume
            End If
            
            ' Increment the output row to prepare for the next ticker's output.
            outRow = outRow + 1
            
            ' Reset the volume accumulator to zero.
            volume = 0
            
            ' Re-initialize the ticker start row for the next ticker.
            tickerStartRow = nextRow
        End If
    Next thisRow
    
    ' Set quarter summary headings and values
    
    ' first the headings,
    Cells(headingsRow, summaryCategoryCol).value = ""
    Cells(headingsRow, summaryTickerCol).value = "Ticker"
    Cells(headingsRow, summaryValueCol).value = "Value"
    
    ' then the greatest increase,
    Cells(greatestIncreaseRow, summaryCategoryCol).value = "Greatest % Increase"
    Cells(greatestIncreaseRow, summaryTickerCol).value = greatestIncreaseTicker
    Cells(greatestIncreaseRow, summaryValueCol).value = greatestIncrease
    
    ' then the greatest decrease,
    Cells(greatestDecreaseRow, summaryCategoryCol).value = "Greatest % Decrease"
    Cells(greatestDecreaseRow, summaryTickerCol).value = greatestDecreaseTicker
    Cells(greatestDecreaseRow, summaryValueCol).value = greatestDecrease
    
    ' and finally the greatest total volume.
    Cells(greatestVolumeRow, summaryCategoryCol).value = "Greatest Total Volume"
    Cells(greatestVolumeRow, summaryTickerCol).value = greatestVolumeTicker
    Cells(greatestVolumeRow, summaryValueCol).value = greatestVolume
    
    ' Format the output and summary.
    AddQuarterlyChangeConditionalFormatting
    FormatOutput
    FormatSummary
    
    ' Go to the top left cell.
    Cells(1, 1).Select
End Sub

Sub AddQuarterlyChangeConditionalFormatting()
    ' Add conditional formatting to the quarterly change output column (J from J2 down).  Set
    ' the fill to solid green color if the change is positive, solid red if it's negative,
    ' or no fill if there was no (zero) change.
    
    ' Clear conditional formatting in this sheet.
    Cells.FormatConditions.Delete
        
    ' Select quarterly change range (column J, from J2 down).
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    ' Add conditional formatting to the selection in reverse, so we can easily set the
    ' order.  Assuming most changes will be positive, then negative, then zero, add zero
    ' condition first, then negative, then positive, setting each to first priority in turn.
    ' Since a cell is either negative, zero, or positive exclusively, set each role to stop
    ' further processing if the condition is true.
    
    ' First, if cell value equal to zero, then set to no pattern.
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    ' Then, if the cell value is negative, then set the fill to solid red.
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = redColor
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    ' Lastly, if the cell value is positive, then set the fill to solid green.
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = greenColor
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = True
End Sub

Sub FormatOutput()
    ' Format the output area (excluding conditional formatting).  Set text data column
    ' (ticker column) alignment to bottom left aligned, and numeric data columns (quarterly
    ' change, percentage change, and total stock volume columns) alignments to right aligned.
    ' The quarterly change should be numeric with comma separation, 2 digit precision, and
    ' a genative sign prefix for negative values.  The percent change should be a percentage
    ' format with 2 decimal digit precision, and the total stock volume should be numeric
    ' with no comma separation. All output columns (I to L) should be auto-fit to make their
    ' content fit in the visible width allotted to them.
    
    ' Format output ticker column (including header) to be left aligned.
    Columns("I:I").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Format output quarterly change column to numeric with "," as the thousands separator,
    ' and two digit decimal precision.
    Columns("J:J").Select
    Selection.NumberFormat = "#,##0.00"
    
    ' Format output quarterly change percentage column to percentage with two decimal digit
    ' precision.
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    
    ' Format the total stock volume column to numeric with no thousands separator and no
    ' decimal digits (whole numbers only).
    Columns("L:L").Select
    Selection.NumberFormat = "#0.00"
    
    ' Right-align the numerical output columns J to L (quarterly change, percentage change,
    ' and total stock volume).
    Columns("J:L").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Auto-fit the output columns I to L (ticker, quarterly change, percentage change, and
    ' total stock volume).
    'Columns("I:L").Select
    Columns("I:L").EntireColumn.AutoFit
End Sub

Sub FormatSummary()
    ' Format the summary area.  Left aligns the summary category and ticker columns, right
    ' align the summary value column, set the greatest increase and decrease value cells
    ' to percentage format with 2 decimal digits of precision, and auto-fit all summary
    ' columns.
    
    ' Left align the summary category and ticker columns (columns O to P).
    Columns("O:P").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Right align the summary value column (column Q).
    Columns("Q:Q").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Set the format for the greatest increase and decrease percentage values to percentage
    ' with two decimal digit precision (cells Q2 and Q3).
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    ' Auto-fit the summary columns (columns O to Q).
    'Columns("O:Q").Select
    Columns("O:Q").EntireColumn.AutoFit
End Sub

Sub DebugClearAllSheets()
    ' Clear the output and summary areas on all sheets.
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
        
        ' Clear output area.
        Columns("H:Q").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Clear
    
        Range("A1").Select
    Next ws
    
    Worksheets(1).Activate
End Sub
