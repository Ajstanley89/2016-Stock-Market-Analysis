Sub EasyAndModerate()

Dim Ticker As String
Dim SummaryTickerColumn As Integer
Dim SummaryPriceDiffColumn As Integer
Dim SummaryPricePercentColumn As Integer
Dim TotalStockVolumeColumn As Integer
Dim SummaryTableRow As Integer

'loop through all worksheets
For Each ws In Worksheets

    'only print headers if worksheet name isn't "Buttons"
    If ws.Name <> "Buttons" Then
        'determine last Row in each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'determine last column in each worksheet
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        'set sumamry table row to row 1
        SummaryTableRow = 1

        'determine location of start of summary table. Starts with Ticker column, and there should be one blank column in between the data and this column
        SummaryTickerColumn = LastColumn + 2

        'create sumary ticker header
        ws.Cells(SummaryTableRow, SummaryTickerColumn).Value = "Ticker"

        'Fill in summary rows to the left of the ticker column and create column label
        SummaryPriceDiffColumn = SummaryTickerColumn + 1
        ws.Cells(SummaryTableRow, SummaryPriceDiffColumn).Value = "Yearly Change"

        'Percent Change column to the left of yearly change and create header
        SummaryPricePercentColumn = SummaryPriceDiffColumn + 1
        ws.Cells(SummaryTableRow, SummaryPricePercentColumn).Value = "Percent Change"

        'Total Stock Column and Create header
        TotalStockVolumeColumn = SummaryPricePercentColumn + 1
        ws.Cells(SummaryTableRow, TotalStockVolumeColumn).Value = "Total Stock Volume"

        'Set summary table row to next row
        SummaryTableRow = SummaryTableRow + 1
    End If

    'Column 1 is ticker, column 3 is opening price, column 6 is closing price, column 7 is volume

    'loop through all rows and find the year opening price, year closing price, and total volume of stock traded.
    'I'm using Brute force, baby! In the future, this code could be streamlined using filters. I could filter by each unique ticker name to sum the stock volume, then grab the first and last entry for each ticker to get the value change.
    For i = 2 To LastRow
        'check if this is the first entry for a ticker. This is the first entry if the ticker doesn't match the one before it
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            'Find the opening price
            YearOpeningPrice = ws.Cells(i, 3).Value

            'find the volume of stock traded the first time
            TotalStockVolume = ws.Cells(i, 7).Value

        'check if this is the last entry for this stock. It's the last entry if the ticker doesn't match the next one in the list
        ElseIf ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then

            'Add the volume of stock traded that day to the running total
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        End If

        'Check if this is the last entry for a stock. Check if the ticker doesn't match the one after it. This is a seperate if statement in case there is only a single entry for a stock.
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'Add the last volume to the running total
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

            'find the year closing price
            YearClosingPrice = ws.Cells(i, 6).Value

            'Find the difference between open and close
            YearDiff = YearClosingPrice - YearOpeningPrice

            'Find the percent difference between open and close

            'weed out if the opening price is 0
            If YearOpeningPrice = 0 Then

                YearPercent = 0

            Else
                YearPercent = YearDiff / YearOpeningPrice

            End If

            'Find the ticker
            Ticker = ws.Cells(i, 1).Value

            'Populate summary table
            'Ticker in table
            ws.Cells(SummaryTableRow, SummaryTickerColumn).Value = Ticker

            'Yearly Change in table
            ws.Cells(SummaryTableRow, SummaryPriceDiffColumn).Value = YearDiff

            'make cell green if value is greater than 0
            If YearDiff > 0 Then

                ws.Cells(SummaryTableRow, SummaryPriceDiffColumn).Interior.ColorIndex = 4

            'make cell red if value is less than 0
            ElseIf YearDiff < 0 Then

                ws.Cells(SummaryTableRow, SummaryPriceDiffColumn).Interior.ColorIndex = 3

            End If

            'Percent Change in table
            ws.Cells(SummaryTableRow, SummaryPricePercentColumn).Value = YearPercent

            'add percent sign formatting to percent change column
            ws.Cells(SummaryTableRow, SummaryPricePercentColumn).Style = "Percent"

            'Total Stock Volume in Table
            ws.Cells(SummaryTableRow, TotalStockVolumeColumn).Value = TotalStockVolume

            'Move on to next sumamry table row
            SummaryTableRow = SummaryTableRow + 1

        End If
    Next i
Next ws

End Sub

Sub Hard()

'dims for hard solution
Dim MaxPercent As Double
Dim MaxPerTicker As String
Dim MinPercent As Double
Dim MinPerTicker As String
Dim MaxVolume As Double
Dim MaxVolTicker As String
Dim LastColumn As Integer
Dim SummaryTableRow As Integer
Dim SummaryTickerColumn As Integer
Dim SummaryPriceDiffColumn As Integer
Dim SummaryPricePercentColumn As Integer
Dim TotalStockVolumeColumn As Integer
Dim GreatestColumn As Integer
Dim GreatestTickerColumn As Integer
Dim GreatestValueColumn As Integer
Dim GreatestRow As Integer

'loop through all worksheets
For Each ws In Worksheets

    'determine last column in each worksheet
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    'Set rows for summary column. In summary row, ticker is col 9, yearly change is 10, percent change is 11, volume is 12

    SummaryTickerColumn = 9
    SummaryPriceDiffColumn = 10
    SummaryPricePercentColumn = 11
    TotalStockVolumeColumn = 12

    'find last used row in summary range

    SummaryTableRow = ws.Range("I" & Rows.Count).End(xlUp).Row

    'set max percent, min percent, and max volume equal to the first ticker

    MaxPercent = ws.Cells(2, SummaryPricePercentColumn).Value

    MaxPerTicker = ws.Cells(2, SummaryTickerColumn).Value

    MinPercent = ws.Cells(2, SummaryPriceDiffColumn).Value

    MinPerTicker = ws.Cells(2, SummaryTickerColumn).Value

    MaxVolume = ws.Cells(2, TotalStockVolumeColumn).Value

    MaxVolTicker = MinPercent = ws.Cells(2, SummaryPriceDiffColumn).Value

    'Can try using formula later: MaxPercent = ws.WorksheetFunction.Max(Range(& SummaryPricePercentColumn).Value)

    'Loop through summary table starting with the second row. Compare the values to the first row
    For i = 3 To SummaryTableRow
        'check if Percent change so far is smaller than the current row
        If MaxPercent < ws.Cells(i, SummaryPricePercentColumn).Value Then
            'if the percent change value is smaller than the current value, then update the MaxPercent Variable
            MaxPercent = ws.Cells(i, SummaryPricePercentColumn).Value
            'update ticker for max
            MaxPerTicker = ws.Cells(i, SummaryTickerColumn).Value
        End If

        'check if Percent change so far is larger than the current row
        If MinPercent > ws.Cells(i, SummaryPricePercentColumn).Value Then
            'if the prercent change value is larger than the current value, then update the MinPercent Variable
            MinPercent = ws.Cells(i, SummaryPricePercentColumn).Value
            'update ticker for min
            MinPerTicker = ws.Cells(i, SummaryTickerColumn).Value
        End If

        'check if Total volume so far is larger than the current row
        If MaxVolume < ws.Cells(i, TotalStockVolumeColumn).Value Then
            'if the total volume value is smaller than the current value, then update the MaxVolume Variable
            MaxVolume = ws.Cells(i, TotalStockVolumeColumn).Value
            'update ticker for max volume
            MaxVolTicker = ws.Cells(i, SummaryTickerColumn).Value
        End If
    Next i
    
    'Start new table with one blank space between summary table
    GreatestColumn = TotalStockVolumeColumn + 2
    GreatestTickerColumn = GreatestColumn + 1
    GreatestValueColumn = GreatestTickerColumn + 1
    GreatestRow = 1

'only print headers if worksheet name isn't "Buttons"
    If ws.Name <> "Buttons" Then
        'Greatest values table header
        ws.Cells(GreatestRow, GreatestTickerColumn).Value = "Ticker"
        ws.Cells(GreatestRow, GreatestValueColumn).Value = "Value"

        'Populate table with values
        'Category
        ws.Cells(GreatestRow + 1, GreatestColumn) = "Greatest % increase"
        ws.Cells(GreatestRow + 2, GreatestColumn) = "Greatest % decrease"
        ws.Cells(GreatestRow + 3, GreatestColumn) = "Greatest Total Volume"

        'ticker values
        ws.Cells(GreatestRow + 1, GreatestTickerColumn) = MaxPerTicker
        ws.Cells(GreatestRow + 2, GreatestTickerColumn) = MinPerTicker
        ws.Cells(GreatestRow + 3, GreatestTickerColumn) = MaxVolTicker

        'values
        'print max
        ws.Cells(GreatestRow + 1, GreatestValueColumn) = MaxPercent
        'format as percent
        ws.Cells(GreatestRow + 1, GreatestValueColumn).Style = "Percent"

        'print min
        ws.Cells(GreatestRow + 2, GreatestValueColumn) = MinPercent
        'format as percent
        ws.Cells(GreatestRow + 2, GreatestValueColumn).Style = "Percent"

        'print volume
        ws.Cells(GreatestRow + 3, GreatestValueColumn) = MaxVolume
        
        End If
Next ws

End Sub

Sub Clear()

Dim LastColumn As Integer
Dim LastRow As Long
For Each ws In Worksheets

    'clear all cells except the given data. Columns I thru P, Rows 1 thru last row

    'determine last column in each worksheet
    LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

    'determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Clear Contents
    ws.Range("I1:P" & LastRow).ClearContents
    
    'Clear formatting
    ws.Range("I1:P" & LastRow).ClearFormats

Next ws

End Sub



