Sub DQAnalysis()
'*****Ignore this first subroutine, this was from the module activities*****
'Compute total volume and yearly return for DQ stock in 2018

    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"


    'Find total volume and yearly return of "DQ" stock in 2018
    Worksheets("2018").Activate

    rowStart = 2
    'Find the number of rows to loop over, source:
    'https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    For i = rowStart To rowEnd
        'Increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If

        'Find starting prices for "DQ" in 2018
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            startingPrice = Cells(i, 6).Value
        End If
        'Find ending prices for "DQ" in 2018
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            endingPrice = Cells(i, 6).Value
        End If
    Next i

    'Print results to the worksheet
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice - startingPrice) / startingPrice

End Sub


Sub AllStocksAnalysis()
'Compute total volume and yearly return for all stocks in a given year

    'Timing variables
    Dim startTime As Single
    Dim endTime As Single

    'Collect user input for analysis
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Assemble array of ticker strings
    Dim tickers(11) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    Worksheets(yearValue).Activate
    'Starting row number
    rowStart = 2
    'Find the number of rows to loop over
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Loop 11 times, once for each ticker
    For stock = 0 To 11
        ticker = tickers(stock)
        totalVolume = 0

        Worksheets(yearValue).Activate
        'r for row
        For r = rowStart To rowEnd
            'Increase totalVolume if ticker cell value is equal to the value of
            'ticker
            If Cells(r, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(r, 8).Value
            End If

            'Find starting prices for ticker in a given year
            If Cells(r, 1).Value = ticker And _
                Cells(r - 1, 1).Value <> ticker Then
                startingPrice = Cells(r, 6).Value
            End If
            'Find ending prices for ticker in a given year
            If Cells(r, 1).Value = ticker And _
                Cells(r + 1, 1).Value <> ticker Then
                endingPrice = Cells(r, 6).Value
            End If
        Next r

        'Print results to the worksheet
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + stock, 1).Value = ticker
        Cells(4 + stock, 2).Value = totalVolume
        Cells(4 + stock, 3).Value = _
            (endingPrice - startingPrice) / startingPrice
    Next stock

    endTime = Timer

    'Apply table formatting
    FormatAllStocksAnalysisTable

    MsgBox ("This code ran in " & (endTime - startTime) & _
        " seconds for the year " & (yearValue) & _
        " using the macro AllStocksAnalysis( )")

End Sub


Sub AllStocksAnalysisRefactored()
'Refactored AllStocksAnalysis to decrease code execution time by looping only
'once through all the rows

    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    'Activate data worksheet
    Worksheets(yearValue).Activate

    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

    '2b) Loop over all the rows in the spreadsheet.
    For Row = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + _
                                     Cells(Row, 8).Value

        '3b) Check if the current row is the first row with the selected
        'tickerIndex.
        If Cells(Row, 1).Value = tickers(tickerIndex) And _
            Cells(Row - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(Row, 6).Value
        End If

        '3c) check if the current row is the last row with the selected ticker.
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(Row, 1).Value = tickers(tickerIndex) And _
            Cells(Row + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(Row, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If

    Next Row

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and
    'Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = _
        (tickerEndingPrices(i) - tickerStartingPrices(i)) / _
        tickerStartingPrices(i)
    Next i

    'Formatting code was removed in favor of my own subroutine
    'FormatAllStocksAnalysisTable found below. The formatting of the numbers
    'was kept the same as challenge_starter_code.vbs

    endTime = Timer

    FormatAllStocksAnalysisTable

    MsgBox ("This code ran in " & (endTime - startTime) & _
        " seconds for the year " & (yearValue) & _
        " using the macro AllStocksAnalysisRefactored( )")

End Sub


'--------------------------------------------
' Formatting and Clear Worksheets subroutines
'--------------------------------------------

Sub FormatAllStocksAnalysisTable()
'Format the "All Stocks Analysis" worksheet

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Font.Bold = True
    Range("A1").Font.Size = 14

    With Range("A3:C3")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(234, 243, 250)
    End With

    Range("A4:C15").BorderAround , LineStyle:=xlContinuous
    Range("A4:C15").Borders(xlInsideVertical).LineStyle = xlContinuous

    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"

    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    'r for row, apply formatting per row
    For r = dataRowStart To dataRowEnd
        'Background color applied to every other table row
        If r Mod 2 <> 0 Then
            rstr = CStr(r)
            'Create a string from the loop index r that will be used in Range()
            rangeStr = "A" & rstr & ":B" & rstr
            Range(rangeStr).Interior.Color = RGB(234, 243, 250)
        End If

        'Return column is green if the percentage is positive
        If Cells(r, 3) > 0 Then
            Cells(r, 3).Interior.Color = RGB(142, 246, 152)
        'Return column is red if the percentage is negative
        ElseIf Cells(r, 3) < 0 Then
            Cells(r, 3).Interior.Color = RGB(255, 109, 109)
        'Return column has no background color if the percentage 0
        Else
            Cells(r, 3).Interior.Color = xlNone
        End If
    Next r
End Sub


Sub ClearWorksheetDQ()
'Clear the DQ Analysis worksheet
    Worksheets("DQ Analysis").Activate
    Cells.Clear

End Sub


Sub ClearWorksheetAllStocks()
'Clear the All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Cells.Clear

End Sub

