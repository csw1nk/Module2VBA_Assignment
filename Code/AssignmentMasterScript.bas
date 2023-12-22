Attribute VB_Name = "AnalyzeStockData"
Sub AnalyzeStockData()
    ' used this source for helping shape this query - https://stackoverflow.com/questions/77682363/vba-script-returning-for-without-next
    ' Loop through all worksheets in the workbook - reference: https://excelchamps.com/vba/loop-sheets/
    For Each ws In ThisWorkbook.Worksheets
        ' Set initial variables to declare later in script - reference in class
        ' using Double for percentage change as that is decimals - reference in class
        ' using Long for last row and i and output row due to volume of data - reference: https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/data-types/long-data-type
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim LastRow As Long
        Dim OutputRow As Long
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim i As Long

        ' Set column headers for output table - reference in class
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Find the last row of data in the worksheet reference in class
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set initial output table row for space for headers reference in class
        OutputRow = 2
        ' Set initial opening price for the first stock ticker - we want the open price to be the first record for each ticket
        OpenPrice = ws.Cells(2, 3).Value
        ' Loop through all rows of data starting from row 2 reference in class
        For i = 2 To LastRow
        ' Check if the ticker symbol has changed from one cell to the next
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ' Paste values into output table
            Ticker = ws.Cells(i, 1).Value
        ' Calculate the yearly change and set open and close values for percent and yearly change
            ClosePrice = ws.Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice
                ' Calculate the percent change
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) 'dont need to *100 because im doing it below
                Else
                    PercentChange = 0
                End If
                ' Add everything to the output table
                ws.Cells(OutputRow, 9).Value = Ticker
                ws.Cells(OutputRow, 10).Value = YearlyChange
                ws.Cells(OutputRow, 11).Value = PercentChange
                ws.Cells(OutputRow, 12).Value = TotalVolume

                ' Format the percent change as a percentage - reference: https://www.statology.org/vba-percentage-format/
                ws.Cells(OutputRow, 11).NumberFormat = "0.00%"

                ' Add conditional formatting for positive and negative values - green, red, and grey - reference: https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
                If YearlyChange > 0 Then
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = 3
                ElseIf YearlyChange = 0 Then
                    ws.Cells(OutputRow, 10).Interior.ColorIndex = 16
                End If

                ' Increment the summary table row
                OutputRow = OutputRow + 1

                ' Reset the total volume
                ' This is necessary because TotalVolume accumulates the volume for each ticker, and must start from 0 for the next ticker.
                TotalVolume = 0
                  ' Update the OpenPrice for the next ticker
                  ' checks if i is not the last row then if true set the open price
                  ' through logic we are finding the open price for each stock and the close price, i.e. first record of open and last record of close
            If i + 1 <= LastRow Then
                OpenPrice = ws.Cells(i + 1, 3).Value
            End If
            Else
                ' Add the stock volume to the total volume
                 'This line adds the volume of the current row (i) to the running total - TotalVolume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub


Sub FindStockHighlights()
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        Dim LastOutputRow As Long
        Dim MaxIncrease As Double, MaxDecrease As Double, MaxVolume As Double
        Dim MaxIncreaseTicker As String, MaxDecreaseTicker As String, MaxVolumeTicker As String

        ' Set the initial values for maximum increase, decrease, and volume because again we can incrementing from first row to last
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0

        ' Find the last row of the summary table not data table
        LastOutputRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

        ' Loop through each row in the summary table
        Dim i As Long
        For i = 2 To LastOutputRow
            ' Checking for maximum increase
            If ws.Cells(i, 11).Value > MaxIncrease Then
                MaxIncrease = ws.Cells(i, 11).Value
                MaxIncreaseTicker = ws.Cells(i, 9).Value
            End If
            ' Checking for maximum decrease
            If ws.Cells(i, 11).Value < MaxDecrease Then
                MaxDecrease = ws.Cells(i, 11).Value
                MaxDecreaseTicker = ws.Cells(i, 9).Value
            End If
            ' Checking for maximum volume
            If ws.Cells(i, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(i, 12).Value
                MaxVolumeTicker = ws.Cells(i, 9).Value
            End If
        Next i

        ' Easter Egg: If the maximum increase percentage is exactly 100% (which is highly unlikely),
    ' this line will trigger a fun message.
    If MaxIncrease = 100 Then
        MsgBox "100% increase? Time to go invest indeed!"
    End If

        ' Output the results to the worksheet cells and instructions
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(2, 17).Value = MaxIncrease
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = MaxDecreaseTicker
        ws.Cells(3, 17).Value = MaxDecrease
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(4, 17).Value = MaxVolume

        ' Format the percent increase and decrease as percentages
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Columns("A:Q").AutoFit
    Next ws
End Sub

