Attribute VB_Name = "FindStockHighlights"
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
        MsgBox "Wow Time to Invest!"
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