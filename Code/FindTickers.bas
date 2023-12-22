Attribute VB_Name = "Module1"
Sub FindTickers()
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set initial variables
        Dim LastRow As Long
        Dim OutputRow As Long
        Dim i As Long

        ' Find the last row of data in the worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set initial output table row
        OutputRow = 2

        ' Set column header for output table
        ws.Cells(1, 9).Value = "Ticker"

        ' Loop through all rows of data starting from row 2
        For i = 2 To LastRow
            ' Check if the ticker symbol has changed from one cell to the next
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' If the ticker symbol has changed, copy the ticker to the output column
                ws.Cells(OutputRow, 9).Value = ws.Cells(i, 1).Value

                ' Increment the summary table row
                OutputRow = OutputRow + 1
            End If
        Next i
    Next ws
End Sub

Sub FindTickersNoText()
    For Each ws In ThisWorkbook.Worksheets
        Dim LastRow As Long
        Dim OutputRow As Long
        Dim i As Long

        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        OutputRow = 2

        ws.Cells(1, 9).Value = "Ticker"

        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(OutputRow, 9).Value = ws.Cells(i, 1).Value
                OutputRow = OutputRow + 1
            End If
        Next i
    Next ws
End Sub

