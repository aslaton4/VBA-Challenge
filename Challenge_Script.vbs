Sub stock_analysis()

    ' Set dimensions
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim summaryRow As Long
    Dim quarterlyChange As Double
    Dim percentChange As Double

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        ' Set title row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Set initial values
        summaryRow = 2
        totalVolume = 0
        openPrice = 0
        closePrice = 0

        ' Get the row number of the last row with data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Loop through all the rows in the worksheet
        For i = 2 To lastRow
            ' Check if the ticker is the same or has changed
            If ws.Cells(i, 1).Value <> ticker Then
                ' If ticker changes then print results (but skip for the first ticker)
                If i > 2 Then
                    ' Stores results in variables
                    closePrice = ws.Cells(i - 1, 6).Value ' Get the close price of the previous row
                    quarterlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = (quarterlyChange / openPrice) * 100
                    Else
                        percentChange = 0
                    End If

                    ' Handle zero total volume (if no volume exists for some tickers)
                    If totalVolume = 0 Then
                        totalVolume = 1 ' Prevent division by zero
                    End If

                    ' Print the results
                    ws.Cells(summaryRow, 9).Value = ticker ' Ticker
                    ws.Cells(summaryRow, 10).Value = quarterlyChange ' Quarterly Change
                    ws.Cells(summaryRow, 11).Value = percentChange ' Percent Change
                    ws.Cells(summaryRow, 12).Value = totalVolume ' Total Stock Volume

                    ' Colors positives green and negatives red
                    If quarterlyChange > 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                    ElseIf quarterlyChange < 0 Then
                        ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                    End If

                    ' Move to the next summary row
                    summaryRow = summaryRow + 1
                End If

                ' Reset variables for new stock ticker
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value ' Get the new opening price
                totalVolume = 0 ' Reset the volume for the new ticker
            End If

            ' If ticker is still the same, add to the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value ' Add volume for current row

        Next i

        ' Print results for the last ticker
        closePrice = ws.Cells(lastRow, 6).Value ' Get the close price of the last row
        quarterlyChange = closePrice - openPrice
        If openPrice <> 0 Then
            percentChange = (quarterlyChange / openPrice) * 100
        Else
            percentChange = 0
        End If

        ws.Cells(summaryRow, 9).Value = ticker ' Ticker
        ws.Cells(summaryRow, 10).Value = quarterlyChange ' Quarterly Change
        ws.Cells(summaryRow, 11).Value = percentChange ' Percent Change
        ws.Cells(summaryRow, 12).Value = totalVolume ' Total Stock Volume

        ' Colors positives green and negatives red for the last ticker
        If quarterlyChange > 0 Then
            ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf quarterlyChange < 0 Then
            ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
        End If

    Next ws ' Move to the next worksheet

End Sub
