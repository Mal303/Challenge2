Attribute VB_Name = "Module1"
Sub CalculateQuarterlyStockDataAllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim currentRow As Long
    Dim outputRow As Long
    Dim quarterStartRow As Long
    
    ' Variables to track the greatest changes and volumes
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        
        ' Initialize tracking variables for each sheet
        maxIncrease = -1000000
        maxDecrease = 1000000
        maxVolume = 0
        
        ' Add headers to the specified columns
        ws.Cells(1, 9).Value = "Ticker" ' Column I
        ws.Cells(1, 10).Value = "Quarterly Change" ' Column J
        ws.Cells(1, 11).Value = "Percentage Change" ' Column K
        ws.Cells(1, 12).Value = "Total Stock Volume" ' Column L
        ws.Cells(1, 16).Value = "Ticker" ' Column P
        ws.Cells(1, 17).Value = "Value" ' Column Q
        
        ' Add headers for the greatest values
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        
        ' Get the last row of data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize the output row, start outputting in row 2
        outputRow = 2
        
        ' Loop through all the rows
        currentRow = 2 ' Assuming headers are in row 1
        Do While currentRow <= lastRow
            If ws.Cells(currentRow, 1).Value <> "" Then
                ticker = ws.Cells(currentRow, 1).Value ' Assume ticker is in column A
                quarterStartRow = currentRow
                
                ' Move currentRow to the last occurrence of the same ticker (end of the quarter)
                Do While ws.Cells(currentRow, 1).Value = ticker And currentRow <= lastRow
                    currentRow = currentRow + 1
                Loop
                
                ' Get the open price (first row of the quarter)
                openPrice = ws.Cells(quarterStartRow, 3).Value ' Assume open price is in column C
                
                ' Get the close price (last row of the quarter)
                closePrice = ws.Cells(currentRow - 1, 6).Value ' Assume close price is in column F
                
                ' Calculate the quarterly change
                quarterlyChange = closePrice - openPrice
                
                ' Calculate the percentage change
                If openPrice <> 0 Then
                    percentageChange = (quarterlyChange / openPrice)
                Else
                    percentageChange = 0
                End If
                
                ' Calculate the total volume for the quarter
                totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(quarterStartRow, 7), ws.Cells(currentRow - 1, 7))) ' Assume volume is in column G
                
                ' Output the results
                ws.Cells(outputRow, 9).Value = ticker ' Output to column I
                ws.Cells(outputRow, 10).Value = quarterlyChange ' Output to column J
                
                ' Color code column J based on positive (green) or negative (red) change
                If quarterlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(144, 238, 144) ' Light green
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.Color = RGB(255, 99, 71) ' Light red
                End If
                
                ws.Cells(outputRow, 11).Value = percentageChange ' Output to column K
                
                ' Format the percentage in column K as a percent with two decimal places
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                ws.Cells(outputRow, 12).Value = totalVolume ' Output to column L
                
                ' Track the greatest percentage increase
                If percentageChange > maxIncrease Then
                    maxIncrease = percentageChange
                    maxIncreaseTicker = ticker
                End If
                
                ' Track the greatest percentage decrease
                If percentageChange < maxDecrease Then
                    maxDecrease = percentageChange
                    maxDecreaseTicker = ticker
                End If
                
                ' Track the greatest total volume
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If
            End If
            
            ' Move to the next output row and process next row
            outputRow = outputRow + 1
        Loop
        
        ' Output the greatest values for the current sheet
        ws.Cells(2, 17).Value = maxIncreaseTicker
        ws.Cells(3, 17).Value = maxDecreaseTicker
        ws.Cells(4, 17).Value = maxVolumeTicker
        
        ' Output the values associated with the greatest changes and volumes
        ws.Cells(2, 18).Value = maxIncrease
        ws.Cells(2, 18).NumberFormat = "0.00%" ' Format as percentage
        ws.Cells(3, 18).Value = maxDecrease
        ws.Cells(3, 18).NumberFormat = "0.00%" ' Format as percentage
        ws.Cells(4, 18).Value = maxVolume
        
    Next ws
    
    MsgBox "All Finished!"

End Sub


