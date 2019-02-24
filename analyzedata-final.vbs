Sub AnalyzeData()

    For Each ws In Worksheets
    
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        Dim numRows As Long
        Dim tickerCount As Integer
        Dim tickerVolume As Double
        numRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        tickerCount = 0
        tickerVolume = 0
        
        Dim yearOpen As Double
        Dim yearClose As Double
        
        ' stores first stock's opening price
        yearOpen = ws.Cells(2, 3)
        ' loop through all the rows
        For i = 2 To numRows
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                tickerCount = tickerCount + 1
                ' display the ticker symbol to coincide with the total volume
                ws.Cells(tickerCount + 1, 9) = ws.Cells(i, 1).Value
                ' total amount of volume each stock had over the year
                ws.Cells(tickerCount + 1, 12) = tickerVolume + ws.Cells(i, 7).Value
                ' Yearly change from what the stock opened the year at to what the closing price was
                yearClose = ws.Cells(i, 6).Value
                ws.Cells(tickerCount + 1, 10) = yearClose - yearOpen
                ' percent change from the what it opened the year at to what it closed.
                If yearOpen > 0 Then ' to avoid dividing by zero
                    ws.Cells(tickerCount + 1, 11) = (yearClose - yearOpen) / yearOpen
                ElseIf yearOpen = 0 And yearClose = 0 Then
                    ' if yearOpen and yearClose are both zero
                    ws.Cells(tickerCount + 1, 11) = 0
                Else
                    ' find first non-zero yearOpen value for current ticker
                    j = 0
                    For j = 0 To i
                        yearOpen = ws.Cells(i - j, 3).Value
                        If yearOpen > 0 Then
                            Exit For
                        End If
                     Next j
                    ws.Cells(tickerCount + 1, 11) = (yearClose - yearOpen) / yearOpen
                End If
                ' highlight positive change in green and negative change in red
                If ws.Cells(tickerCount + 1, 10) < 0 Then
                    ws.Cells(tickerCount + 1, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(tickerCount + 1, 10).Interior.ColorIndex = 4
                End If
                ' reset
                tickerVolume = 0
                yearClose = 0
                ' stores next stock's opening price
                yearOpen = ws.Cells(i + 1, 3).Value
            Else
                tickerVolume = tickerVolume + ws.Cells(i, 7).Value
            End If

        Next i
        ' ws.Range("J1:J" & tickerCount + 1).NumberFormat = "0.00000000"
        ws.Range("K1:K" & tickerCount + 1).NumberFormat = "0.00%"
    
        ' locate the stock with the "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Volume"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim maxIncreaseRow As Integer
        Dim maxDecreaseRow As Integer
        Dim maxVolumeRow As Integer
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        For i = 2 To tickerCount + 1
                If ws.Cells(i, 12) > maxVolume Then
                    maxVolume = ws.Cells(i, 12)
                    maxVolumeRow = i
                End If
                If ws.Cells(i, 11) > maxIncrease Then
                    maxIncrease = ws.Cells(i, 11)
                    maxIncreaseRow = i
                End If
                If ws.Cells(i, 11) < maxDecrease Then
                    maxDecrease = ws.Cells(i, 11)
                    maxDecreaseRow = i
                End If
        Next i
        ' Display Greatest % increase, % decrease, total volume
        ws.Cells(2, 16) = ws.Cells(maxIncreaseRow, 9)
        ws.Cells(2, 17) = ws.Cells(maxIncreaseRow, 11)
        ws.Cells(3, 16) = ws.Cells(maxDecreaseRow, 9)
        ws.Cells(3, 17) = ws.Cells(maxDecreaseRow, 11)
        ws.Cells(4, 16) = ws.Cells(maxVolumeRow, 9)
        ws.Cells(4, 17) = ws.Cells(maxVolumeRow, 12)
        ws.Columns("J:Q").AutoFit
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    Next ws
End Sub

