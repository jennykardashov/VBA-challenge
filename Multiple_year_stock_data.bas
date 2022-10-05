Attribute VB_Name = "Module1"
Sub tickerSummaries()

     Dim i, tickerIndex As Long
     Dim tickerOpen, tickerClose, tickerVolume, bigWinner, bigLoser, bigVolume As Double
     Dim bigWinnerTick, bigLoserTick, bigVolumeTick As String
     Dim newSection As Integer
    
     For Each ws In Worksheets
        tickerIndex = 2
        newSection = 1
        bigWinner = 0
        bigLoser = 0
        bigVolume = 0
            For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
                If newSection = 1 Then
                    tickerOpen = ws.Cells(i, 3).Value
                End If
                newSection = 0
                tickerVolume = tickerVolume + ws.Cells(i, 7)
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' Last row of this ticker, print summary
                    ws.Cells(1, 10).Value = "Ticker"
                    ws.Cells(1, 11).Value = "Yearly Change"
                    ws.Cells(1, 12).Value = "% Change"
                    ws.Cells(1, 13).Value = "Total Stock Volume"
                    tickerClose = ws.Cells(i, 6).Value
                    ws.Cells(tickerIndex, 10).Value = ws.Cells(i, 1).Value
                    ws.Cells(tickerIndex, 11).Value = tickerClose - tickerOpen
                    ws.Cells(tickerIndex, 12).Value = ((tickerClose - tickerOpen) / tickerOpen) * 100
                    ws.Cells(tickerIndex, 13).Value = tickerVolume
                    If tickerVolume > bigVolume Then
                        bigVolumeTick = ws.Cells(i, 1).Value
                        bigVolume = tickerVolume
                    End If
                    If ws.Cells(tickerIndex, 11).Value > 0 Then
                        ' positive change
                        ws.Cells(tickerIndex, 11).Interior.ColorIndex = 4
                        If ws.Cells(tickerIndex, 12).Value > bigWinner Then
                            bigWinnerTick = ws.Cells(i, 1).Value
                            bigWinner = ws.Cells(tickerIndex, 12).Value
                        End If
                    Else
                        ' negative change
                        ws.Cells(tickerIndex, 11).Interior.ColorIndex = 3
                        If ws.Cells(tickerIndex, 12).Value < bigLoser Then
                            bigLoserTick = ws.Cells(i, 1).Value
                            bigLoser = ws.Cells(tickerIndex, 12).Value
                        End If
                    End If
                    tickerIndex = tickerIndex + 1
                    tickerVolume = 0
                    newSection = 1
                End If
            Next i
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(2, 16).Value = bigWinnerTick
            ws.Cells(2, 17).Value = bigWinner
            ws.Cells(3, 16).Value = bigLoserTick
            ws.Cells(3, 17).Value = bigLoser
            ws.Cells(4, 16).Value = bigVolumeTick
            ws.Cells(4, 17).Value = bigVolume
            ws.Columns("J:Q").AutoFit
    Next ws

End Sub

