Sub StockData():

    For Each ws In Worksheets
    
        Dim TickerCount As Long
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestVolume As Double
        Dim i As Long
        Dim j As Long
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        TickerCount = 2
        
        j = 2

        
            For i = 2 To 800000
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                        If ws.Cells(TickerCount, 10).Value < 0 Then
                        ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                        Else
                
                        ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                        End If
                    
                                If ws.Cells(j, 3).Value <> 0 Then
                                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                                ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                    
                                Else
                    
                                ws.Cells(TickerCount, 11).Value = 0
                    
                                End If
                    
                                    ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                                    
                                    TickerCount = TickerCount + 1
                
                                    j = i + 1
                
                End If
            
                Next i
            
        
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        GreatestVolume = ws.Cells(2, 12).Value
        
            For i = 2 To 10000
            

                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestIncrease = GreatestIncrease
                
                End If
                
                    If ws.Cells(i, 11).Value < GreatestDecrease Then
                    GreatestDecrease = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                    Else
                
                    GreatestDecrease = GreatestDecrease
                
                    End If
                        
                        
                        If ws.Cells(i, 12).Value > GreatestVolume Then
                        GreatestVolume = ws.Cells(i, 12).Value
                        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                        Else
                
                        GreatestVolume = GreatestVolume
                
                        End If
                
                            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
                            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
                            ws.Cells(4, 17).Value = GreatestVolume
            
            Next i
        
    Next ws
        
End Sub





