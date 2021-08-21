Attribute VB_Name = "Module2"
Sub extras()
    For Each WS In Worksheets

       
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        maxRow = 0
        minRow = 0
        maxVolRow = 0
        maxTicker = ""
        Max = 0
        Min = WS.Cells(2, 11).Value
        
        Name = WS.Name


        For i = 2 To LastRow
            If WS.Cells(i, 11).Value > Max Then
            
                Max = WS.Cells(i, 11).Value
                maxRow = i
                maxTicker = WS.Range("I" & i).Value
                maxVol = WS.Range("L" & i).Value
            End If
            
            If WS.Cells(i, 11).Value < Min Then
            
                Min = WS.Cells(i, 11).Value
                minRow = i
                minTicker = WS.Range("I" & i).Value
                minVol = WS.Range("L" & i).Value
            End If
        Next i
            WS.Range("O2").Value = "Ticker"
            WS.Range("P2").Value = "Value %"
            WS.Range("Q2").Value = "Volume"
            
            
            WS.Range("N3").Value = "Greatest % increase: "
            WS.Range("O3").Value = maxTicker
            WS.Range("P3").Value = Max
            WS.Range("Q3").Value = maxVol
            
            
            WS.Range("N4").Value = "Greatest % decrease: "
            WS.Range("O4").Value = minTicker
            WS.Range("P4").Value = Min
            WS.Range("Q4").Value = minVol
            
    Next WS
End Sub

