Attribute VB_Name = "Module3"
Sub color()
For Each WS In Worksheets

       
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
         If WS.Range("J" & i).Value > 0 Then
            WS.Cells(i, 10).Interior.ColorIndex = 4
        
         Else
            WS.Cells(i, 10).Interior.ColorIndex = 3
         End If
        
        Next i
        
        
        
        Next WS
        
End Sub

