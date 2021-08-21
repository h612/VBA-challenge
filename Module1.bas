Attribute VB_Name = "Module1"
Sub homework()


    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each WS In Worksheets

        
        
        WS.Columns(9).EntireColumn.Delete
        WS.Columns(9).EntireColumn.Delete
        WS.Columns(9).EntireColumn.Delete
        WS.Columns(9).EntireColumn.Delete
        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String
        Dim change As Double
        Dim symbNum As Long
        Dim isFirst As Double
        Dim sum As Double
        
        
        
        
        
       ' Dim total As Integer
        Total = 0
        change = 0
        symbNum = 2
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        ' Determine the Last Column Number
        LastColumn = 2 + WS.Cells(1, Columns.Count).End(xlToLeft).Column
                ' Grabbed the WorksheetName
        WorksheetName = WS.Name
        ' MsgBox WorksheetName
        

        ' Add the Header
        WS.Cells(1, LastColumn).Value = "Ticker"
        WS.Cells(1, LastColumn + 1).Value = "Yearly Change"
        WS.Cells(1, LastColumn + 2).Value = "Percentage Change"
        WS.Cells(1, LastColumn + 3).Value = "TOTAL STOCK VOLUME"
        
        'Range(LastColumn&"1:"&(LastColumn-4)&LastRow).Clear
       tempChange = 0
       isFirst = 0
       sum = 0
       'open=0
    
        For i = 2 To LastRow
                If isFirst = 0 Then
                
                        isFirst = Cells(i, 3).Value
                        End If
                        
               If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                  
                     
                     last = WS.Cells(i, 6).Value
                     WS.Cells(symbNum, LastColumn).Value = WS.Cells(i, 1).Value
                     WS.Cells(symbNum, LastColumn + 1).Value = last - isFirst
                     If isFirst = 0 Then
                        WS.Cells(symbNum, LastColumn + 2).Value = last
                     Else
                        WS.Cells(symbNum, LastColumn + 2).Value = ((last - isFirst) * 100) / isFirst
                     End If
                     
                     WS.Cells(symbNum, LastColumn + 3).Value = sum + WS.Cells(i, 7).Value
                     
                     symbNum = symbNum + 1
                     isFirst = 0
                     sum = 0
                     
                     
               Else
                      'change = change + Range("D" & i).Value - Range("F" & i).Value
                      sum = sum + WS.Cells(i, 7).Value
                      
                      
               End If
               
    
    
         Next i
    
    
        Next WS
    
End Sub



