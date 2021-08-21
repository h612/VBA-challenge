Attribute VB_Name = "Module1"
Sub homework()


    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each WS In Worksheets

        
        'Re-running creates several columns. To avoid that, delete te columns that are created in the last run
        WS.Columns(9).EntireColumn.Delete
        WS.Columns(9).EntireColumn.Delete
        WS.Columns(9).EntireColumn.Delete
        WS.Columns(9).EntireColumn.Delete
        
        Dim WorksheetName As String
        Dim change As Double
        Dim symbNum As Long
        Dim isFirst As Double
        Dim sum As Double
        
        
        
        
        
  
        Total = 0
        change = 0
        symbNum = 2
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        LastColumn = 2 + WS.Cells(1, Columns.Count).End(xlToLeft).Column
        
        WorksheetName = WS.Name
     
        

        ' Add the Header
        WS.Cells(1, LastColumn).Value = "Ticker"
        WS.Cells(1, LastColumn + 1).Value = "Yearly Change"
        WS.Cells(1, LastColumn + 2).Value = "Percentage Change"
        WS.Cells(1, LastColumn + 3).Value = "TOTAL STOCK VOLUME"
        
    
       tempChange = 0
       isFirst = 0
       sum = 0
       
    
        For i = 2 To LastRow ' traverse all rows
                If isFirst = 0 Then
                
                        isFirst = Cells(i, 3).Value 'take first value of each ticker brand
                        End If
                        
               If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then 'check if the new brand starts
                  
                     
                     last = WS.Cells(i, 6).Value 'last value of current ticker
                     WS.Cells(symbNum, LastColumn).Value = WS.Cells(i, 1).Value 'store ticker name in row
                     WS.Cells(symbNum, LastColumn + 1).Value = last - isFirst 'yearly change
                     If isFirst = 0 Then
                        WS.Cells(symbNum, LastColumn + 2).Value = last 'avoid divide by 0
                     Else
                        WS.Cells(symbNum, LastColumn + 2).Value = ((last - isFirst) * 100) / isFirst
                     End If
                     
                     WS.Cells(symbNum, LastColumn + 3).Value = sum + WS.Cells(i, 7).Value 'all the volumes
                     
                     symbNum = symbNum + 1
                     isFirst = 0
                     sum = 0
                     
                     
               Else
                      
                      sum = sum + WS.Cells(i, 7).Value
                      
                      
               End If
               
    
    
         Next i
    
    
        Next WS
    
End Sub



