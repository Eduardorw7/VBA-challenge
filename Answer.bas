Attribute VB_Name = "Module1"
Sub Multipleyearstockdata():
    For Each ws In Worksheets
    
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim ticketcount As Long
    Dim lastrowa As Long
    Dim lastrowi As Long
    Dim perchange As Double
    Dim greatincrease As Double
    Dim greatdecrease As Double
    Dim greatvolume As Double
    
    WorksheetName = ws.Name
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ticketcount = 2
    
    j = 2
    
    lastrowa = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastrowa
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(ticketcount, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(ticketcount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
                If ws.Cells(ticketcount, 10).Value < 0 Then
                ws.Cells(ticketcount, 10).Interior.ColorIndex = 3
                Else
                ws.Cells(ticketcount, 10).Interior.ColorIndex = 4
                
                End If
                
                If ws.Cells(j, 3).Value <> 0 Then
                perchange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                ws.Cells(ticketcount, 11).Value = Format(perchange, "Percent")
                
                Else
                
                ws.Cells(ticketcount, 11).Value = Format(0, "Percent")
                
                End If
                
            ws.Cells(ticketcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            ticketcount = ticketcount + 1
            
            j = i + 1
            
            End If
            
        Next i
        
        lastrowi = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        greatvolume = ws.Cells(2, 12).Value
        greatincrease = ws.Cells(2, 11).Value
        greatdecrease = ws.Cells(2, 11).Value
        
            For i = 2 To lastrowi
            
                If ws.Cells(i, 12).Value > greatvolume Then
                greatvolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatvolume = greatvolume
                
                End If
                
                If ws.Cells(i, 11).Value < greatincrease Then
                greatincrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatincerase = greatincrease
                
                End If
                
                If ws.Cells(i, 11).Value < greatdecrease Then
                greatdecrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatdecrease = greatdecrease
                
                End If
                
            ws.Cells(2, 17).Value = Format(greatincrease, "Percent")
            ws.Cells(3, 17).Value = Format(greatdecrease, "Percent")
            ws.Cells(4, 17).Value = Format(greatvolume, "Scientific")
            
            Next i
            
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
 
 
  
End Sub
