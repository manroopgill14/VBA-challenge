# VBA-challenge
Module 2 Challenge

*Tutor helped with following code
    
    If ws.cells(row + 1, 1).Value <> ws.cells(row, 1).Value Then
    
        ws.cells(ARow, 9).Value = ws.cells(row, 1).Value
        
        ws.cells(ARow, 10).Value = ws.cells(row, 6).Value - opens
        
        ws.cells(ARow, 11).Value = ws.cells(ARow, 10).Value / opens
        
        ws.cells(ARow, 11).NumberFormat = "0.00%"
        
        ws.cells(ARow, 12).Value = Total
       
