Sub TickerSummary()

Dim ws As Worksheet

For Each ws In Sheets
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
Next ws


End Sub