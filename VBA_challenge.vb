Sub TickerSummary()

Dim ws As Worksheet

'loop through sheets
For Each ws In Sheets
    dim FirstPrice as double
    dim LastPrice as double
    dim RowCounter as integer
    dim TotChange as double
    dim PctChange as double

    'set initial price for first stock, intialize volume, find last filled row of sheet
    'and start row counter to print totals
    FirstPrice=ws.Cells(2,3).value
    volume=0
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    RowCounter=2
    
    'loop through each row adding volume
    for i = 2 to last_row
        volume=volume+ws.Cells(i,7).value

        'check for ticker change
        'if ticker changes, pull last price and calculate total and percent change
        'print final values into summary columns, reset volume, and iterate print totals RowCounter
        if ws.Cells(i+1,1).value <> ws.Cells(i,1).value Then
            LastPrice=ws.Cells(i,6).value
            TotChange=LastPrice-FirstPrice
            if FirstPrice<>0 Then
                PctChange=TotChange/FirstPrice
            elseif FirstPrice=0 and LastPrice>0 Then
                PctChange=999999999
            elseif FirstPrice=0 and LastPrice<0 Then
                PctChange=-999999999
            else
                PctChange=0
            end if
            ws.Cells(RowCounter,9).Value=ws.Cells(i,1).value
            ws.Cells(RowCounter,10).Value=TotChange
            ws.Cells(RowCounter,11).Value=PctChange
            ws.Cells(RowCounter,12).Value=volume
            FirstPrice=ws.Cells(i+1,3).Value
            volume=0
            RowCounter=RowCounter+1
        End if
    Next i

    'set summary headers
    ws.Cells(1,9).Value="Ticker"
    ws.Cells(1,10).Value="Yearly Change"
    ws.Cells(1,11).Value="Percent Change"
    ws.Cells(1,12).Value="Total Volume"

    'format percentage change cells
    for s = 2 to RowCounter-1
        ws.Cells(s,11).NumberFormat="0.00%"
    next s

    'conditional formatting to show direction of change
    for p = 2 to RowCounter-1
        if ws.Cells(p,10).value > 0 Then
            ws.Cells(p,10).Interior.ColorIndex=4
        elseif ws.Cells(p,10).Value < 0 Then
            ws.Cells(p,10).Interior.ColorIndex=3
        else
            ws.Cells(p,10).Interior.ColorIndex=6
        end if
    next p      
Next ws


End Sub