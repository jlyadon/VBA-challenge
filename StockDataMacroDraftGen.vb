Dim lastrow as double
Dim n as integer 'n will be used to count ticker symbols
Dim totalvolume as double
Dim GInc as double
Dim GDec as double
Dim GTot as double
Dim q as integer 'an index to track which quarter...
Dim d as integer '...so that we can assign d the value of trading days in the quarter

q = 1

For each ws in Worksheets

'Make a column with a list of the ticker symbols:
    lastrow = ws.Cells(rows.count,1).End(xlUp).row
    n = 0

    ws.Cells(1,9).value = "Ticker"

    for i = 2 to lastrow
        if ws.cells(i,1).value <> ws.cells(i-1,1).value then
            ws.cells(n+2,9).value = ws.cells(i,1).value
            n = n + 1
        end if
    next i
    'n is now stored as the number of distinct ticker symbols in the sheet.

'Assign d the value of the number of trading days in the quarter:
if q = 1 then
    d = 62
elseif q = 2 then
    d = 63
elseif q = 3 then
    d = 64
elseif q = 4 then
    d = 64
end if
    

'Calculate the quarterly changes:

    ws.Cells(1, 10).value = "Quarterly Change"

    For i = 0 to n - 1
        'Subtract the closing value on the last day from the opening value on the first day:
        ws.Cells(i+2,10).value = ws.Cells(d + 1 + d * i, 6).Value - ws.Cells(2 + d * i, 3).Value
        'Assign the appropriate color:
        If ws.Cells(i+2,10).value > 0 then
            ws.Cells(i+2,10).interior.colorindex = 4
        elseif ws.Cells(i+2,10).value < 0 then
            ws.Cells(i+2,10).interior.colorindex = 3
        End if
    next i

'Calcluate the percentage quarterly change:

    ws.Cells(1,11).value = "Percent Change"

    For i = 0 to n - 1
        ws.Cells(i+2,11).value = FormatPercent((ws.Cells(i+2,10).value / ws.Cells(2 + d * i, 3).Value))
    Next i

'Calculate the total stock volume for each stock:

    ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 0 to n - 1
        totalvolume = 0
        'The pattern check script confirmed that there are 62 trading days in each quarter.
        for j = 1 to 62
            totalvolume = totalvolume + ws.Cells(2 + d*i + j, 7)
        Next j
        ws.Cells(i + 2, 12).value = totalvolume
    Next i

'Make the chart for the greatest increase, decrease, and volume:

    'Label the chart:
    ws.cells(2,15).value = "Greatest % Increase"
    ws.cells(3,15).value = "Greatest % Decrease"
    ws.cells(4,15).value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    GInc = 0
    GDec = 0
    GTot = 0

    'Find the greatest percent increase:
    for i = 2 to n + 1
        if ws.Cells(i,11).value > GInc then
            GInc = ws.Cells(i,11).value
        End if
    Next i

    'Find the greatest percent decrease:
    for i = 2 to n + 1
        if ws.Cells(i,11).value < GDec then
            GDec = ws.Cells(i,11).value
        End if
    Next i

    'Find the greatest total stock volume:
    for i = 2 to n + 1
        if ws.Cells(i,12).value > GTot then
            GTot = ws.Cells(i,12).value
        End if
    Next i

    ws.cells(2,17).value = FormatPercent(GInc)
    ws.cells(3,17).value = FormatPercent(GDec)
    ws.cells(4,17).value = GTot

    'Find the ticker symbol for the greatest increase:
    For i = 2 to n + 1  
        if ws.Cells(i,11).value = GInc then
            ws.Cells(2,16).value = ws.cells(i, 9).value
        End if
    Next i

    'Find the ticker symbol for the greatest decrease:
    For i = 2 to n + 1 
        if ws.Cells(i,11).value = GDec then
            ws.Cells(3,16).value = ws.cells(i, 9).value
        End if
    Next i

    'Find the ticker symbol for the greatest total stock volume:
    For i = 2 to n + 1   
        if ws.Cells(i,12).value = GTot then
            ws.Cells(4,16).value = ws.cells(i, 9).value
        End if
    Next i
q = q + 1 'next quarter, for assigning the value of d
Next ws