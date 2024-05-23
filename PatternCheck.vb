'This subroutine determines whether the pattern of dates repeats throughout column B

Dim i As Double
Dim j As Double 'i and j are used as indices
Dim n As Integer 'n will be used to count ticker symbols
Dim q as integer 'will track which quarter/sheet we're looping over
q = 1
Dim d as integer 'the number of days in the quarter
Dim lastrow As Double

Dim patternholds As Boolean

For each ws in Worksheets
    n = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow 'This loop counts ticker symbols
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            n = n + 1
        End If
    Next i

    patternholds = True

    'Assign d the number of trading days in the quarter:
    if q = 1 then
        d = 62
    else if q = 2 then
        d = 63
    else if q = 3 then
        d = 64
    else if q = 4 then
        d = 64
    end if

    'If, at any point, the pattern breaks, this loop will switch patternholds to "false"
    For i = 1 To d 'looping over d days of trading
        For j = 0 To n - 1 'looping over n ticker symbols
            If ws.Cells(i + 1 + d * j, 2).Value <> ws.Cells(1 + i, 2).Value Then
                patternholds = False
            End If
        Next j
    Next i

    If patternholds = True Then
    MsgBox ("The pattern holds for " & ws.Name & ". All stocks are traded on all days.")
    Else: MsgBox ("The pattern does not hold for all stocks in " & ws.Name & ".")
    End If
    q = q + 1 'next quarter/sheet
Next ws