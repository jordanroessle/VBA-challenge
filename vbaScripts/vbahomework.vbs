Sub stockData()
    'intialize headers'
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"

    'making columns wider to fit text'
    Columns("j").ColumnWidth = "12"
    Columns("k").ColumnWidth = "13.5"
    Columns("l").ColumnWidth = "16"

    'formatting k column to percentage'
    Columns("k").NumberFormat = "0.00%"

    'using unique ticker counter to know where to output results'
    Dim uniqueTicker As Double: uniqueTicker = 0

    'loop until end of sheet'
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

        'add count to uniqueTicker'
        uniqueTicker = uniqueTicker + 1

        'variabes for each ticker, using doubles for larger numbers'
        Dim ticker As String: ticker = Cells(i, 1).Value
        Dim counter As Double: counter = 1
        Dim stockVol As Double: stockVol = Cells(i, 7).Value
        Dim yearlyChange As Double
        Dim percentChange As Double

        'loop until ticker changes'
        Do While (ticker = Cells(i + counter, 1).Value)
            'sum stock as we iterate'
            stockVol = stockVol + Cells(i + counter, 7).Value
            'iterate'
            counter = counter + 1
        Loop
        'set up so that i + counter = first row with different ticker'

        'calculate yearly change'
        yearlyChange = Cells(i + counter - 1, 6).Value - Cells(i, 3).Value

        'calculate percent change'
        If (Cells(i, 3) = 0) Then
            percentChange = 0
        Else
            percentChange = yearlyChange / Cells(i, 3).Value
        End If

        'output results'
        Cells(uniqueTicker + 1, 9).Value = ticker
        Cells(uniqueTicker + 1, 10).Value = yearlyChange
        Cells(uniqueTicker + 1, 11).Value = percentChange
        Cells(uniqueTicker + 1, 12).Value = stockVol

        'color for percent change, leaving no change white'
        If (yearlyChange > 0) Then
            Cells(uniqueTicker + 1, 10).Interior.ColorIndex = 4
        ElseIf (yearlyChange < 0) Then
            Cells(uniqueTicker + 1, 10).Interior.ColorIndex = 3
        End If
        'iterate by counter, -1 because next adds 1'
        i = i + counter - 1
        Next
End Sub
