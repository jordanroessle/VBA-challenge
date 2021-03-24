Sub allWorksheets()
    'Loop through each ws'
    for each ws in Worksheets

        'intialize headers'
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"

        'Bonus: Setup new table'
        ws.Range("o2").Value = "Great % Increase"
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"

        'making columns wider to fit text'
        ws.Columns("j").ColumnWidth = "12"
        ws.Columns("k").ColumnWidth = "13.5"
        ws.Columns("l").ColumnWidth = "16"

        'Bonus: making o column wider'
        ws.Columns("o").ColumnWidth = "19.5"

        'formatting k column to percentage'
        ws.Columns("k").NumberFormat = "0.00%"

        'Bonus: formatting q2 and q3 to percentage'
        ws.Range("q2").NumberFormat = "0.00%"
        ws.Range("q3").NumberFormat = "0.00%"

        'using unique ticker counter to know where to output results'
        Dim uniqueTicker As Double: uniqueTicker = 0

        'Bonus: New Variables to store values'
        Dim maxTicker As String
        Dim minTicker As String
        Dim volTicker As String
        Dim maxPercentage As Double: maxPercentage = 0
        Dim minPercentage As Double: minPercentage = 0
        Dim maxVol As Double: maxVol = 0

        'loop until end of sheet'
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            'add count to uniqueTicker'
            uniqueTicker = uniqueTicker + 1

            'variabes for each ticker, using doubles for larger numbers'
            Dim ticker As String: ticker = ws.Cells(i, 1).Value
            Dim counter As Double: counter = 0
            Dim stockVol As Double: stockVol = ws.Cells(i, 7).Value
            Dim yearlyChange As Double
            Dim percentChange As Double

            'loop until ticker changes'
            Do While (ticker = ws.Cells(i + counter + 1, 1).Value)
                'sum stock as we iterate'
                stockVol = stockVol + ws.Cells(i + counter + 1, 7).Value
                'iterate'
                counter = counter + 1
            Loop
            'set up so that i + counter = last row with same ticker'

            'calculate yearly change'
            yearlyChange = ws.Cells(i + counter, 6).Value - ws.Cells(i, 3).Value

            'calculate percent change'
            If (ws.Cells(i, 3) = 0) Then
                percentChange = 0
            Else
                percentChange = yearlyChange / ws.Cells(i, 3).Value
                'Bonus: Compare to max and min'
                If (percentChange > maxPercentage) Then
                    maxPercentage = percentChange
                    maxTicker = ticker
                ElseIf (percentChange < minPercentage) Then
                    minPercentage = percentChange
                    minTicker = ticker
                End If
            End If

            'Bonus: Compare stockVol to maxVol'
            If (stockVol > maxVol) Then
                maxVol = stockVol
                volTicker = ticker
            End If

            'output results'
            ws.Cells(uniqueTicker + 1, 9).Value = ticker
            ws.Cells(uniqueTicker + 1, 10).Value = yearlyChange
            ws.Cells(uniqueTicker + 1, 11).Value = percentChange
            ws.Cells(uniqueTicker + 1, 12).Value = stockVol

            'color for percent change, leaving no change as white'
            If (yearlyChange > 0) Then
                ws.Cells(uniqueTicker + 1, 10).Interior.ColorIndex = 4
            ElseIf (yearlyChange < 0) Then
                ws.Cells(uniqueTicker + 1, 10).Interior.ColorIndex = 3
            End If
            'iterate by counter'
            i = i + counter
        Next
        'Bonus: Output results'
        ws.Range("p2").Value = maxTicker
        ws.Range("p3").Value = minTicker
        ws.Range("p4").Value = volTicker
        ws.Range("q2").Value = maxPercentage
        ws.Range("q3").Value = minPercentage
        ws.Range("q4").Value = maxVol
    Next ws
End Sub

