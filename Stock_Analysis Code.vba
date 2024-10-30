Sub stock_analysis()
    ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim percentChange As Double
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim start As Long
    Dim rowCount As Long
    Dim ticker As String
    Dim openprice, closeprice As Double
    
    
    'Iterate over worksheets
    For Each ws In Worksheets
    
    ' Set title row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' Set initial values
    total = 0
    change = 0
    start = 2
    openprice = ws.Cells(2, "C").Value

    
    ' get the row number of the last row with data
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    outputRow = 2
    
    'Initialize openprice for the first ticker
    openprice = ws.Cells(2, "C").Value
    
    For i = 2 To rowCount
    
        ' If ticker changes then print results, use Range!
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            total = total + ws.Cells(i, "G").Value
            ticker = ws.Cells(i, 1).Value
            closeprice = ws.Cells(i, "F").Value
            
            
            'When the ticker is different,put that total in column L
            ws.Cells(start, "I").Value = ticker
            ws.Cells(start, "J").Value = closeprice - openprice
            If openprice <> 0 Then
                ws.Cells(start, "K").Value = FormatPercent((closeprice - openprice) / openprice, 2)
            Else
                ws.Cells(start, "K").Value = Null
            End If
            
            ws.Cells(start, "L").Value = total
            
            ' reset variables for new stock ticker
            start = start + 1
            total = 0
            
            
            'Set the new open price for the next ticker
            openprice = ws.Cells(i + 1, "C").Value
        
        Else
            'Accumulate total stock volume for the current ticker
            total = total + ws.Cells(i, "G").Value
        
        'Before the conditional formatting check, calculate the change
            change = closeprice - openprice
        
        'Conditional Formatting
            If (change > 0) Then
                ws.Cells(start, "J").Interior.ColorIndex = 4 'Green for increase
            ElseIf (change < 0) Then
                ws.Cells(start, "J").Interior.ColorIndex = 3 'Red for decrease
            Else
                'Do Nothing (default White)
            End If
        
        End If
        Next i
    
    ' take the max and min and place them in a separate part in the worksheet
        Dim max_price As Double
        Dim min_price As Double
        Dim max_volume As Double
        Dim max_price_stock As String
        Dim min_price_stock As String
        Dim max_volume_stock As String
        Dim j As Integer
        
        'Inits
        max_price = ws.Cells(2, "K").Value
        min_price = ws.Cells(2, "K").Value
        max_volume = ws.Cells(2, "L").Value
        max_price_stock = ws.Cells(2, "I").Value
        min_price_stock = ws.Cells(2, "I").Value
        max_volume_stock = ws.Cells(2, "I").Value
        
        For j = 2 To row_count
            'Compare current row to the inits (first row)
            If (ws.Cells(j, "K").Value > max_price) Then
                'We have a new Max Percent Change!
                max_price = ws.Cells(j, "K").Value
                max_price_stock = ws.Cells(j, "I").Value
            End If
            
            If (ws.Cells(j, "K").Value < min_price) Then
                'We have a new Min Percent Change!
                min_price = ws.Cells(j, "K").Value
                min_price_stock = ws.Cells(j, "I").Value
            End If
            
            If (ws.Cells(j, "L").Value > max_volume) Then
                'We have a new Max Volume Change!
                max_volume = ws.Cells(j, "L").Value
                max_volume_stock = ws.Cells(j, "I").Value
            End If
        Next j
        
    'Write out to Excel Notebook
    ws.Range("P2").Value = max_price_stock
    ws.Range("P3").Value = min_price_stock
    ws.Range("P4").Value = max_volume_stock
    
    ws.Range("Q2").Value = max_price
    ws.Range("Q3").Value = min_price
    ws.Range("Q4").Value = max_volume
    
    Next ws
    
End Sub

