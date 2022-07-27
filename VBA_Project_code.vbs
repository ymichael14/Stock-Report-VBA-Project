Sub stocks()

Dim ws As Worksheet
For Each ws In Sheets

    outindex = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    yearlychange = 0
    totalstockvol = 0
    percentchange = 0
    greatestincrease = 0
    greatdecrease = 0
    topvol = 0
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    open_price = 0
    
 '------------------------------------------------------------------------------------------------------------
    For i = 2 To lastrow
    
    ticker_symbol = ws.Cells(i, "A")
    close_price = ws.Cells(i, "F").Value
    volume = ws.Cells(i, "G").Value
    
    'yearly change, percent change, totalvolume calculations
    
    If open_price = 0 Then
        open_price = ws.Cells(i, "C").Value
    End If
    
    totalstockvol = totalstockvol + volume
       
        If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
            yearlychange = ws.Cells(i, "F").Value - open_price
            percentchange = yearlychange / open_price
            ws.Cells(outindex, "I") = ws.Cells(i, "A") ' ticker symbol
            ws.Cells(outindex, "J") = yearlychange 'yearly change
            ws.Cells(outindex, "K") = FormatPercent(percentchange, 2) 'percent change with % formatting
            ws.Cells(outindex, "L") = totalstockvol 'total stock volume
            
            outindex = outindex + 1
            open_price = 0
            percentchange = 0
            totalstockvol = 0
        End If
    Next i
'-------------------------------------------------------------------------------------------------------------
    ws.[P1] = "Ticker"
    ws.[Q1] = "Value"
    ws.[O2] = "Greatest % Increase"
    ws.[O3] = "Greatest % Decrease"
    ws.[O4] = "Great Total Volume"

ws.Range("I1:R1").EntireColumn.AutoFit

Next ws

Dim WS_count As Integer
WS_count = ActiveWorkbook.Worksheets.Count

For i = 1 To WS_count

        ' counts the rows in summary report
        lastrowsum = ActiveWorkbook.Worksheets(i).Cells(Rows.Count, "I").End(xlUp).Row
        
    For j = 2 To lastrowsum
    
        '------formatting conditions here
        If ActiveWorkbook.Worksheets(i).Cells(j, "J").Value >= 0 Then
            ActiveWorkbook.Worksheets(i).Cells(j, "J").Interior.ColorIndex = 4
        Else
            ActiveWorkbook.Worksheets(i).Cells(j, "J").Interior.ColorIndex = 3
        End If
        
        '-----sorts the greatest total stock volume
        topvol = Application.WorksheetFunction.Max(ActiveWorkbook.Worksheets(i).Range("L:L"))
        ActiveWorkbook.Worksheets(i).[q4] = topvol
        
        '------sorts the greatest decrease
        greatdecrease = Application.WorksheetFunction.Min(ActiveWorkbook.Worksheets(i).Range("K:K"))
        
        ActiveWorkbook.Worksheets(i).[q3] = greatdecrease
        ActiveWorkbook.Worksheets(i).[q3].NumberFormat = "0.00%"
        
        '----- sorts the great increase
        greatincrease = Application.WorksheetFunction.Max(ActiveWorkbook.Worksheets(i).Range("K:K"))
        ActiveWorkbook.Worksheets(i).[q2].NumberFormat = "0.00%"
        ActiveWorkbook.Worksheets(i).[q2] = greatincrease
        
        '------Find ticker for corresponding value
        If ActiveWorkbook.Worksheets(i).Cells(j, "L").Value = topvol Then
            ActiveWorkbook.Worksheets(i).Cells(4, "P").Value = ActiveWorkbook.Worksheets(i).Cells(j, "I").Value
        ElseIf ActiveWorkbook.Worksheets(i).Cells(j, "K").Value = greatdecrease Then
            ActiveWorkbook.Worksheets(i).Cells(3, "P").Value = ActiveWorkbook.Worksheets(i).Cells(j, "I").Value
        ElseIf ActiveWorkbook.Worksheets(i).Cells(j, "K").Value = greatincrease Then
            ActiveWorkbook.Worksheets(i).Cells(2, "P").Value = ActiveWorkbook.Worksheets(i).Cells(j, "I").Value
        End If
        
    Next j
Next i


End Sub
