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

    
 '------------------------------------------------------------------------------------------------------------
    For I = 2 To lastrow
    
    close_price = ws.Cells(I, "F").Value
    open_price = ws.Cells(I, "C").Value
    volume = ws.Cells(I, "G").Value
    
    'yearly change, percent change, totalvolume calculations
    
    yearlychange = yearlychange + close_price - open_price
    totalstockvol = totalstockvol + volume
    percentchange = percentchange + (close_price - open_price) / open_price
       
        If ws.Cells(I, "A") <> ws.Cells(I + 1, "A") Then
            ws.Cells(outindex, "I") = ws.Cells(I, "A") ' ticker symbol
            ws.Cells(outindex, "J") = yearlychange 'yearly change
            ws.Cells(outindex, "K") = FormatPercent(percentchange, 2) 'percent change with % formatting
            ws.Cells(outindex, "L") = totalstockvol 'total stock volume
            
            outindex = outindex + 1
            yearlychange = 0
            percentchange = 0
            totalstockvol = 0
        End If
    Next I
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

For I = 1 To WS_count

        ' counts the rows in summary report
        lastrowsum = ActiveWorkbook.Worksheets(I).Cells(Rows.Count, "I").End(xlUp).Row
        
    For j = 2 To lastrowsum
    
        '------formatting conditions here
        If ActiveWorkbook.Worksheets(I).Cells(j, "J").Value >= 0 Then
            ActiveWorkbook.Worksheets(I).Cells(j, "J").Interior.ColorIndex = 4
        Else
            ActiveWorkbook.Worksheets(I).Cells(j, "J").Interior.ColorIndex = 3
        End If
        
        '-----sorts the greatest total stock volume
        topvol = Application.WorksheetFunction.Max(ActiveWorkbook.Worksheets(I).Range("L:L"))
        ActiveWorkbook.Worksheets(I).[q4] = topvol
        
        '------sorts the greatest decrease
        greatdecrease = Application.WorksheetFunction.Min(ActiveWorkbook.Worksheets(I).Range("K:K"))
        
        ActiveWorkbook.Worksheets(I).[q3] = greatdecrease
        ActiveWorkbook.Worksheets(I).[q3].NumberFormat = "0.00%"
        
        '----- sorts the great increase
        greatincrease = Application.WorksheetFunction.Max(ActiveWorkbook.Worksheets(I).Range("K:K"))
        ActiveWorkbook.Worksheets(I).[q2].NumberFormat = "0.00%"
        ActiveWorkbook.Worksheets(I).[q2] = greatincrease
        
        '------Find ticker for corresponding value
        If ActiveWorkbook.Worksheets(I).Cells(j, "L").Value = topvol Then
            ActiveWorkbook.Worksheets(I).Cells(4, "P").Value = ActiveWorkbook.Worksheets(I).Cells(j, "I").Value
        ElseIf ActiveWorkbook.Worksheets(I).Cells(j, "K").Value = greatdecrease Then
            ActiveWorkbook.Worksheets(I).Cells(3, "P").Value = ActiveWorkbook.Worksheets(I).Cells(j, "I").Value
        ElseIf ActiveWorkbook.Worksheets(I).Cells(j, "K").Value = greatincrease Then
            ActiveWorkbook.Worksheets(I).Cells(2, "P").Value = ActiveWorkbook.Worksheets(I).Cells(j, "I").Value
        End If
        
    Next j
Next I

End Sub