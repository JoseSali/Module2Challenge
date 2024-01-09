Function StockAnalysis()
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim ticker As String
    Dim openprice As Double
    Dim closingprice As Double
    Dim yearchange As Double
    Dim percent As Double
    Dim volumesum As Double
    Dim sumrow As Double
    
    ' loop worksheets
    For Each ws In ThisWorkbook.Worksheets
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        sumrow = 2
        
        For i = 2 To lastrow
            If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
                openprice = ws.Cells(i, "C").Value
            End If
            
            volumesum = volumesum + ws.Cells(i, "G").Value
            
            If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
                closingprice = ws.Cells(i, "F").Value
                yearchange = closingprice - openprice
                percent = yearchange / openprice
                
                ws.Cells(sumrow, "I").Value = ws.Cells(i, "A").Value
                ws.Cells(sumrow, "J").Value = yearchange
                ws.Cells(sumrow, "K").Value = percent
                ws.Cells(sumrow, "L").Value = volumesum
                
                sumrow = sumrow + 1
                volumesum = 0
            End If
        Next i
        'the following was genereated by ai  
        ' Find Greatest % Increase, Greatest % Decrease, Greatest Total Volume
        Dim maxPercentIncrease As Double
        Dim minPercentDecrease As Double
        Dim maxVolume As Double
        Dim maxPercentIncreaseTicker As String
        Dim minPercentDecreaseTicker As String
        Dim maxVolumeTicker As String
        
        maxPercentIncrease = -1000000 ' Initialize with a very low value
        minPercentDecrease = 1000000 ' Initialize with a very high value
        maxVolume = 0 ' Initialize with zero
        
        lastrow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i, "K").Value > maxPercentIncrease Then
                maxPercentIncrease = ws.Cells(i, "K").Value
                maxPercentIncreaseTicker = ws.Cells(i, "I").Value
            End If
            
            If ws.Cells(i, "K").Value < minPercentDecrease Then
                minPercentDecrease = ws.Cells(i, "K").Value
                minPercentDecreaseTicker = ws.Cells(i, "I").Value
            End If
            
            If ws.Cells(i, "L").Value > maxVolume Then
                maxVolume = ws.Cells(i, "L").Value
                maxVolumeTicker = ws.Cells(i, "I").Value
            End If
        Next i
        
        ' Output results in column O
        ws.Cells(1, "O").Value = "Greatest % Increase"
        ws.Cells(2, "O").Value = maxPercentIncreaseTicker
        ws.Cells(3, "O").Value = maxPercentIncrease
        
        ws.Cells(5, "O").Value = "Greatest % Decrease"
        ws.Cells(6, "O").Value = minPercentDecreaseTicker
        ws.Cells(7, "O").Value = minPercentDecrease
        
        ws.Cells(9, "O").Value = "Greatest Total Volume"
        ws.Cells(10, "O").Value = maxVolumeTicker
        ws.Cells(11, "O").Value = maxVolume
    Next ws
End Function

