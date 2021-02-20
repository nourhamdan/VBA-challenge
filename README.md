# VBA-challenge
homework





Sub Stonks()


Dim ticker As String
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStockVolume As Double
Dim openPrice
Dim closePrice As Double
Dim summaryTableRow As Double
Dim yearStart
Dim wsCount As Integer



'worksheet iterate
wsCount = ActiveWorkbook.Worksheets.Count

For ws = 1 To wsCount
    ' adding colomns for worksheet
    Worksheets(ws).Range("i1") = "Ticker"
    Worksheets(ws).Range("J1") = "Yearly Change"
    Worksheets(ws).Range("k1") = "Percent Change"
    Worksheets(ws).Range("L1") = "Total Stock Volume"
    
  

    summaryTableRow = 2
    
    For i = 2 To Worksheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
        ' set ticker & begin stock volume increment
        ticker = Worksheets(ws).Cells(i, 1)
        totalStockVolume = totalStockVolume + Cells(i, 7)
        
        'set opening price
        If openPrice = "" Then
            openPrice = Worksheets(ws).Cells(i, 3)
        End If
        
        'iterate
        If ticker <> Worksheets(ws).Cells((i + 1), 1) Then
        
        'set close price
        closePrice = Worksheets(ws).Cells(i, 6)
        'calc yearly change
        yearlyChange = openPrice - closePrice
        
        ' ticker output to worksheet
        Worksheets(ws).Range("i" & summaryTableRow).Value = ticker
        
        ' yearly change output to worksheet with formatting
        Worksheets(ws).Range("J" & summaryTableRow).Value = yearlyChange
        If yearlyChange > 0 Then
        Worksheets(ws).Range("J" & summaryTableRow).Interior.ColorIndex = 4
        Else
        Worksheets(ws).Range("J" & summaryTableRow).Interior.ColorIndex = 3
        End If
        
        ' percent changed 
        If startPrice <> closePrice Then
            percentChange = yearlyChange / closePrice
        Else
            percentChange = 0
        End If
        Worksheets(ws).Range("k" & summaryTableRow).Value = percentChange
        Worksheets(ws).Range("k" & summaryTableRow).NumberFormat = "0.00%"
        
        ' stock volume output to worksheet
        Worksheets(ws).Range("l" & summaryTableRow).Value = totalStockVolume
        
        
     
        
        ' reset variables and increment
        summaryTableRow = summaryTableRow + 1
        totalStockVolume = 0
        
        End If
        
    Next i

Next ws

End Sub
