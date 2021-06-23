Sub calculateStockInfo()
    'Define all variables
    endData = Cells(Rows.Count, 1).End(xlUp).Row
    Dim yChange As Double
    Dim pChange As Double
    Dim ticker As String
    Dim printStock As Integer
    Dim totalVolume As Double
    Dim stockClose As Double
    Dim stockStart As Double
    
    'Initialize Variables
    yChange = 0
    pChange = 0
    printStock = 2
    totalVolume = 0
    stockStart = Cells(2, 3)
    
    'Loop through all of the rows of data
    For i = 2 To endData
        
        'Increment totalVolume as we go through the rows
        totalVolume = Cells(i, 7).Value + totalVolume
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            stockClose = Cells(i, 6).Value
            
            yChange = stockClose - stockStart
            If stockStart <> 0 Then
                pChange = yChange / stockStart
            Else
                pChange = 0
            End If
            
            Cells(printStock, 8).Value = ticker
            Cells(printStock, 9).Value = yChange
            Cells(printStock, 10).Value = pChange
            Cells(printStock, 11).Value = totalVolume
            
            If Cells(printStock, 9).Value < 0 Then
                Cells(printStock, 9).Interior.ColorIndex = 3
            Else
                Cells(printStock, 9).Interior.ColorIndex = 4
            End If
            
            totalVolume = 0
            stockStart = Cells(i + 1, 3)
            printStock = printStock + 1
        
        End If
    Next i
End Sub
