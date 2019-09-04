Attribute VB_Name = "Module1"
Sub LoopOnStocks()

    'Define Variables

    Dim lastRowA As Long
    ' Dim lastColumn As Long
    Dim tickerSymbol As String
    Dim yearOpenPrice As Double
    Dim yearClosePrice As Double
    Dim yearPriceDifference As Double
    Dim totalStockVolume As Double
    
    Dim rowHolder As Double
    Dim columnHolder As Double
    
    rowHolder = 2
    columnHolder = 9
    
    'Label Headers

    Range("I1") = "Ticker Symbol"
    Range("J1") = "Yearly Price Change"
    Range("K1") = "Yearly Price Percent Change"
    Range("L1") = "Total Stock Volume"

    lastRowA = Cells(Rows.Count, 1).End(xlUp).Row
    ' lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
    
    For i = 2 To lastRowA
        
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            tickerSymbol = Cells(i, 1).Value
            yearOpenPrice = Cells(i, 3).Value
            totalStockVolume = Cells(i, 7).Value
            
        ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
            
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
            yearClosePrice = Cells(i, 6).Value
            yearPriceDifference = yearClosePrice - yearOpenPrice
            
            Cells(rowHolder, columnHolder).Value = tickerSymbol
            Cells(rowHolder, columnHolder + 1).Value = yearPriceDifference
            'Cells(rowHolder, columnHolder + 1).NumberFormat = "###.########"
            If yearOpenPrice <> 0 Then
                Cells(rowHolder, columnHolder + 2).Value = yearPriceDifference / yearOpenPrice
            Else
                Cells(rowHolder, columnHolder + 2).Value = "At the beginning of the year, the open price was 0."
            End If
            Cells(rowHolder, columnHolder + 2).NumberFormat = "0.000000%"
            Cells(rowHolder, columnHolder + 3).Value = totalStockVolume
            
            If yearPriceDifference > 0 Then
        
                Cells(rowHolder, columnHolder + 1).Interior.ColorIndex = 4
                Cells(rowHolder, columnHolder + 2).Interior.ColorIndex = 4
         
            ElseIf yearPriceDifference < 0 Then
        
                Cells(rowHolder, columnHolder + 1).Interior.ColorIndex = 3
                Cells(rowHolder, columnHolder + 2).Interior.ColorIndex = 3
            
            End If
            
            rowHolder = rowHolder + 1
            
        End If
        
    Next i
        
    'Hard Solution
    'Create a Loop that iterates through the summarized data for each stock, and find the stock with the greatest % increase, the greatest % decrease, and the greatest total stock volumen. Also,
    'save the ticker symbols as well for each.

    'Label headers
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker Symbol"
    Range("P1") = "Value"

    'Define variables for hard solution
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalStockVolume As Double
    
    Dim GPItickerSymbol As String
    Dim GPDtickerSymbol As String
    Dim GTSVtickerSymbol As String

    Dim lastRowB As Double
    
    lastRowB = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i1 = 2 To lastRowB
    
        If Cells(i, 11).Row <> lastRowB Then
    
            If Cells(i, 11).Value < Cells(i + 1, 11).Value Then
        
                greatestPercentIncrease = Cells(i + 1, 11).Value
                GPItickerSymbol = Cells(i + 1, 9).Value
             
            ElseIf Cells(i, 11).Value > Cells(i + 1, 11).Value Then
        
                greatestPercentDecrease = Cells(i + 1, 11).Value
                GPDtickerSymbol = Cells(i + 1, 9).Value
        
            End If
            
            If Cells(i, 12).Value < Cells(i + 1, 12).Value Then
    
                greatestTotalStockVolume = Cells(i + 1, 12).Value
                GTSVtickerSymbol = Cells(i + 1, 9).Value
                
            End If
                
        End If
        
    Next i1
    
    'Populate Chart
    
    Range("O2") = GPItickerSymbol
    Range("O3") = GPDtickerSymbol
    Range("O4") = GTSVtickerSymbol
    
    Range("P2") = greatestPercentIncrease
    Range("O3") = greatestPercentDeccrease
    Range("O4") = greatestTotalStockVolume

End Sub
