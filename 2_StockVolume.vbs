Sub StockVolume():
    Dim TickerName  As String
    
    'Name the variable for total volume of stock
    Dim TotalStockVolume As LongLong
    
    TotalStockVolume = 0
    
    'Keep track of the location for each stock
    Dim Results_row As Long
    Results_row = 2
    
    'Identify the lastrow of Stock Volumes
    lastrow = Cells(Rows.Count, 7).End(xlUp).row
    
    'Loop through all the stock volumes
    For i = 2 To lastrow
        
        'Write the conditional
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            
            'Add to the total stock volume
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
        Else
            'Set the TickerName
            TickerName = Cells(i, 1).Value
            
            'Add to the TotalStock Volume
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
            'Print the total stock volume
            Range("L" & Results_row).Value = TotalStockVolume
            
            'Print the ticker name
            Range("M" & Results_row).Value = TickerName
            
            'Add a row
            Results_row = Results_row + 1
            
            'Reset the Stock amount to 0
            TotalStockVolume = 0
            
        End If
    Next i
End Sub