Sub Ticker():
    
    'Write code for #1 of assignment: Ticker symbol
    'Set a variable for the <ticker> column
    Dim OriginalTicker As Long
    OriginalTicker = 1
    
    'Common commands for all 3 points of assignment:
    
    'Name the columns for placing the outputs
    Range("H1").Value = "New_Date"
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly_Change"
    Range("K1").Value = "Percent_Change"
    Range("L1").Value = "Total_Stock_Volume"
    
    'Count the number of rows and retreive last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).row
    
    'Create rows to be added one below the other
    Dim TickerList_Row As Long
    TickerList_Row = 2
    
    'Set the TickerName
    For i = 2 To lastrow
        Dim TickerName As String
        TickerName = Cells(i, 1).Value
        
        'Loop through rows in the original ticker column: Write a for loop to meet condition that if next cell is not the same as previous cell, then print the value of previous cell in the ticker symbol list column
        
        If Cells(i + 1, OriginalTicker).Value <> Cells(i, OriginalTicker).Value Then
            
            'MsgBox ("TickerName" & TickerName & TickerList_Row)
            Range("I" & TickerList_Row).Value = TickerName
            
            'Add one to theTickerList_Row
            TickerList_Row = TickerList_Row + 1
            
        End If
    Next i
End Sub