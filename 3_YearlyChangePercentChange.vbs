Sub PercentYearlyChange():
    
    'Define variables
    Dim TickerName  As String
    Dim OpeningBalance As Variant
    Dim ClosingBalance As Variant
    Dim PercentChange As Variant
    Dim YearlyChange As Variant
    
    'Count the number of rows of the results Ticker column and retreive last row
    Lastrow_TickerResults = Cells(Rows.Count, 9).End(xlUp).row
    
    'Count the number of rows of the 1st Ticker column
    Lastrow_Ticker = Cells(Rows.Count, 1).End(xlUp).row
    
    'Begin the outer loop. Here it will loop through the list of unique TickerNames generated in the first module.
    For i = 2 To Lastrow_TickerResults
        TickerName = Cells(i, 9).Value
        'Begin the inner loop. Here it will retreive the values of opening balance and closing balance. 
        For j = 2 To Lastrow_Ticker
            
            If Cells(j, 1).Value = TickerName Then
                
                If (Cells(j, 2).Value = "20200102") Then
                    OpeningBalance = Cells(j, 3).Value
                End If
                
                If (Cells(j, 2).Value = "20201231") Then
                    ClosingBalance = Cells(j, 6).Value
                    
                    'Calculate Yearly Change and Percent Change
                    YearlyChange = ClosingBalance - OpeningBalance
                    PercentChange = (YearlyChange / OpeningBalance) * 100
                    
                End If
            End If
        Next j

        'Print the values
        Cells(i, 10).Value = YearlyChange
        Cells(i, 11).Value = PercentChange
        'Add a row
        Results_row = Results_row + 1
        'Reset the amounts to 0
        OpeningBalance = 0
        ClosingBalance = 0
        PercentChange = 0
        YearlyChange = 0        
    Next i
End Sub
