Sub PercentYearlyChange():

'Define variables
Dim TickerName As String
Dim OpeningBalance As Variant
Dim ClosingBalance As Variant
Dim PercentChange As Variant
Dim YearlyChange As Variant


'Count the number of rows of the results Ticker column and retreive last row
Lastrow_TickerResults = Cells(Rows.Count, 9).End(xlUp).row
MsgBox ("Lastrow_TickerResults" & " " & Lastrow_TickerResults)
    
'Count the number of rows of the 1st Ticker column
Lastrow_Ticker = Cells(Rows.Count, 1).End(xlUp).row
MsgBox ("Lastrow_Ticker" & " " & Lastrow_Ticker)

'Begin the 1st loop
'For i = 2 To Lastrow_TickerResults
For i = 2 To 3

TickerName = Cells(i, 9).Value
          
          'For j = 2 To Lastrow_Ticker
            
            For j = 2 To 600
        
        'MsgBox (Cells(j, 1).Value & " " & TickerName)
        
            If Cells(j, 1).Value = TickerName Then
            
                    MsgBox (Cells(j, 2).Value)
                
                    If (Cells(j, 2).Value = "20200102") Then
                        OpeningBalance = Cells(j, 3).Value
                        MsgBox (TickerName & " works " & OpeningBalance)
                    
                    
                    If (Cells(j, 2).Value = "20201231") Then
                        ClosingBalance = Cells(j, 6).Value
                        MsgBox (TickerName & " closing " & ClosingBalance)
                        
                        'Calculate Yearly Change and Percent Change
                        YearlyChange = ClosingBalance - OpeningBalance
                        PercentChange = (YearlyChange / OpeningBalance) * 100
                        MsgBox (" YearlyChange " & YearlyChange & " PercentChange " & PercentChange)
                    
                    End If
                    End If
                    End If
                    Next j
                    
                    'Add a row
                    Results_row = Results_row + 1
                    'Reset the amounts to 0
                    OpeningBalance = 0
                    ClosingBalance = 0
                    PercentChange = 0
                    YearlyChange = 0
                    
                
            End If
        
        
                            
Next i
End Sub

 'Till the above line, the code works. Afer that, something is going wrong.
        'Print PercentChange
        'Range("K1:K2").Value = PercentChange
        'Range("J1:J2").Value = YearlyChange
        

