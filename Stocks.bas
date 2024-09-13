Sub Stock()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ws.[I1] = "Ticker"
        ws.[J1] = "Quaterly Change"
        ws.[K1] = "Percent Change"
        ws.[L1] = "Total Stock Volume"
        ws.[O2] = "Greatest % Increase"
        ws.[O3] = "Greatest % Decrease"
        ws.[O4] = "Greatest Total Volume"
        ws.[P1] = "Ticker"
        ws.[Q1] = "Value"
        
        si = 2
        Total = 0
        Inc_Value = 0
        Dec_Value = 0
        first_open = 0
        Greatest_Total = 0
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastRow
        

        'Total Stock Value
            Total = Total + ws.Cells(i, "G")
            
            If first_open = 0 Then
                first_open = ws.Cells(i, "C")
            End If
        
        'Lastrow of the Ticker
            If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
                
                ws.Cells(si, "I") = ws.Cells(i, "A")
        'Quaterly Change
                qCh = ws.Cells(i, "F") - first_open
                ws.Cells(si, "J") = qCh
                
        'Quaterly Change Color
                If qCh > 0 Then
                    ws.Cells(si, "J").Interior.ColorIndex = 4
                End If
                
                If qCh < 0 Then
                    ws.Cells(si, "J").Interior.ColorIndex = 3
                End If
                
                
        'Percentage Change
                pCh = (qCh / first_open)
                ws.Cells(si, "K") = pCh
            
        'Greates Increase %
            If pCh > Inc_Value Then
            
                Inc_Value = pCh
                ws.Range("P2") = ws.Cells(i, "A")
                ws.Range("Q2") = Inc_Value
            
            End If
        
        'Greatest Decreased %
            If pCh < Dec_Value Then
        
                Dec_Value = pCh
                ws.Range("P3") = ws.Cells(i, "A")
                ws.Range("Q3") = Dec_Value
            
            End If
                
        'Total Stock Volume
                ws.Cells(si, "L") = Total
                        
        'Greatest Total
            If Total > Greatest_Total Then
                
                Greatest_Total = Total
                ws.Range("Q4") = Greatest_Total
                ws.Range("P4") = ws.Cells(i, "A")
            
            End If
            
        'Reset area
                si = si + 1
                first_open = 0
                Total = 0
            End If
        Next i
        
        ws.Columns.AutoFit
        ws.Columns("K").NumberFormat = "##.##%"
        ws.Range("Q2:Q3").NumberFormat = "##.##%"
        ws.Range("Q4").NumberFormat = "$#,###"
    
    Next ws
    
End Sub


