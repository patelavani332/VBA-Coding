# VBA-Coding

... VBA Code...

Sub stock()

    Dim i As Long
    Dim j As Long
    Dim total As Double
    Dim Yearly_Change As Double
    Dim Precent_Change As Double
    Dim lastrow As Long
    

    For Each ws In Worksheets
        'To calculate the last row in every sheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
                
       j = 2
       total = 0
       Yearly_Change = 0
       percent_change = 0
       k = 2
       
       'MsgBox (lastrow)
            
        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' List tickers in the sheets
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                ' calculates total volume of stocks for a perticular ticker
                total = total + ws.Cells(i, 7).Value
                
                    If ws.Cells(k, 3).Value = 0 Then
                        ' To find the open value of tickers
                        For open_value = k To i
                        
                            If ws.Cells(k, 3).Value <> 0 Then
                            
                                k = open_value
                                
                                Exit For
                            
                            End If
                        
                         Next open_value
                         
                    End If
                    ' Calculates yearly change and percent change
                    Yearly_Change = (ws.Cells(i, 6).Value - ws.Cells(k, 3).Value)
                    percent_change = (Yearly_Change / ws.Cells(k, 3).Value)
                    
                    k = i + 1
                    ' Assign the calculated yearly change, percent change and total stock volume values to respective cells
                    ws.Cells(j, 10).Value = Yearly_Change
                    ws.Cells(j, 10).NumberFormat = "0.00"
                    ws.Cells(j, 11).Value = percent_change
                    ws.Cells(j, 11).NumberFormat = "0.00%"
                    ws.Cells(j, 12).Value = total
                
                ' Conditional formatting of the yearly change column
                    If ws.Cells(j, 10).Value > 0 Then
                 
                        ws.Cells(j, 10).Interior.ColorIndex = 4
                    
                    ElseIf ws.Cells(j, 10).Value < 0 Then
                    
                        ws.Cells(j, 10).Interior.ColorIndex = 3
                    
                    Else
                    
                        ws.Cells(j, 10).Interior.ColorIndex = 0
                    
                    End If
                
                
                j = j + 1
                total = 0
                percent_change = 0
                Yearly_Change = 0
                                     
            Else
                
                total = total + ws.Cells(i, 7).Value
                
           
            End If
                         
        
        Next i
        
        Dim max As Double
        Dim min As Double
        
        ' find the greatest percent increase
        max = WorksheetFunction.max(ws.Range("K2:K" & lastrow))
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("Q2") = max
        
        ' find the greatest percent decrease
        min = WorksheetFunction.min(ws.Range("K2:K" & lastrow))
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("Q3") = min
        
        ' find the greatest total volume
        ws.Range("Q4") = WorksheetFunction.max(ws.Range("L2:L" & lastrow))
        
        ' find the ticker for the greatest percent increase
        greatest_ticker = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        ws.Range("P2") = ws.Cells(greatest_ticker + 1, 9)
        
        ' find the ticker for the greatest percent decrease
        lowest_ticker = WorksheetFunction.Match(WorksheetFunction.min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        ws.Range("P3") = ws.Cells(lowest_ticker + 1, 9)
        
        ' find the ticker for the greatest total volume
        max_volume = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
        ws.Range("P4") = ws.Cells(max_volume + 1, 9)
        
        
              
    Next ws


End Sub
