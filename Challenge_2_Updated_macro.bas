Attribute VB_Name = "Module1"
Sub sample_stock():

    For Each ws In Worksheets

       
        Dim WorksheetName As String
        Dim LastRow As Double
       
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        WorksheetName = ws.Name
        
    
      
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
       
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim ticker_symbol As String
        Dim yearly_change As Double
        Dim summary_table_row As Integer
        Dim opening_price As Double
        Dim closing_price As Double
        Dim total_volume As Double
        
    
        
        summary_table_row = 2
        
        
        
        
        For i = 2 To LastRow
            
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                If i > 2 Then
                    ws.Cells(summary_table_row - 1, 12).Value = total_volume
                End If
                    
            
                ticker_symbol = ws.Cells(i, 1).Value
                opening_price = ws.Cells(i, 3).Value
                total_volume = 0
                
            End If
                total_volume = total_volume + CDbl(ws.Cells(i, 7).Value)
            
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                closing_price = ws.Cells(i, 6).Value
                yearly_change = closing_price - opening_price
                
            
                ws.Cells(summary_table_row, 9).Value = ticker_symbol
                ws.Cells(summary_table_row, 10).Value = yearly_change
                
                 If yearly_change < 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3  ' Red
                ElseIf yearly_change > 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4 'Green
                End If
                
                If opening_price <> 0 Then
                    ws.Cells(summary_table_row, 11).Value = Round((yearly_change / opening_price) * 100, 2) & "%"
                Else
                    ws.Cells(summary_table_row, 11).Value = "0.00%"
                End If
                
                
                
                summary_table_row = summary_table_row + 1
                
         
        
                
        
            End If
            
            
    
                
        Next i
                
                
          ws.Cells(summary_table_row - 1, 12).Value = total_volume
          ws.Range("P1") = "Ticker"
          ws.Range("Q1") = "Value"
         ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
         ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
         ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
         ws.Range("Q4").NumberFormat = "0.00"
         increase_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
         ws.Range("P2") = Cells(increase_index + 1, 9)
         decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
         ws.Range("P3") = Cells(decrease_index + 1, 9)
         volume_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)
         ws.Range("P4") = Cells(volume_index + 1, 9)
        



Next ws

End Sub


