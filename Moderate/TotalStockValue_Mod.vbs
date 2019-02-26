Sub TotalStockValue_mod()
    
    'Loop through all worksheets
     For Each ws In Worksheets
       'Set initial Variables
        Dim row As Integer
        Dim m As Long
        Dim total_vol As Double
        Dim open_vol As Double
        Dim close_vol As Double
        

       'Add the header in summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Determine the last row in each worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Keep track of the row for ticker name in summary table
         row = 2
        'Keep track of the row for the first open volume of each ticker
         m = 2
        'Set an initial varible for holding the total volume
         total_vol = 0
        
        'loop through all rows in each worksheet
               
        For i = 2 To LastRow
            'Add value to the column for  Total Stock Volume
            total_vol = total_vol + ws.Cells(i, 7).Value
            ws.Cells(row, 12).Value = total_vol
            ws.Cells(row, 9).Value = ws.Cells(i, 1).Value
            open_vol = ws.Cells(m, 3).Value

            'Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Reset the total volume
                total_vol = 0
                'Set value for close volume
                close_vol = ws.Cells(i, 6).Value
                'Set value for column of Yearly Change
                ws.Cells(row, 10).Value = close_vol - open_vol

                'Set value for column of Percent Change convert the cell format to percentage
                If open_vol = 0 Then
                    ws.Cells(row, 11).Value = 0
                Else
                    ws.Cells(row, 11).Value = ws.Cells(row, 10).Value / open_vol
                    'Convert cell format to percentage
                    ws.Cells(row, 11).NumberFormat = "0.00%"
                    
                End If
                
                'Conditional highlight positive/negtive
                If ws.Cells(row, 11).Value > 0 Then
                    ws.Cells(row, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(row, 11).Interior.ColorIndex = 3
                End If
                'Reset the row of summary table
                row = row + 1
                'Reset the row of open volume
                 m = i + 1
            End If
        Next i
        
          
            
    Next ws
    
End Sub
