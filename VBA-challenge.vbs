Sub stocks():
    'Create variables to hold all values
    Dim ws As Worksheet
    Dim summary_table_row As Integer
    Dim ticker As String
    Dim opening As Double
    Dim closing As Double
    Dim vol As Double
    Dim change As Double
    Dim percent As Double
    Dim percent_chng As String
    
    For Each ws In Sheets
        'Give summary table columns headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
    
        'Exclude header from analysis
        summary_table_row = 2
        
        'Find last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all data
        For i = 2 To lastrow
        
        'Get opening cost (we moved to a new ticker)
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Get ticker and add to summary table
                ticker = ws.Cells(i, 1).Value
                ws.Cells(summary_table_row, 9).Value = ticker
                
                'Get opening value
                opening = ws.Cells(i, 3).Value
                
                'Get Volumes
                vol = 0
        
            End If
            
            'increment volumes outside If statement
            vol = vol + ws.Cells(i, 7).Value
            
            'If a cell is equal to previous cell (we are about to move to next ticker)
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Get closing
                closing = ws.Cells(i, 6).Value
                
                'Calculate yearly change
                change = closing - opening
                
                'Calculate percent change
                If change = 0 Then
                    percent = 0
                Else
                    percent = ((closing - opening) / opening)
                End If
                
                'Color the cells red and green
                If change >= 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
                End If
               
               'Write values to summary table
                ws.Cells(summary_table_row, 11).Value = percent
                ws.Cells(summary_table_row, 12).Value = vol
                ws.Cells(summary_table_row, 10).Value = change
                
                'Format percent
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                
                'Reset the variables, jump to next row
                summary_table_row = summary_table_row + 1
    
            End If
        Next i
    Next ws
End Sub

