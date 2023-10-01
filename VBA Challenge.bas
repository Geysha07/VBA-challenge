Attribute VB_Name = "Module1"
Sub Stock_Analysis()

'Set worksheet as variable

Dim ws As Worksheet

    'Run code on every worksheet in the workbook
    
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
 
        'Label columns for Analysis Table
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        'Set variables
        
        Dim ticker_ID As String
        Dim total_stock_volume As Double
        Dim year_open As Double
        Dim year_close As Double
        Dim yearly_change As Double
        Dim percentage_change As Double
        Dim summary_table_row As Long
        opening_price_start = 2
        summary_table_row = 2

        'Define last row of each worksheet

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set stock volume counter to 0
          
        total_stock_volume = 0
        
        'Run code until you reach the last row in the worksheet
        
        For i = 2 To lastrow

            'If ticker ID for next cell is different, then add the ticker ID, yearly change change, percent change, and total stock volume for each ticker to the summary table
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker_ID = ws.Cells(i, 1).Value
                year_open = ws.Cells(opening_price_start, 3).Value
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
                percentage_change = (yearly_change / year_open) * 100
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                ws.Range("I" & summary_table_row) = ticker_ID
                ws.Range("J" & summary_table_row) = yearly_change
                ws.Range("K" & summary_table_row) = percentage_change
                ws.Range("L" & summary_table_row) = total_stock_volume
          
            'Highlight cells in yearly change and percent change green if positive or red if negative
        
            If Cells(summary_table_row, 10).Value > 0 Then
                Cells(summary_table_row, 10).Interior.ColorIndex = 4
                
            ElseIf Cells(summary_table_row, 10).Value < 0 Then
                Cells(summary_table_row, 10).Interior.ColorIndex = 3
                
            End If
            
            If Cells(summary_table_row, 11).Value > 0 Then
                Cells(summary_table_row, 11).Interior.ColorIndex = 4
                
            ElseIf Cells(summary_table_row, 11).Value < 0 Then
                Cells(summary_table_row, 11).Interior.ColorIndex = 3
                
            End If
                 
                'Add a new row to summary table
                
                summary_table_row = summary_table_row + 1
                
                'Check the next ticker opening price
                
                opening_price_start = i + 1
               
               'Reset stock volume counter to 0
                
                total_stock_volume = 0
                
            'If ticker ID for the next cell is the same, then start adding up the volumes again
               
            Else: ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If
            
        'Move to next row
        
        Next i
        
    'Find value and ticker ID of ticker with the greatest % increase
    
    max_percent_value = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow))
    Range("Q2") = max_percent_value
    max_percent_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    Range("P2") = Cells(max_percent_index + 1, 9)
    
    'Find value and ticker ID of ticker with the greatest % decrease
    
    min_percent_value = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow))
    Range("Q3") = min_percent_value
    min_percent_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    Range("P3") = Cells(min_percent_index + 1, 9)
    
    'Find value and ticker ID of ticker with the greatest total volume
    
    max_stock_volume = WorksheetFunction.Max(Range("L2:L" & lastrow))
    Range("Q4") = max_stock_volume
    max_stock_volume_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
    Range("P4") = Cells(max_stock_volume_index + 1, 9)

    'Move to next worksheet and run code again
    Next ws

'Once finished running code through all worksheets, end subroutine

End Sub






