Sub mult_year_stock_data()

'Create a loop that goes through all the worksheets
Dim ws As Worksheet

'Work way throuh each sheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

'Find the last row of data
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Column Headers
Range("I1").Value = "ticker"
Range("J1").Value = "yearly change"
Range("K1").Value = "percent change"
Range("L1").Value = "total stock volume"
Range("O2").Value = "greatest % increase"
Range("O3").Value = "greatest % decrease"
Range("O4").Value = "greatest total volume"
Range("P1").Value = "ticker"
Range("Q1").Value = "value"

'Declare your Variables:
Dim ticker_symbol As String

'Set an intial variab;e for holding the total per ticker symbol
Dim ticker_volume As Double
ticker_volume = 0

'Set open value as a variable
Dim open_year As Double
open_year = 0

'Set Close Value as Double
Dim end_year As Double
end_year = 0

'Keep track of each location of each each summary table: ticker, total vol, yearly change, % change
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

    'loop through all ticker symbols
    For i = 2 To lastrow

        'Check if we are still within the same ticker symbol, if it is not...
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'set ticker
            ticker_symbol = ws.Cells(i, 1).Value
            
            'Add to the total stock volume
            ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            
            'Print the ticker symbol in the summary table
            Range("I" & Summary_Table_Row).Value = ticker_symbol
            
            'Print the Total volume in the summary table
            Range("L" & Summary_Table_Row).Value = ticker_volume
            
            'Add One to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Ticker volume total
            ticker_volume = 0
            
        'If the cell immediately following the row is the same brand
        Else
            'add to the ticker volume total
            ticker_volume = ticker_volume + ws.Cells(i, 7).Value
        End If
        
        
      'Calculate the yearly change to the right cell and set color for positive or negative
        ws.Cells(Summary_Table_Row, 10).Value = (ws.Cells(i, 6).Value - open_year)
        If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(Summary_Table_Row, 10).Value > 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else: ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 6
        End If

    'Calculate the percent change and write to the appropriate cell and set result to two decimals
        If open_year = 0 Then
            ws.Cells(Summary_Table_Row, 11).Value = "NaN"
        Else
            ws.Cells(Summary_Table_Row, 11).Value = ((ws.Cells(i, 6).Value - open_year) / open_year)
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        End If
        
'Sets the opening year value to the new value
    open_year = ws.Cells(i + 1, 3).Value
    
    
    Next i
    
Next ws


End Sub

