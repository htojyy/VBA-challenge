Sub stock_data():

    Dim ws As Worksheet
    
    Dim ticker As String
    Dim open_price, close_price, pct_change, yearly_change As Double
    Dim max_pct_change, min_pct_change As Double
    
    Dim LastRow_ticker, LastRow_smy As Long
    Dim ticker_counter As Long
    Dim max_pct_change_row, min_pct_change_row, max_stock_vol_row As Long
    
    '1. https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/longlong-data-type
    Dim stock_volume, max_stock_vol As LongLong

    Dim pct_change_range, stock_vol_range As Range
    
    Dim i As Long
    
    'Loop through each worksheet and get the yearly change, % change, total stock volume
    For Each ws In Worksheets
    
        '2. https://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/
        ws.Activate
        
        'Get last row cell value
        LastRow_ticker = Cells(Rows.Count, 1).End(xlUp).Row
        LastRow_smy = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Set/Reset initial values of numeric variables for each worksheet loop
        open_price = 0
        close_price = 0
        yearly_change = 0
        pct_change = 0
        stock_volume = 0
        ticker_counter = 0
        smy_tbl_row = 2
        
        'Set new table headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        '3. https://learn.microsoft.com/en-us/office/vba/api/excel.font.bold
        Range("I1:P1").Font.Bold = True
                
        For i = 2 To LastRow_ticker
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                'Get and print ticker name - 1 line per ticker name
                ticker = Cells(i, 1).Value
                Range("I" & smy_tbl_row).Value = ticker
                
                'Get and print total stock volume for each ticker
                stock_volume = stock_volume + Cells(i, 7).Value
                Range("L" & smy_tbl_row).Value = stock_volume
                
                'Get open and close price for each ticker
                '4. https://www.automateexcel.com/vba/offset-range-cell
                open_price = Cells(i, 3).Offset(-ticker_counter, 0).Value
                close_price = Cells(i, 6).Value
    '            Range("N" & smy_tbl_row).Value = open_price
    '            Range("M" & smy_tbl_row).Value = close_price
                
                'Get and print yearly change
                yearly_change = close_price - open_price
                Range("J" & smy_tbl_row).Value = yearly_change
    
                'Get and print Percent change
                percent_change = yearly_change / open_price
                fmt_percent_change = Format(percent_change, "Percent")
                Range("K" & smy_tbl_row).Value = fmt_percent_change
                
                'Cell background formatting: if % change >0 then green, if <0 then red
                If percent_change > 0 Then
                    Range("K" & smy_tbl_row).Interior.ColorIndex = 4
                    Range("J" & smy_tbl_row).Interior.ColorIndex = 4
                    
                ElseIf percent_change < 0 Then
                    Range("K" & smy_tbl_row).Interior.ColorIndex = 3
                    Range("J" & smy_tbl_row).Interior.ColorIndex = 3
                    
                End If
    
                smy_tbl_row = smy_tbl_row + 1
                
                'Reset counters for each row loop
                stock_volume = 0
                close_price = 0
                open_price = 0
                ticker_counter = 0
            
            Else
                'If ticker values are the same, sum stock volume and get total number of 'rows' of ticker names
                stock_volume = stock_volume + Cells(i, 7).Value
                ticker_counter = ticker_counter + 1
     
            End If
            
        Next i
    
        'Set range for pct change and total stock volume (smy table)
        Set pct_change_range = ws.Range("K2:K" & LastRow_smy)
        Set stock_vol_range = ws.Range("L2:L" & LastRow_smy)
        
        'Find min/max pct change & max stock volume figures
        '5. https://forum.ozgrid.com/forum/index.php?thread/87224-min-max-values-from-row-column-using-vba/
        max_pct_change = Application.WorksheetFunction.Max(pct_change_range)
        min_pct_change = Application.WorksheetFunction.Min(pct_change_range)
        max_stock_vol = Application.WorksheetFunction.Max(stock_vol_range)
        
        'Get row number for each of the summary stats
        '6. https://forum.ozgrid.com/forum/index.php?thread/1228100-find-row-of-max-value-copy-and-paste/
        max_pct_change_row = Application.WorksheetFunction.Match(max_pct_change, pct_change_range, 0) + pct_change_range.Row - 1
        min_pct_change_row = Application.WorksheetFunction.Match(min_pct_change, pct_change_range, 0) + pct_change_range.Row - 1
        max_stock_vol_row = Application.WorksheetFunction.Match(max_stock_vol, stock_vol_range, 0) + stock_vol_range.Row - 1
'        MsgBox "Row number of max pct change is " & max_pct_change_row
        
        'Format % change numbers
        '7. https://www.vbforums.com/showthread.php?606675-RESOLVED-Format-variable-as-percentage
        fmt_max_pct_change = Format(max_pct_change, "Percent")
        fmt_min_pct_change = Format(min_pct_change, "Percent")
        
        'Add value of the 3 summary stats to relevant cells
        Range("P2").Value = fmt_max_pct_change
        Range("P3").Value = fmt_min_pct_change
        Range("P4").Value = max_stock_vol
        
        '8. https://stackoverflow.com/questions/8153800/how-to-set-numberformat-to-number-with-0-decimal-places
        Range("P4").NumberFormat = "0"
        
        'Get and Add ticker name to each of the 3 summary stats
        max_pct_change_ticker = Range("I" & max_pct_change_row).Value
        min_pct_change_ticker = Range("I" & min_pct_change_row).Value
        max_stock_vol_ticker = Range("I" & max_stock_vol_row).Value

        Range("O2").Value = max_pct_change_ticker
        Range("O3").Value = min_pct_change_ticker
        Range("O4").Value = max_stock_vol_ticker
   
        'Reset counters
        LastRow_ticker = 0
        LastRow_smy = 0
        
        'Adjust width of columns
        '9. https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
        ws.Columns("I:P").AutoFit
        
    Next ws


End Sub
