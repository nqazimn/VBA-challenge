Sub VBAStock_Hard()
    ' Declarations: Data related variables
    Dim this_ticker As String
    Dim next_ticker As String
    Dim opening_price, closing_price As Double
    Dim yearly_change As Double
    Dim total_stock_volume, stock_volume As Long
    Dim max_change, min_change As Double
    Dim max_total As Long
    Dim max_change_idx, min_change_idx, max_total_idx As Long
    
    ' Declarations: Loop variables
    Dim row_iterator As Long
    Dim ticker_counter As Long
    
    'Dim work_sheet As Worksheet
    
    'For Each work_sheet In Worksheets
        'work_sheet.Activate
    
        ' Clear, label headers and setup output columns
        'Call delete_output_columns
        ActiveSheet.Range("I:Q").Delete
        ActiveSheet.Range("I1").Value = "Ticker"
        ActiveSheet.Range("J1").Value = "Yearly Change"
        ActiveSheet.Range("K1").Value = "Percent Change"
        ActiveSheet.Range("L1").Value = "Total Stock Volume"
        ActiveSheet.Range("O2").Value = "Greatest % Increase"
        ActiveSheet.Range("O3").Value = "Greatest % Decrease"
        ActiveSheet.Range("O4").Value = "Greatest Total Volume"
        ActiveSheet.Range("P1").Value = "Ticker"
        ActiveSheet.Range("Q1").Value = "Value"
        
        ActiveSheet.Range("I:L").Columns.AutoFit
        ActiveSheet.Range("O:Q").Columns.AutoFit
        
        ' Set everything to zero to for first iteration
        max_change = 0
        min_change = 0
        max_total = 0
        
        ' Setting indices to -1 for first iteration
        max_change_index = -1
        min_change_index = -1
        max_total_index = -1
        
        row_iterator = 2
        ticker_counter = 1
        total_stock_volume = 0
        
        ' Save data for first ticker to begin the loop
        this_ticker = ActiveSheet.Cells(row_iterator, 1)
        opening_price = ActiveSheet.Cells(row_iterator, 3)
        
        Do While (this_ticker <> "")    ' Loop until end of record
            this_ticker = ActiveSheet.Cells(row_iterator, 1)
            next_ticker = ActiveSheet.Cells(row_iterator + 1, 1)
            
            stock_volume = ActiveSheet.Cells(row_iterator, 7)
                    
            total_stock_volume = total_stock_volume + stock_volume
            
            If (next_ticker <> this_ticker) Then
                ' If next ticker is different, then grab the closing price of this_ticker
                closing_price = ActiveSheet.Cells(row_iterator, 6)
                
                ' Increment ticker_counter for dumping result in sheet
                ticker_counter = ticker_counter + 1
                
                yearly_change = closing_price - opening_price
                
                ' Dump data
                ActiveSheet.Cells(ticker_counter, 9) = this_ticker
                ActiveSheet.Cells(ticker_counter, 10) = Format(yearly_change, "0.00")
                
                ' Calculate percent change
                If (opening_price <> 0) Then
                    ActiveSheet.Cells(ticker_counter, 11) = Format(yearly_change / Abs(opening_price), "0.00%")
                ElseIf (opening_price = 0) Then
                    ' If opening price is 0, the above formula is not valid.
                    ActiveSheet.Cells(ticker_counter, 11) = "N/A"
                End If
                
                ' Conditional formatting Yearly Change Cell
                If (yearly_change < 0) Then
                    ActiveSheet.Cells(ticker_counter, 10).Select
                    Selection.Interior.ColorIndex = 3 'Red interior
                    Selection.Font.Bold = True
                    Selection.Font.ColorIndex = 1
                Else
                    ActiveSheet.Cells(ticker_counter, 10).Select
                    Selection.Interior.ColorIndex = 4 'Green interior
                    Selection.Font.Bold = True
                End If
                
                ' Dump total_stock_volume in corresponding cell
                ActiveSheet.Cells(ticker_counter, 12).Value = total_stock_volume
                
                ' Set total back to 0 to re-start summation for new ticker value
                total_stock_volume = 0
                
                ' Update opening price for new ticker
                opening_price = ActiveSheet.Cells(row_iterator + 1, 3)
                
            End If
              
            row_iterator = row_iterator + 1
        Loop
                    
        ' Find maximum value of percent_increase
        ActiveSheet.Range("Q2").Value = WorksheetFunction.Max(Range("K:K"))
        ActiveSheet.Range("Q2").NumberFormat = "0.00%"
        ActiveSheet.Range("Q2").Columns.AutoFit
            
        ' Find index of maximum value in percent_increase and place in corresponding cell
        max_total_idx = WorksheetFunction.Match(WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
        ActiveSheet.Range("P2").Value = ActiveSheet.Range("I" & max_total_idx)
            
        ' Find minimum value of percent_increase
        ActiveSheet.Range("Q3").Value = WorksheetFunction.Min(Range("K:K"))
        ActiveSheet.Range("Q3").NumberFormat = "0.00%"
        ActiveSheet.Range("Q3").Columns.AutoFit
            
        ' Find index of minimum value in percent_increase and place in corresponding cell
        min_total_idx = WorksheetFunction.Match(WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
        ActiveSheet.Range("P3").Value = ActiveSheet.Range("I" & min_total_idx)
            
        ' Find maximum value of total_stock_volume
        ActiveSheet.Range("Q4").Value = WorksheetFunction.Max(Range("L:L"))
        ActiveSheet.Range("Q4").NumberFormat = "0.0000E+00"
        ActiveSheet.Range("Q4").Columns.AutoFit
            
        ' Find index of maximum value in total_stock_volume and place in corresponding cell
        max_total_idx = WorksheetFunction.Match(WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
        ActiveSheet.Range("P4").Value = ActiveSheet.Range("I" & max_total_idx)
            
        ' Take cursor to result area of the sheet
        ActiveSheet.Cells(2, 11).Select
            
    'Next work_sheet
    
    MsgBox ("DONE!!!")
End Sub


