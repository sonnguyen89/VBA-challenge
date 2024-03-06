Sub Process_tickers()

    ' Loop through all sheets
    For Each ws In Worksheets
        '----------------------------
        'reset the columns by  delete their content before update
        ' Specify the delete column
        Set IColumn = ws.Columns("I")
        Set JColumn = ws.Columns("J")
        Set KColumn = ws.Columns("K")
        Set LColumn = ws.Columns("L")
        Set OColumn = ws.Columns("O")
        Set PColumn = ws.Columns("P")
        Set QColumn = ws.Columns("Q")
        
        
        ' Delete the entire column
        IColumn.Delete
        JColumn.Delete
        KColumn.Delete
        LColumn.Delete
        OColumn.Delete
        PColumn.Delete
        QColumn.Delete
        
        
        '---------------------------
        ' Find the last row of the sheet after each paste
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        NumberOfRowWithoutHeader = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        
         ' Find the last used cell in the first row (row 1)
        lastColumnIndex = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        Set Rng = ws.Range("I1") 'new column ticker
        'extract the unique value from column A to Column Ticker
        ws.Range("A1:A" & LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True
        ws.Range("I1").Value = "Ticker"
       
        Dim value_open As Double
        Dim value_close As Double
        Dim yearly_change As Double
        
        EndRow = Rng.End(xlDown).Row
        ws.Range("J1").Value = "Yearly Change"
        For x = 2 To EndRow
            value_open = WorksheetFunction.SumIf(ws.Range("A1:A" & LastRow), ws.Cells(x, 9), ws.Range("C:C"))
            value_close = WorksheetFunction.SumIf(ws.Range("A1:A" & LastRow), ws.Cells(x, 9), ws.Range("F:F"))
            total_stock_volumn = WorksheetFunction.SumIf(ws.Range("A1:A" & LastRow), ws.Cells(x, 9), ws.Range("G:G"))
            
            'add value to column yearly change, percentage change and total stock volumn
            yearly_change = value_close - value_open
            yearly_change_rate = yearly_change / value_open
            ws.Cells(x, 10).Value = yearly_change
            ws.Cells(x, 11).Value = yearly_change_rate * 100
            ws.Cells(x, 12).Value = total_stock_volumn
        Next x
        
        'set value for Percent Change Column
        ws.Range("K1").Value = "Percent Change"
        Set targetColumn = Columns("K")
        '--------------------------------
        'update cell colour based on the value
        LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row 'find the last row index of yearly change
        Set targetRange = ws.Range("J2:J" & LastRow) ' Change this to the range you want
        ' Loop through each cell in the target range
        For Each cell In targetRange
            ' Check if the cell value is positive, negative, or zero
            If cell.Value > 0 Then
                ' Positive value, set cell color to green
                cell.Interior.Color = RGB(0, 255, 0) ' RGB color for green
            ElseIf cell.Value < 0 Then
                ' Negative value, set cell color to red
                cell.Interior.Color = RGB(255, 0, 0) ' RGB color for red
            Else
                ' Zero value, clear cell color
                cell.Interior.ColorIndex = xlNone
            End If
        Next cell
        '--------------------------------
        'set number format to percentage
        targetColumn.NumberFormat = "0.00%"
        
       'set value for Total  Stock Volumn Column
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim smallest As Double
        Dim biggest As Double
        Dim biggest_total_volume As LongLong
        
        biggest = 0
        smallest = 0
        biggest_ticker = ""
        smallest_ticker = ""
        biggest_total_volume_ticker = ""
        
        For x = 2 To EndRow
            If ws.Cells(x, 11).Value > biggest Then
                biggest = ws.Cells(x, 11).Value
                biggest_ticker = ws.Cells(x, 9).Value
            End If
             If ws.Cells(x, 11).Value <= smallest Then
                smallest = ws.Cells(x, 11).Value
                smallest_ticker = ws.Cells(x, 9).Value
            End If
            If ws.Cells(x, 12).Value > biggest_total_volume Then
                biggest_total_volume = ws.Cells(x, 12).Value
                biggest_total_volume_ticker = ws.Cells(x, 9).Value
            End If
         Next x
         
        ws.Range("Q2").Value = biggest
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("Q3").Value = smallest
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("Q4").Value = biggest_total_volume
        
        ws.Range("P2").Value = biggest_ticker
        ws.Range("P3").Value = smallest_ticker
        ws.Range("P4").Value = biggest_total_volume_ticker
        
    
    Next ws
    
    
End Sub
