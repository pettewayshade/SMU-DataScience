Attribute VB_Name = "StockTotalScript"
Sub Stock_Totals()
    Dim tickerArray() As Variant
    Dim startRow As Long
    Dim lastRow As Long
    Dim tickerRow As Long
    
    
''''Get number of worksheets in workbook
wrkSheetNum = ActiveWorkbook.Worksheets.Count

'''Starting loop at worksheet 1 to last worksheet
For i = 1 To wrkSheetNum
        '''Activate worksheet to run macro on
        Worksheets(i).Activate
        '''Get unqiue values, copy them to column H2, then store them in an array
        Range("A1", Range("A1").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
        '''Get the size of the ticker column
        tickersize = WorksheetFunction.CountA(Columns(9))
        
        '''Resize the array to the tickersize
        ReDim tickerArraySZ(tickersize)
        
            For t = 2 To tickersize
                tickerArraySZ(t) = Cells(t, 9).Value
                tickerArray = tickerArraySZ
            Next t
        
            '''set column headers
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            
            
            For Each ticker In tickerArray()
                '''Gets start and end row for each ticker value
                startRow = Range("A:A").Find(what:=ticker, after:=Range("A1")).Row
                lastRow = Range("A:A").Find(what:=ticker, after:=Range("A1"), LookAt:=xlWhole, searchDirection:=xlPrevious).Row
                tickerRow = Range("I:I").Find(what:=ticker, after:=Range("I1")).Row
                '''Yearly change
                Cells(tickerRow, 10).Value = (Cells(lastRow, 6) - Cells(startRow, 3))
                '''Percent change
                If Cells(startRow, 3) > 0 Then
                    Cells(tickerRow, 11).Value = Format(((Cells(lastRow, 6) - Cells(startRow, 3)) / Cells(startRow, 3)), "0.000000000%")
                End If
                ''''sum up cells for total stock volume for each ticker
                Cells(tickerRow, 12).Value = Application.WorksheetFunction.Sum(Range(Cells(startRow, 7), Cells(lastRow, 7)))
            Next ticker
            
            For n = 2 To tickersize
                If Cells(n, 10).Value > 0 Then
                    Cells(n, 10).Interior.Color = vbGreen
                Else
                    Cells(n, 10).Interior.Color = vbRed
                End If
            Next n
Next i
End Sub
