Sub StockAnalysis()

Dim i As Long
Dim j As Long
Dim dates() As Variant
Dim rowIndex() As Variant
Dim ticker As String
Dim counter As Integer
Dim counter2 As Integer
Dim openPrice As Double
Dim closePrice As Double
Dim temp As Integer





'Counts the number of rows
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Pulls unique tickers and lists them in column i
ActiveSheet.Range("A1:A" & LastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("I1:I" & LastRow), Unique:=True

'Changes headers in new columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

LastUniqueRow = Cells(Rows.Count, 9).End(xlUp).Row

For j = 2 To LastUniqueRow

    ticker = Cells(j, 9).Value
    counter = 0
    temp = 0
    openPrice = 0
    closePrice = 0
    'Counts the number of values per ticker
    For i = 2 To LastRow
        If Cells(i, 1).Value = ticker Then
            counter = counter + 1
        End If
    Next i

    'Creates an array full of dates for each ticker entry and indexes which rows you can find the current ticker
    ReDim dates(counter)
    ReDim rowIndex(counter)
    counter2 = 0

    For i = 2 To LastRow
        If Cells(i, 1).Value = ticker Then
            dates(temp) = Cells(i, 2).Value
            rowIndex(temp) = i
            temp = temp + 1
        End If
    Next i

    'Determines the opening price on the first day and the closing price on the last day
    For i = rowIndex(0) To rowIndex(temp - 1)
        If Cells(i, 2).Value = dates(0) Then
            openPrice = Cells(i, 3).Value
        ElseIf Cells(i, 2).Value = dates(counter - 1) Then
            closePrice = Cells(i, 6).Value
        End If
    Next i

    'Calculates percent change and writes it to Column J
    Cells(j, 10).Value = closePrice - openPrice
    Cells(j, 11).Value = (closePrice - openPrice) / openPrice
    Cells(j, 12).Value = Application.Sum(Range(Cells(rowIndex(0), 7), Cells(rowIndex(temp - 1), 7)))
    
    
Next j




        

End Sub

