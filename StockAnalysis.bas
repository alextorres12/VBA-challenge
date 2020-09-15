Attribute VB_Name = "Module1"
Sub StockAnalysis()

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

MsgBox (Str(LastRow))

End Sub
