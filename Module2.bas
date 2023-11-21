Sub years()

   Dim ws As Worksheet
   Dim WorksheetName As String
   
For Each ws In ThisWorkbook.Sheets

   ' Setting variables and initial values
   Dim Ticker As String
   Dim Vol_Total As Double
   Vol_Total = 0
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2
   Dim open_price As Double
   open_price = 0
   Dim close_price As Double
   close_price = 0
   Dim price_change As Double
   price_change = 0
   Dim price_change_percent As Double
   price_change_percent = 0
   Dim i As Long
   Dim j As Integer
   j = 0
   Dim start As Long
   start = 2
   Dim rowCount As Long

   ' Setting column headers and table labels
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   
   ' Setting up the last row
   rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
  
  
' Loop through rows
  For i = 2 To rowCount
     If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        open_price = ws.Cells(i, 3).Value
     End If
     
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       Ticker = ws.Cells(i, 1).Value
       Vol_Total = Vol_Total + ws.Cells(i, 7).Value
       close_price = ws.Cells(i, 6).Value
       price_change = close_price - open_price
       price_change_percent = (price_change / open_price)
       
       ws.Cells(Summary_Table_Row, 9).Value = Ticker
       ws.Cells(Summary_Table_Row, 10).Value = price_change
       ws.Cells(Summary_Table_Row, 11).Value = price_change_percent
       ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
       ws.Cells(Summary_Table_Row, 12).Value = Vol_Total
     
     
       If price_change >= 0# Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
       ElseIf price_change < 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
       End If
     
       If price_change_percent >= 0# Then
            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
       ElseIf price_change_percent < 0 Then
            ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
       End If
       
       Summary_Table_Row = Summary_Table_Row + 1
       Vol_Total = 0
   
       Else
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value
        price_change = close_price - open_price
        price_change_percent = (price_change / open_price)
       
       End If
  Next i

    ' Working on the final summary chart
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

    ws.Range("P2") = Cells(increase_number + 1, 9)
    ws.Range("P3") = Cells(decrease_number + 1, 9)
    ws.Range("P4") = Cells(volume_number + 1, 9)

Next ws

End Sub



