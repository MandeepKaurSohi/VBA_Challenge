STOCK_MARKET ANALYSIS CODING :: VBA CHALLENGE

MODULE 1_ANALYSIS TOGET THE TICKER SYMBOL

    
Sub StockMarket()
'worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate


'create the variables
Dim i As Long
Dim ticker As String
Dim row As Integer

'set the row value
row = 2
 'find the last row of the table
 last_row = ws.Cells(Rows.Count, 1).End(xlUp).row
 
 'column creation
 Cells(1, 9).Value = "Ticker"
 Cells(1, 10).Value = "Quarterly Change"
 Cells(1, 11).Value = "Percent Change"
 Cells(1, 12).Value = "Total Stock Volume"
 
 'loop through ticker
 
 For i = 2 To last_row
 
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 ticker = Cells(i, 1).Value
 
 ' add ticker into table
   Range("I" & row).Value = ticker
 
 'reset the value
 row = row + 1
 End If
 Next i
 Next
End Sub


MODULE 2_analysis for quarterly change

Sub StockMarket_Quart()

'worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'create thevariables
Dim i As Long
Dim j As Long
Dim quarterly_change As Double
Dim opening_price As Double
Dim closing_price As Double
Dim row As Integer
 ' set the row value
 row = 2
 
 'find th elast row
 last_row = Cells(Rows.Count, 1).End(xlUp).row
  'initial the opening price
  opening_price = Cells(2, 3).Value

   
   'loop through the ticker
   For i = 2 To last_row
     
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
   'set the close price
   closing_price = Cells(i, 6).Value
   
   'calculate the quarterly change
   quarterly_change = (closing_price - opening_price)
   Range("j" & row).Value = quarterly_change
   
   'loop to next row
   row = row + 1
   opening_price = Cells(i + 1, 3).Value
   End If
 
   Next i
Next
End Sub



MODULE 3_' set the colors

Sub StockMarket_colors()

'worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate
last_row = Cells(Rows.Count, 9).End(xlUp).row

'iniate the loop
For i = 2 To last_row
If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
Else
Cells(i, 10).Interior.ColorIndex = 3
End If
Next i
Next

End Sub


MODULE 4_Analysis for Percent Change

Sub StockMarket_Percent()
'worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'variables creation
Dim opening_price As Double
Dim closing_price As Double
Dim percent_change As Double
Dim row As Integer
'initial th erow value
row = 2

'find the last row
last_row = Cells(Rows.Count, 1).End(xlUp).row

' set the open price
opening_price = Cells(2, 3).Value

'iniate the loop
For i = 2 To last_row
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
closing_price = Cells(i, 6).Value

    If opening_price <> 0 Then
    percent_change = (closing_price - opening_price) / opening_price
    Cells(row, 11).NumberFormat = "0.00%"
    End If
    Range("k" & row).Value = percent_change
    
    'loop to next row
    row = row + 1
    opening_price = Cells(i + 1, 3)
    End If
    Next i
Next

End Sub


MODULE 5_' Analysis for total stock volume

 Sub StcokMarket_Volume()
 'worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

'create the variables
Dim i As Long
Dim volume  As Double
Dim opening_price As Double
Dim closing_price As Double
Dim row As Integer
  
  'set the row and volume
  row = 2
  volume = 0
  
  'find th elast row
  last_row = Cells(Rows.Count, 1).End(xlUp).row
  
  ' start the loop
  For i = 2 To last_row
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      volume = volume + Cells(i, 7).Value
      Range("L" & row).Value = volume
      closing_price = Cells(i, 6).Value
      
      ' loop to next row
      row = row + 1
      volume = 0
      opening_price = Cells(i + 1, 3)
      Else
      volume = volume + Cells(i, 7).Value
      End If
      Next i
      Next
    
      
 End Sub


MODULE 6_analysis forBonu: greatest % increase, decrease andgreatest total volume

Sub StockMarket_Bonus()
  'worksheet loop
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

' create the variables
Dim i As Long
Dim max_percent_change As Double
Dim min_percent_change As Double
Dim max_volume As Double
Dim last_row As Long

'find the last row
last_row = Cells(Rows.Count, 11).End(xlUp).row

'column creation
Cells(2, 15).Value = "Greates % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = " Greatest Total volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "value"

max_percent_change = Application.WorksheetFunction.max(Range("K2:K" & last_row))
min_percent_change = Application.WorksheetFunction.Min(Range("K2:K" & last_row))
max_volume = Application.WorksheetFunction.max(Range("L2:L" & last_row))

'initial the loop

For i = 2 To last_row
If Cells(i, 11).Value = max_percent_change Then
Cells(2, 16).Value = Cells(i, 9).Value
Cells(2, 17).Value = Cells(i, 11).Value
Cells(2, 17).NumberFormat = "0.00%"

ElseIf Cells(i, 11).Value = min_percent_change Then
Cells(3, 16).Value = Cells(i, 9).Value
Cells(3, 17).Value = Cells(i, 11).Value
Cells(3, 17).NumberFormat = "0.00%"

ElseIf Cells(i, 12).Value = max_volume Then
Cells(4, 16).Value = Cells(i, 9).Value
Cells(4, 17) = Cells(i, 12).Value
End If
Next i
Next

End Sub
![image](https://github.com/user-attachments/assets/44511504-e649-440d-8fa3-534f3f1687ca)


'References:: Stack_overflow, 
Week_2:Day_3- Class activities, 
https://www.tutorialspoint.com/vba/index.htm , 
chatGpt

