# VBA_Challenge
Stock Market Analysis coding:

Sub StockMarket()

 ' worksheet loop
 
Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate
    
    'create the variables
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim ticker As String
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim volume As Double
    Dim opening_price As Double
    Dim closing_price As Double
    Dim row As Double
    Dim column As Double
    
    
        ' find the last row of the table
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row

        ' Column creation
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quaterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
      

        'reset the value
        
        volume = 0
        row = 2
        column = 1
       
             
        ' setting the opening price
        opening_price = Cells(2, column + 2).Value
        
          ' loop through all ticker
        For i = 2 To last_row
        
         
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            
                ' setting ticker name
                ticker = Cells(i, column).Value
                Cells(i, 9).Value = ticker
                
                  ' set the Colors as per positive and negative change
                  
        For j = 2 To ticker_last_row
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
                ' setting closing price
                closing_price = Cells(i, column + 5).Value
                
                ' calculate quarterly change
                quarterly_change = closing_price - opening_price
                
               ' calculate percent change
                    percent_change = quarterly_change / opening_price
                    Cells(row, column + 10).NumberFormat = "0.00%"
               
               
               ' calculate total volume per quarter
                volume = volume + Cells(i, column + 6).Value
                
                ' loop to the next row
                row = row + 1
                
                ' reset open price to next ticker
                opening_price = Cells(i + 1, column + 2)
                
                ' reset volume for next ticker
                volume = 0
                
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i
        
        
        ' find the last row of ticker column
        ticker_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
      
        ' BONUS: set Ticker, Value, Greatest %, Increase, % Decrease, and Total volume headers
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        ' find the highest value of each ticker
        For k = 2 To ticker_last_row
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & ticker_last_row)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & ticker_last_row)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & ticker_last_row)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k
        
    Next w
End Sub


'References:: Stack_overflow
Week_2:Day_3- Class activities
https://www.tutorialspoint.com/vba/index.htm
chatGpt
![image](https://github.com/user-attachments/assets/3d0fe156-c93a-4027-81f0-8667d5eff2d4)
