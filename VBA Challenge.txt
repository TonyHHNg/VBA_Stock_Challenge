Sub multi_year()

  ' Loop through all worksheets '
  For Each ws In Worksheets
  
  ' Identify range'
  Dim rng As Range
  
  ' Identify variable for holding the ticker name '
  Dim t_name As String

  ' Identify variable for holding the total per stock volume '
  Dim t_total As Double
  t_total = 0
  
  ' Keep track of the location for each stock in the summary table '
  Dim T_Table_Row As Integer
  T_Table_Row = 2
   
  ' Identify varible for yearly close price '
  Dim cl_value As Double
    
  ' Identify last row in different column '
  last_row = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count
  RowCount = ws.Range("K1", ws.Range("K1").End(xlDown)).Rows.Count
  L_RowCount = ws.Range("L1", ws.Range("L1").End(xlDown)).Rows.Count
  working_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
  ' Print headers for results in each worksheet '
  ws.Range("I1") = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percentage Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("P1") = "Ticker"
  ws.Range("Q1") = "Value"
  ws.Range("O2") = "Greatest % Increase"
  ws.Range("O3") = "Greatest % Decrease"
  ws.Range("O4") = "Greatest Total Volume"
    
  ' Define the yearly open value  '
  op_value = ws.Range("C2").Value

  ' Loop through all stock purchases
  For i = 2 To last_row
          
    ' Check if we are still within the same stock
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the closing value
      cl_value = ws.Cells(i, 6).Value
         
      ' Set the stock name
      t_name = ws.Cells(i, 1).Value
          
      'Print the change between opening and closing value into the summary table
      ws.Range("J" & T_Table_Row).Value = cl_value - op_value

      'Print the percentage change between opening and closing value into the summary table
      ws.Range("K" & T_Table_Row).Value = Round(((cl_value - op_value) / op_value) * 100, 2) & "%"
        
      ' Add to the stock Total
      t_total = t_total + ws.Cells(i, 7).Value

      ' Print the Stock Name in the Summary Table
      ws.Range("I" & T_Table_Row).Value = t_name

      ' Print the stock Amount to the Summary Table
      ws.Range("L" & T_Table_Row).Value = t_total

      ' Add one to the summary table row
      T_Table_Row = T_Table_Row + 1
       
      ' Reset the Brand Total
      t_total = 0
           
      ' If the cell immediately following a row is the same brand, add the next one until different name
      op_value = ws.Cells(i + 1, 3).Value
      
    Else

      ' Add to the Brand Total
      t_total = t_total + ws.Cells(i, 7).Value
           
    End If

 Next i

   ' Loop through every J row '
   For Each rng In ws.Range("J2:J" & working_row)
   
    ' To check if the cell is a number '
    If IsNumeric(rng.Value) Then
      
      ' Set up a loop to check the condition then highlight in Red colour '
      If rng.Value < 0 Then
             
        rng.Interior.Color = vbRed
            
      ' Set up a loop to check the condition then highlight in Green colour '
      ElseIf rng.Value > 0 Then
                
        rng.Interior.Color = vbGreen
           
     End If
    
   End If
   
  Next rng

 ' Print the Maximum value of Row K '
 ws.Range("Q2").Value = (WorksheetFunction.Max(ws.Range("K2:K" & RowCount))) * 100

 ' Print the Minimum value of Row K '
 ws.Range("Q3").Value = (WorksheetFunction.Min(ws.Range("K2:K" & RowCount))) * 100

 ' Print the Maximum value of Row L '
 ws.Range("Q4").Value = (WorksheetFunction.Max(ws.Range("L2:L" & RowCount)))


    'Loop through every K row '
    For k = 2 To RowCount
      
        ' If confition to get the relevant Ticker Name '
        If ws.Cells(k, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) Then
    
            ' Print relevant Ticker Name by using offset function '
            ws.Range("P2").Value = ws.Cells(k, 11).Offset(0, -2)
         
        ' If confition to get the relevant Ticker Name '
        ElseIf ws.Cells(k, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) Then
            
            ' Print relevant Ticker Name by using offset function '
            ws.Range("P3").Value = ws.Cells(k, 11).Offset(0, -2)
            
        End If
    
    Next k

    'Loop through every K row '
    For l = 2 To L_RowCount
        
        ' If confition to get the relevant Ticker Name '
        If ws.Cells(l, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount)) Then
            
            ' Print relevant Ticker Name by using offset function '
            ws.Range("P4").Value = ws.Cells(l, 12).Offset(0, -3)
         
        End If
    
    Next l

Next ws

End Sub


