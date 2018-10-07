Sub wall_street()

'Part2 - Create a script that will loop through all the stocks and take the following info.
'Yearly change from what the stock opened the year at to what the closing price was.
'The percent change from the what it opened the year at to what it closed.
'The total Volume of the stock
'Ticker Symbol
'You should also have conditional formatting that will highlight positive change in green and negative change in red.



  ' Set an initial variable for holding the type of stock (brand name)
  Dim Stock_Type As String

  ' Set an initial variable for holding the total per each stock (credit card brand)
  Dim Stock_Total As Double
  Stock_Total = 0
  
  ' Set initial variable for holding the total change per each stock
  Dim Stock_Change As Double
  Stock_Change = 0

  ' Keep track of the location for each stock in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock listings
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  For I = 2 To lastrow
   
   
   
    ' Check if we are still within the same stock type, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Brand name
      Stock_Type = Cells(I, 1).Value


    ' While in each row, loop through each stock change column
    For j = 3 To 6
        
        Stock_Beg = Cells(2, 3).Value

    Next j



      ' Set the Change initial Value
        'Stock_Beg = Cells(i, 3).Value
        Stock_End = Cells(I, 6).Value
        Stock_Change = Stock_End - Stock_Beg
        Stock_Percent = Format(Round(((Stock_Change / Stock_Beg)), 2), "Percent")
        'Stock_Percent = Format(Round(((Stock_Change / Stock_Beg) * 100), 2), "Percent")
        
      ' Add to the Brand Total
      Stock_Total = Stock_Total + Cells(I, 3).Value
      
      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = Stock_Type

      'Print the Stock_Change in the Summary Table
      Range("J" & Summary_Table_Row).Value = Stock_Change
      
            Set r1 = Range("J" & Summary_Table_Row)
            If r1.Value >= 0 Then r1.Interior.Color = vbGreen
            If r1.Value < 0 Then r1.Interior.Color = vbRed
      
      'Print the Stock_Change in the Summary Table
      Range("K" & Summary_Table_Row).Value = Stock_Percent

      ' Print the Brand Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Stock_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Stock_Total = Stock_Total + Cells(I, 3).Value

      'Add to the Stock_Change (- Cells(i, 6).Value
      ' Stock_Change = Stock_Change
                 

    End If
   

  Next I

  
End Sub


