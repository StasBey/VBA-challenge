Attribute VB_Name = "Module5"
Sub Stocks()

For Each ws In ActiveWorkbook.Worksheets
ws.Activate
 
  Dim Ticker As String
  Dim Yearly_Change As Double
  Yearly_Change = 0
  Dim Percent_Change As Double
  Percent_Total = 0
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  

  
  Range("I1") = "Ticker"
  Range("J1") = "Yearly Change"
  Range("K1") = "Percent Change"
  Range("L1") = "Total Stock Volume"
  

    
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Opening_Price = Cells(2, 3).Value
    Range("J10").Value = Opening_Price
        
    For i = 2 To LastRow
  
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker = Cells(i, 1).Value
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
      
      Closing_Price = Cells(i, 6).Value
      Range("I" & Summary_Table_Row).Value = Ticker
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      Range("J" & Summary_Table_Row).Value = (Closing_Price - Opening_Price)
      If Opening_Price <> 0 Then Range("K" & Summary_Table_Row).Value = (Closing_Price - Opening_Price) / Opening_Price Else Range("K" & Summary_Table_Row).Value = 0
      Opening_Price = Cells(i + 1, 3).Value
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      Total_Stock_Volume = 0
      
      
      Closing_Price = Cells(i, 6).Value


    Else

      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If
    
  Next i

    
    Range("K:K").Style = "Percent"
    Range("K:K").NumberFormat = "0.00%"
    


  ' Setup a counter to track cell number
    Dim y As Integer
    ' Loop through each row of the board
  For y = 2 To 100

      If Cells(y, 10) < 0 Then
       Cells(y, 10).Interior.ColorIndex = 3
      Else
       Cells(y, 10).Interior.ColorIndex = 4
       
    End If

  Next y
    
Next ws

For Each ws In Worksheets
ws.Activate

Dim rng1 As Range
  Dim rng2 As Range
  Dim dblMin As Double
  Dim dblMax As Double

  Dim Row As Double
  
  Dim Greatest_Increase As Double
  Dim Greatest_Decrease As Double
  Dim Greatest_Volume As Double
  Greatest_Increase = 0
  Greatest_Decrease = 0
  Greatest_Volume = 0
  Row = 0

  Dim max_ticker As Integer
  
  
  Range("P1") = "Ticker"
  Range("Q1") = "Value"
  Range("O2") = "Greatest % Increase"
  Range("O3") = "Greatest % Decrease"
  Range("O4") = "Greatest Total Volume"
    
  LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

  Set rng1 = Sheet1.Range("K1:K" & LastRow)
  Set rng2 = Sheet1.Range("L1:L" & LastRow)
  
  dblMin = Application.WorksheetFunction.Min(rng1)
  dblMax = Application.WorksheetFunction.max(rng1)
  dblMax1 = Application.WorksheetFunction.max(rng2)
    
  Range("Q2") = dblMax
  Range("Q3") = dblMin
  Range("Q4") = dblMax1
  
  For x = 2 To LastRow
  If Cells(x, 11) = dblMax Then Cells(2, 16) = Cells(x, 9)
  If Cells(x, 11) = dblMin Then Cells(3, 16) = Cells(x, 9)
  If Cells(x, 12) = dblMax1 Then Cells(4, 16) = Cells(x, 9)
  Next x
  
  x = 0
    
    Range("Q2:Q3").Style = "Percent"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
Next ws

End Sub

