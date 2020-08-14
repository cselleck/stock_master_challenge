Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Worksheets(xSh.Name).Activate
        Call stock_master
    Next
    Application.ScreenUpdating = True
End Sub


Sub stock_master()


    ' Set variables for the values needed
     Dim Stock_Name As String
     Dim StockVolume As Double
     Dim MaxVolume As Double
     Dim Minchange As Double
     Dim Maxchange As Double
     Dim PriceChange As Double
     Dim Percentchange As Integer
     Dim Lowprice As Double
     Dim HighPrice As Double
     Dim EndPrice As Double
  
     ' Set Headers in Summary Table
     Cells(1, 9).Value = "Ticker"
     Cells(1, 10).Value = "Total Stock Volume"
     Cells(1, 11).Value = "Yearly Change"
     Cells(1, 12).Value = "Yearly Percent Change"
     Cells(1, 15).Value = "Ticker"
     Cells(1, 16).Value = "Value"
     Cells(2, 14).Value = "Greatest % Increase"
     Cells(3, 14).Value = "Greatest % Decrease"
     Cells(4, 14).Value = "Greatest Volume"
     

  
    'Find the last row
  
     LastRow = Cells(Rows.Count, 1).End(xlUp).Row

     ' Set an initial variable values
  
     Lowprice = Cells(2, 6).Value
     HighPrice = Cells(2, 6).Value
     StockVolume = 0
     PriceChange = 0
     Percentchange = 0
     StartPrice = Cells(2, 3).Value

     ' Keep track of the location for the summary chart
     Dim Summary_Table_Row As Integer
     Summary_Table_Row = 2

     ' Loop through all stock entries
     For i = 2 To LastRow

     ' Check if we are still within the same stock ticker, if it is not...
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

         ' Set the Stock name
         Stock_Name = Cells(i, 1).Value

         ' Add to the volume total
         StockVolume = StockVolume + Cells(i, 7).Value
      
         ' Set EndPrice value
         EndPrice = Cells(i, 6).Value

         ' Print Stock name on summary
         Range("I" & Summary_Table_Row).Value = Stock_Name

         ' Print the StockVolume to the Summary Table
         Range("J" & Summary_Table_Row).Value = StockVolume
      
         ' Print the price change to the summary table
         Range("K" & Summary_Table_Row).Value = EndPrice - StartPrice
      
         ' Print the Percent change to the table
         If StartPrice = 0 Then
            Range("L" & Summary_Table_Row).Value = FormatPercent(EndPrice, 2, vbFalse, vbFalse, vbFalse)
         
         Else: Range("L" & Summary_Table_Row).Value = FormatPercent(((EndPrice - StartPrice) / StartPrice), 2, vbFalse, vbFalse, vbFalse)
         End If
         
         ' Set the color of the cell
         If Range("K" & Summary_Table_Row).Value >= 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
         Else: Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
         End If
        
         ' Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
      
         ' Reset the StockVolume Total
         StockVolume = 0
      
      
         ' Get the next StartPrice
         StartPrice = Cells(i + 1, 3).Value
      
         ' Get the next EndPrice
         EndPrice = Cells(i + 1, 6).Value
      

    ' If the cell immediately following a row is the same brand...
     Else

        ' Add to the Brand Total
         StockVolume = StockVolume + Cells(i, 7).Value
      
         'Check to see if HighPrice or LowPrice need to be changed
         If Cells(i, 6).Value < Lowprice Then
             Lowprice = Cells(i, 6).Value
         End If
         If Cells(i, 6).Value > HighPrice Then
             HighPrice = Cells(i, 6).Value
         End If
    
      
     End If

  Next i
  
  '**** Challenge questions ********
  
  
  'Find length of summary chart
  LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row
  
  'Set variables
  Minchange = Cells(2, 12).Value
  Maxchange = Cells(2, 12).Value
  MaxVolume = Cells(2, 10).Value
  
  'Loop through summary chart
  For k = 2 To LastRow2
    
    'check for MinChange Value
    If Cells(k, 12).Value < Minchange Then
        Minchange = Cells(k, 12).Value
        MinIndex = k
    End If
      
    'Check for MaxChange value
    If Cells(k, 12).Value > Maxchange Then
        Maxchange = Cells(k, 12).Value
        MaxIndex = k
    End If
    
    'Check for MaxVolume value
    If Cells(k, 10).Value > MaxVolume Then
        MaxVolume = Cells(k, 10).Value
        VolIndex = k
    End If
     
  Next k
  
  'Print Tickers to Summary Table

  Cells(2, 15).Value = Cells(MaxIndex, 9).Value
  Cells(3, 15).Value = Cells(MinIndex, 9).Value
  Cells(4, 15).Value = Cells(VolIndex, 9).Value
  
  'Print values to Summary table
  Cells(2, 16).Value = FormatPercent(Maxchange, 2, vbFalse, vbFalse, vbFalse)
  Cells(3, 16).Value = FormatPercent(Minchange, 2, vbFalse, vbFalse, vbFalse)
  Cells(4, 16).Value = MaxVolume
        
  'reset variables
  Minchange = 0
  Maxchange = 0
  MinIndex = 0
  
  
End Sub



