Attribute VB_Name = "Module1"
 Sub stock_exchange()
 
'Set an initial variable
Dim ticker_symbol As String
Dim yearly_change As Double
Dim percent_change As Double
Dim Total_stock_volume As Double
' Adding two more variables to help calculating yearly and percent changes
Dim open_total As Double
Dim close_total As Double



'Populating header information
Range("H1").Value = "Summarised ticker symbol"
Range("I1").Value = "Yearly Change"
Range("J1").Value = "Percent Change"
Range("K1").Value = "Total Stock Volume"

'Assigning initial values
Total_stock_volume = 0
yearly_change = 0
percent_change = 0


  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all ticker symbols
  For i = 2 To 70926

    ' Check if we are still within the same ticker brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker symbol
      ticker_symbol = Cells(i, 1).Value

      ' Add to the Total Stock Volume
      Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
      
      ' Calculating Open and Close totals
      close_total = close_total + Cells(i, 6).Value
      open_total = open_total + Cells(i, 3).Value
      
      yearly_change = (close_total - open_total)
      ' For the percent changes I had to take care of the 0 open values as Excel returned error when dividing with zero
      If open_total <> 0 Then
            percent_change = close_total / open_total - 1
      Else: percent_change = 0
      End If
      

      ' Print the Ticker in the Summary Table
      Range("H" & Summary_Table_Row).Value = ticker_symbol
      
      'Print the Yearly Change in the Summary Table
      Range("I" & Summary_Table_Row).Value = yearly_change
      
      'Change the Percent Change format in the Summary Table
      Range("J" & Summary_Table_Row).Value = percent_change
      Range("J" & Summary_Table_Row).NumberFormat = "0.00%"

            
     ' Print the Total Stock Volume to the Summary Table
      Range("K" & Summary_Table_Row).Value = Total_stock_volume
      
    ' Adding conditional formatting
         If yearly_change > 0 Then
            Range("I" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            Range("I" & Summary_Table_Row).Interior.ColorIndex = 3
        End If


      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the values
        close_total = 0
        open_total = 0
        yearly_change = 0
        percent_change = 0
        Total_stock_volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

        close_total = close_total + Cells(i, 6).Value
        open_total = open_total + Cells(i, 3).Value
        yearly_change = (close_total - open_total)
        If open_total <> 0 Then
            percent_change = close_total / open_total - 1
        Else: percent_change = 0
        End If
        
        Total_stock_volume = Total_stock_volume + Cells(i, 7).Value

    End If

  Next i


End Sub

