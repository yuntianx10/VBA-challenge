' Create a script that loops through all the stocks for one year and outputs the following information:
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.
' Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Sub StockSummary():

' Set the header of the Summary Table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Set an initial variable for holding the ticket name
Dim Ticket_Name As String

' Set initial variables for holding the open price, close price of the year, and the yearly change
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double

Year_Open = Cells(2, 3).Value
Year_Close = 0
Yearly_Change = 0

' Set an initial variable for holding the percent change.
Dim Percent_Change As Double
Percent_Change = 0

' Set an initial variable for holding the total stock volumn of the stock.
Dim Total_Volume As Double
Total_Volume = 0

' To track the row number of ticker in a new summary table
Dim Summary_Row As Integer
Summary_Row = 2

' Determine the Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

' 1. Loop through all stocks.
For i = 2 To LastRow
    
    ' 2. Check every row and see if still the same name. If loop to i+1 row, the ticker name is not the same
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
      ' Set the Ticket name
      Ticket_Name = Cells(i, 1).Value
        
    ' Record the Close price of the year
      Year_Close = Cells(i, 6).Value
      
      ' Calculate the yearly change
      Yearly_Change = Year_Close - Year_Open
      
      ' Calculate the percent change
      Percent_Change = Yearly_Change / Year_Open
      
      ' Calculate the total volumn
      Total_Volume = Total_Volume + Cells(i, 7).Value
      
      ' Print the Ticket name, Yearly Change, Percent Change and Total Stock Volumn into the summary table
      Cells(Summary_Row, 9).Value = Ticket_Name
      Cells(Summary_Row, 10).Value = Yearly_Change
      Cells(Summary_Row, 11).Value = Percent_Change
      Cells(Summary_Row, 12).Value = Total_Volume
      
      ' Set the Conditional formatting for yearly change
            If Yearly_Change < 0 Then
                Cells(Summary_Row, 10).Interior.ColorIndex = 3
            Else
                Cells(Summary_Row, 10).Interior.ColorIndex = 4
            End If
      
      ' Add one to the summary table row
      Summary_Row = Summary_Row + 1
      
      ' Reset the year open price, yearly change, percent change and total volumn to 0
      Year_Open = Cells(i + 1, 3).Value
      Year_Close = 0
      Yearly_Change = 0
      Percent_Change = 0
      Total_Volumn = 0
    
      
    ' If the cell immediately following a row is the same brand ...
    Else
    
    ' Add to the Total_Volume
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
    End If
    
Next i
        
End Sub

