# VBA-challenge
Use VBA scripting to analyze generated stock market data. 

The VBA script loops through all the stocks for each year and outputs the following information:
1. The ticker symbol.
2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
4. The total stock volume of the stock.

The bonus code returns the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

# Key VBA functionalities used in this challenge: 
1. Define and initialize variables
2. If statement to check if still the same name when looping each row by using Cells(i+1,column).Value <> Cells(i,column).Value
3. Conditional formatting of the interior color of each cell
4. If statement to check the greatest value
5. Determine the number of rows: Cells(Rows.Count, 1).End(xlUp).Row
6. For Each ws In Worksheets


