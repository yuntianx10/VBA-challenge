# VBA-challenge
## Project Description and Motivation
In the second week, we focused on the microsoft's event-driven programming languate - Visual Basic of Applications (VBA). Start from here, we gradually transition to use programming tools for data analysis. In this VBA-challenge, we aim to use VBA scripting to analyze generated stock market data. 

## Tasks
The VBA script loops through all the stocks for each year and outputs the following information:
1. The ticker symbol.
2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
4. The total stock volume of the stock.

The bonus code returns the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.

## Results
![alt=""](https://github.com/yuntianx10/VBA-challenge/blob/main/Results%20Screenshots/Results_2018.jpg "Summary Table 2018")
![alt=""](https://github.com/yuntianx10/VBA-challenge/blob/main/Results%20Screenshots/Results_2019.jpg "Summary Table 2019")
![alt=""](https://github.com/yuntianx10/VBA-challenge/blob/main/Results%20Screenshots/Results_2020.jpg "Summary Table 2020")


## Key VBA Functionalities Used
1. Define and initialize variables
2. If statement to check if still the same name when looping each row by using Cells(i+1,column).Value <> Cells(i,column).Value
3. Conditional formatting of the interior color of each cell using Cells(row,column).Interior.ColorIndex = #
4. If statement to determine the greatest value in a certain range of cells
5. Determine the number of rows: Cells(Rows.Count, 1).End(xlUp).Row
6. For Each ws In Worksheets


