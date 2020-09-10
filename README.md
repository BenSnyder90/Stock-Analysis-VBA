# VBA-challenge
VBA Homework - The VBA of Wall Street



## Objectives

* Create a script that will loop through all the stocks for one year for each run and take the following information:
  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year
  * The total stock volume of the stock
* Conditional formatting that will highlight positive change in green and negative change in red
* Return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"
* Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once

--------------------------------------------------------------------
## Contents

* ## /VBAStocks/
  * StockSummary.VBS - Uses For Loops to move through all of the stocks in the data set, tracks the changes and prints out the totals on a Summary Table. A For Loop is also included that enables script to run through every sheet in the workbook.
  * 2014 Stock Summary.PNG - 2014 stocks worksheet after the script is run. Shows conditional formatting and summary cells
  * 2015 Stock Summary.PNG - 2015 stocks worksheet after the script is run. Shows conditional formatting and summary cells
  * 2016 Stock Summary.PNG - 2016 stocks worksheet after the script is run. Shows conditional formatting and summary cells

* ## /Challenge/
  * Challenge.VBS -Uses For Loops to look through the Summary Table and keeps track of the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume. As the loop runs through, the current stock is compared to stored stocks. If the current stock satisfies any of the conditions for the new table, it is stored instead. It also includes a For Loop that allows it to run through each table.
