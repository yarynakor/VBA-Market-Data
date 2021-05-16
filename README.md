# The VBA of Wall Street

## Background

VBA scripting to analyze real stock market data. 


### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Used this while developing scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Ranscripts on this data to generate the final report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Instructions

* Create a script that will loop through all the stocks for one year for each run and take the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

### CHALLENGES

1. Return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume". 

2. adjustments to the VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
