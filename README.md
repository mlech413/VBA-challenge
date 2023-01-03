# VBA-challenge
## Stock analysis for years 2018  - 2020

The VBA routine StockData was run first in the test file alphabetical_testing.xlsm, and then in the actual data Multiple_year_stock_data.xlsm (which was too large to upload to github).

The single script process all worksheets, cycling through and processing each yearly page. Within each year, each stock is read and calculated, and an annual summary is written which contains:
- Ticker
- Yearly Change (amount)
- Percent Change
- Total Stock Volume
- 
The Yearly Change and Percent Change are formated within the routine to display as a green background for an increase, red for a decrease, and white for no change.

In addition, three rows are written for each year, containing the stock with the overall:
- Greatest % Increase
- Greatest % Decrease
- Greatest Total Volume
- 
Files included are screenshots from each year, the exported StockData VBA script, and the test file alphabetical_testing.xlsm (which also includes the VBA script).
