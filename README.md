Google Spreadsheets Get Quotes
==============================

This is a small script based on https://gist.github.com/markandey/5632923
It is used to obtain quotes from YQL and insert them directly into a google spreadsheet for further analisys/logic.

### Usage Examples
It is good for building watch lists, doing some technical/statistical analysis without having to type everything in.

### Usage

Create a new script on your google spreadsheet, add a trigger for onOpen.

You will see the "Get Quotes for Active Range" under the Add-Ons menu.

You will need to modify the script to select which fields to show; by default it shows all data obtained from Yahoo. Simply add // in front of the lines you don't want.
You can also rearrange your columns as needed

Select a sequential range of cells.
The script assumes that all ticker symbols are on a single column, the position result will be the last column on the right of your range.

####Example:

Symbols located in range A2:A10
Range selected is A2:C10

The resulting data will be displayed starting at C2

### Questions

Contact me.

### References
####Yahoo Query Language

#####Quotes Data
- https://developer.yahoo.com/yql/console/?q=show%20tables&env=store://datatables.org/alltableswithkeys#h=select+*+from+yahoo.finance.quotes+where+symbol+in+(%22YHOO%22%2C%22AAPL%22%2C%22GOOG%22%2C%22MSFT%22)

#####Historical Data
- https://developer.yahoo.com/yql/console/?q=show%20tables&env=store://datatables.org/alltableswithkeys#h=select+*+from+yahoo.finance.historicaldata+where+symbol+in(%22YHOO%22%2C%22AAPL%22)+and+startDate+%3D+%222009-09-11%22+and+endDate+%3D+%222010-03-10%22


