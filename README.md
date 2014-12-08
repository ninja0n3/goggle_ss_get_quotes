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


