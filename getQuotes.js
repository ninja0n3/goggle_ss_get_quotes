/**
 * Created by Jamil on 12/8/2014.
 *
 * Based and improved from https://gist.github.com/markandey/5632923
 *
 *
 */

/********************************************
 * TODO:
 *
 * - Implement keystats data query and add columns - https://developer.yahoo.com/yql/console/?q=show%20tables&env=store://datatables.org/alltableswithkeys#h=SELECT+*+FROM+yahoo.finance.keystats+WHERE+symbol%3D'AAPL'
 *
 *
 ********************************************/

function onOpen() {
    var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or FormApp.
    menu.addItem('Get Quotes for Active Range', 'pullData');
    menu.addToUi();
}

function pullData(){

    /*************************************
     * Display Titles - Change below
     *************************************/
    var display_titles = true; // true to display titles (default), false to remove titles
    /*************************************
     * Display Titles - Change Above
     *************************************/

    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getActiveRange();

    var data = range.getValues(); //range.getValues();

    //var data = r1;

    // Query string
    var query = "";
    var lines = [];
    var j=0;

    // Loop through values - get first column only
    for (var i=0; i < data.length; i++){
        symbol = data[i][0];

        // Validate data to be a symbol
        if(symbol.match(/\w{3,5}/)){
            query += "%22"+symbol.capitalize()+"%22%2C"; //Add symbol to query
            lines[j]=i;
            j++;
        }
    }
    query = query.replace(/\%2C$/g,"");

    scol = range.getLastColumn();//cr2.getColumn();
    srow = range.getRow();//cr2.getRow();

    Logger.log(scol);
    Logger.log(srow);

    var url="https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20("+query+")&format=json&diagnostics=true&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback=";
    var response = UrlFetchApp.fetch(url);
    var json=JSON.parse(response.getContentText());

    var fields = getQuoteFields();

    // Displays the titles above the selected range
    if(display_titles){
        trow = srow - 1;
        if(trow > 0){
            dcol=0;
            for(i=0; i<fields.length; i++){
                sheet.getRange(trow, scol+dcol).setValue(fields[i]);
                sheet.getRange(trow, scol+dcol).setFontWeight("bold");
                sheet.autoResizeColumn(scol+dcol);
                dcol++;
            }
        }
    }

    // Displays the data collected from Yahoo finance YQL for quotes
    var quotes=json.query.results.quote;

    dcol = 0;
    for(i=0;i<quotes.length;i++){
        row = srow+lines[i];

        dcol = 0;
        // Loop Through selected fields
        for(j=0; j<fields.length; j++){
            var val = quotes[i][fields[j]];

            if(typeof(val) == "string" ){
                val = val.replace(/\+/g,"");
            }

            sheet.getRange(row, scol+dcol++).setValue(val);
        }


    }

    // TODO: Get Historical Data from YQL
    startdate = "2014-12-01";
    enddate = "2014-12-05";
    url="https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.historicaldata%20where%20symbol%20in%20("+query+")%20and%20startDate%20%3D%20%22"+startdate+"%22%20and%20endDate%20%3D%20%22"+enddate+"%22&format=json&diagnostics=true&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback=";
    response = UrlFetchApp.fetch(url);
    json=JSON.parse(response.getContentText());

    var hist=json.query.results.quote;
    var symbol_count = range.getValues().length;
    var record_count = json.query.count;
    var rec_per_sym = parseInt(record_count/symbol_count);

    //Logger.log("Records per symbol: "+rec_per_sym);

    fields= getHistoricalFields();

    for(i=0; i<symbol_count; i++){
        offset = i+1;

        ecol = dcol;
        for(j=rec_per_sym*i; j< rec_per_sym*offset; j++){
            row = srow+lines[i];



            // Loop Through selected fields
            for(k=0; k<fields.length; k++){

                // add titles on top
                if(i == 0 && display_titles){
                    trow = srow - 1;
                    if(trow > 0){
                        sheet.getRange(trow, scol+ecol).setValue(fields[k]+" for "+hist[j]["Date"]);
                        sheet.getRange(trow, scol+ecol).setFontWeight("bold");
                        sheet.autoResizeColumn(scol+ecol);
                    }

                }

                val = hist[j][fields[k]];
                sheet.getRange(row, scol+ecol++).setValue(val);
            }



        }




    }//

    // TODO: get Key stats data from YQL
}

function getQuoteFields(){

    i=0;
    fields = [];

    /******************************************
     *
     *  Select the desired fields below
     *  add "//" in front to disable a field as such:
     *  // fields[i++] = "symbol"; // symbol
     *
     *  Fields can be reordered as required by cutting and pasting at a higher/lower position
     *
     *  These are all the fields available on yahoo YQL.
     *  For more info: https://developer.yahoo.com/yql/console/?q=show%20tables&env=store://datatables.org/alltableswithkeys#h=select+*+from+yahoo.finance.quotes+where+symbol+in+(%22YHOO%22%2C%22AAPL%22%2C%22GOOG%22%2C%22MSFT%22)
     *
     ******************************************/

    fields[i++] = "symbol"; // symbol
    fields[i++] = "Ask"; // Ask
    fields[i++] = "AverageDailyVolume"; // AverageDailyVolume
    fields[i++] = "Bid"; // Bid
    fields[i++] = "AskRealtime"; // AskRealtime
    fields[i++] = "BidRealtime"; // BidRealtime
    fields[i++] = "BookValue"; // BookValue
    fields[i++] = "Change_PercentChange"; // Change_PercentChange
    fields[i++] = "Change"; // Change
    fields[i++] = "Commission"; // Commission
    fields[i++] = "Currency"; // Currency
    fields[i++] = "ChangeRealtime"; // ChangeRealtime
    fields[i++] = "AfterHoursChangeRealtime"; // AfterHoursChangeRealtime
    fields[i++] = "DividendShare"; // DividendShare
    fields[i++] = "LastTradeDate"; // LastTradeDate
    fields[i++] = "TradeDate"; // TradeDate
    fields[i++] = "EarningsShare"; // EarningsShare
    fields[i++] = "ErrorIndicationreturnedforsymbolchangedinvalid"; // ErrorIndicationreturnedforsymbolchangedinvalid
    fields[i++] = "EPSEstimateCurrentYear"; // EPSEstimateCurrentYear
    fields[i++] = "EPSEstimateNextYear"; // EPSEstimateNextYear
    fields[i++] = "EPSEstimateNextQuarter"; // EPSEstimateNextQuarter
    fields[i++] = "DaysLow"; // DaysLow
    fields[i++] = "DaysHigh"; // DaysHigh
    fields[i++] = "YearLow"; // YearLow
    fields[i++] = "YearHigh"; // YearHigh
    fields[i++] = "HoldingsGainPercent"; // HoldingsGainPercent
    fields[i++] = "AnnualizedGain"; // AnnualizedGain
    fields[i++] = "HoldingsGain"; // HoldingsGain
    fields[i++] = "HoldingsGainPercentRealtime"; // HoldingsGainPercentRealtime
    fields[i++] = "HoldingsGainRealtime"; // HoldingsGainRealtime
    fields[i++] = "MoreInfo"; // MoreInfo
    fields[i++] = "OrderBookRealtime"; // OrderBookRealtime
    fields[i++] = "MarketCapitalization"; // MarketCapitalization
    fields[i++] = "MarketCapRealtime"; // MarketCapRealtime
    fields[i++] = "EBITDA"; // EBITDA
    fields[i++] = "ChangeFromYearLow"; // ChangeFromYearLow
    fields[i++] = "PercentChangeFromYearLow"; // PercentChangeFromYearLow
    fields[i++] = "LastTradeRealtimeWithTime"; // LastTradeRealtimeWithTime
    fields[i++] = "ChangePercentRealtime"; // ChangePercentRealtime
    fields[i++] = "ChangeFromYearHigh"; // ChangeFromYearHigh
    fields[i++] = "PercebtChangeFromYearHigh"; // PercebtChangeFromYearHigh
    fields[i++] = "LastTradeWithTime"; // LastTradeWithTime
    fields[i++] = "LastTradePriceOnly"; // LastTradePriceOnly
    fields[i++] = "HighLimit"; // HighLimit
    fields[i++] = "LowLimit"; // LowLimit
    fields[i++] = "DaysRange"; // DaysRange
    fields[i++] = "DaysRangeRealtime"; // DaysRangeRealtime
    fields[i++] = "FiftydayMovingAverage"; // FiftydayMovingAverage
    fields[i++] = "TwoHundreddayMovingAverage"; // TwoHundreddayMovingAverage
    fields[i++] = "ChangeFromTwoHundreddayMovingAverage"; // ChangeFromTwoHundreddayMovingAverage
    fields[i++] = "PercentChangeFromTwoHundreddayMovingAverage"; // PercentChangeFromTwoHundreddayMovingAverage
    fields[i++] = "ChangeFromFiftydayMovingAverage"; // ChangeFromFiftydayMovingAverage
    fields[i++] = "PercentChangeFromFiftydayMovingAverage"; // PercentChangeFromFiftydayMovingAverage
    fields[i++] = "Name"; // Name
    fields[i++] = "Notes"; // Notes
    fields[i++] = "Open"; // Open
    fields[i++] = "PreviousClose"; // PreviousClose
    fields[i++] = "PricePaid"; // PricePaid
    fields[i++] = "ChangeinPercent"; // ChangeinPercent
    fields[i++] = "PriceSales"; // PriceSales
    fields[i++] = "PriceBook"; // PriceBook
    fields[i++] = "ExDividendDate"; // ExDividendDate
    fields[i++] = "PERatio"; // PERatio
    fields[i++] = "DividendPayDate"; // DividendPayDate
    fields[i++] = "PERatioRealtime"; // PERatioRealtime
    fields[i++] = "PEGRatio"; // PEGRatio
    fields[i++] = "PriceEPSEstimateCurrentYear"; // PriceEPSEstimateCurrentYear
    fields[i++] = "PriceEPSEstimateNextYear"; // PriceEPSEstimateNextYear
    fields[i++] = "Symbol"; // Symbol
    fields[i++] = "SharesOwned"; // SharesOwned
    fields[i++] = "ShortRatio"; // ShortRatio
    fields[i++] = "LastTradeTime"; // LastTradeTime
    fields[i++] = "TickerTrend"; // TickerTrend
    fields[i++] = "OneyrTargetPrice"; // OneyrTargetPrice
    fields[i++] = "Volume"; // Volume
    fields[i++] = "HoldingsValue"; // HoldingsValue
    fields[i++] = "HoldingsValueRealtime"; // HoldingsValueRealtime
    fields[i++] = "YearRange"; // YearRange
    fields[i++] = "DaysValueChange"; // DaysValueChange
    fields[i++] = "DaysValueChangeRealtime"; // DaysValueChangeRealtime
    fields[i++] = "StockExchange"; // StockExchange
    fields[i++] = "DividendYield"; // DividendYield
    fields[i++] = "PercentChange"; // PercentChange
    //*/
    return fields;

}

function getHistoricalFields(){

    i=0;
    fields = [];

    /******************************************
     *
     *  Select the desired fields below
     *  add "//" in front to disable a field as such:
     *  // fields[i++] = "symbol"; // symbol
     *
     *  Fields can be reordered as required by cutting and pasting at a higher/lower position
     *
     *  These are all the fields available on yahoo YQL.
     *  For more info: https://developer.yahoo.com/yql/console/?q=show%20tables&env=store://datatables.org/alltableswithkeys#h=select+*+from+yahoo.finance.historicaldata+where+symbol+%3D+%22AAPL%22+and+startDate+%3D+%222014-12-01%22+and+endDate+%3D+%222014-12-05%22
     *
     ******************************************/

    //fields[i++] = "Symbol"; // Symbol
    //fields[i++] = "Date"; // Date
    fields[i++] = "Open"; // Open
    fields[i++] = "High"; // High
    fields[i++] = "Low"; // Low
    fields[i++] = "Close"; // Close
    fields[i++] = "Volume"; // Volume
    fields[i++] = "Adj_Close"; // Adj_Close
    //*/

    return fields;


}




String.prototype.capitalize = function() {
    return this.replace(/(?:^|\s)\S/g, function(a) { return a.toUpperCase(); });
};