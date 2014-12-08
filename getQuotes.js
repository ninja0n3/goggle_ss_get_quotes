/**
 * Created by Jamil on 12/8/2014.
 *
 * Based and improved from https://gist.github.com/markandey/5632923
 *
 *
 */
function onOpen() {
    var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or FormApp.
    menu.addItem('Get Quotes for Active Range', 'pullData');
    menu.addToUi();
}

function pullData(){

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

    var url="https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20("+query+")&format=json&diagnostics=true&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback="
    var response = UrlFetchApp.fetch(url);
    var json=JSON.parse(response.getContentText());
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Displays the data collected from Yahoo finance YQL
    var quotes=json.query.results.quote;
    for(var i=0;i<quotes.length;i++){
        row = srow+lines[i];
        dcol = 0;
        sheet.getRange(row,scol).setValue(quotes[i].Name);
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LastTradePriceOnly); dcol++;
        sheet.getRange(row, scol+dcol).setValue(quotes[i].symbol); dcol++; // symbol
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Ask); dcol++; // Ask
        sheet.getRange(row, scol+dcol).setValue(quotes[i].AverageDailyVolume); dcol++; // AverageDailyVolume
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Bid); dcol++; // Bid
        sheet.getRange(row, scol+dcol).setValue(quotes[i].AskRealtime); dcol++; // AskRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].BidRealtime); dcol++; // BidRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].BookValue); dcol++; // BookValue
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Change_PercentChange); dcol++; // Change_PercentChange
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Change); dcol++; // Change
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Commission); dcol++; // Commission
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Currency); dcol++; // Currency
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangeRealtime); dcol++; // ChangeRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].AfterHoursChangeRealtime); dcol++; // AfterHoursChangeRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DividendShare); dcol++; // DividendShare
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LastTradeDate); dcol++; // LastTradeDate
        sheet.getRange(row, scol+dcol).setValue(quotes[i].TradeDate); dcol++; // TradeDate
        sheet.getRange(row, scol+dcol).setValue(quotes[i].EarningsShare); dcol++; // EarningsShare
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ErrorIndicationreturnedforsymbolchangedinvalid); dcol++; // ErrorIndicationreturnedforsymbolchangedinvalid
        sheet.getRange(row, scol+dcol).setValue(quotes[i].EPSEstimateCurrentYear); dcol++; // EPSEstimateCurrentYear
        sheet.getRange(row, scol+dcol).setValue(quotes[i].EPSEstimateNextYear); dcol++; // EPSEstimateNextYear
        sheet.getRange(row, scol+dcol).setValue(quotes[i].EPSEstimateNextQuarter); dcol++; // EPSEstimateNextQuarter
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DaysLow); dcol++; // DaysLow
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DaysHigh); dcol++; // DaysHigh
        sheet.getRange(row, scol+dcol).setValue(quotes[i].YearLow); dcol++; // YearLow
        sheet.getRange(row, scol+dcol).setValue(quotes[i].YearHigh); dcol++; // YearHigh
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HoldingsGainPercent); dcol++; // HoldingsGainPercent
        sheet.getRange(row, scol+dcol).setValue(quotes[i].AnnualizedGain); dcol++; // AnnualizedGain
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HoldingsGain); dcol++; // HoldingsGain
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HoldingsGainPercentRealtime); dcol++; // HoldingsGainPercentRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HoldingsGainRealtime); dcol++; // HoldingsGainRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].MoreInfo); dcol++; // MoreInfo
        sheet.getRange(row, scol+dcol).setValue(quotes[i].OrderBookRealtime); dcol++; // OrderBookRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].MarketCapitalization); dcol++; // MarketCapitalization
        sheet.getRange(row, scol+dcol).setValue(quotes[i].MarketCapRealtime); dcol++; // MarketCapRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].EBITDA); dcol++; // EBITDA
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangeFromYearLow); dcol++; // ChangeFromYearLow
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PercentChangeFromYearLow); dcol++; // PercentChangeFromYearLow
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LastTradeRealtimeWithTime); dcol++; // LastTradeRealtimeWithTime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangePercentRealtime); dcol++; // ChangePercentRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangeFromYearHigh); dcol++; // ChangeFromYearHigh
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PercebtChangeFromYearHigh); dcol++; // PercebtChangeFromYearHigh
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LastTradeWithTime); dcol++; // LastTradeWithTime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LastTradePriceOnly); dcol++; // LastTradePriceOnly
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HighLimit); dcol++; // HighLimit
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LowLimit); dcol++; // LowLimit
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DaysRange); dcol++; // DaysRange
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DaysRangeRealtime); dcol++; // DaysRangeRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].FiftydayMovingAverage); dcol++; // FiftydayMovingAverage
        sheet.getRange(row, scol+dcol).setValue(quotes[i].TwoHundreddayMovingAverage); dcol++; // TwoHundreddayMovingAverage
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangeFromTwoHundreddayMovingAverage); dcol++; // ChangeFromTwoHundreddayMovingAverage
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PercentChangeFromTwoHundreddayMovingAverage); dcol++; // PercentChangeFromTwoHundreddayMovingAverage
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangeFromFiftydayMovingAverage); dcol++; // ChangeFromFiftydayMovingAverage
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PercentChangeFromFiftydayMovingAverage); dcol++; // PercentChangeFromFiftydayMovingAverage
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Name); dcol++; // Name
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Notes); dcol++; // Notes
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Open); dcol++; // Open
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PreviousClose); dcol++; // PreviousClose
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PricePaid); dcol++; // PricePaid
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ChangeinPercent); dcol++; // ChangeinPercent
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PriceSales); dcol++; // PriceSales
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PriceBook); dcol++; // PriceBook
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ExDividendDate); dcol++; // ExDividendDate
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PERatio); dcol++; // PERatio
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DividendPayDate); dcol++; // DividendPayDate
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PERatioRealtime); dcol++; // PERatioRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PEGRatio); dcol++; // PEGRatio
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PriceEPSEstimateCurrentYear); dcol++; // PriceEPSEstimateCurrentYear
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PriceEPSEstimateNextYear); dcol++; // PriceEPSEstimateNextYear
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Symbol); dcol++; // Symbol
        sheet.getRange(row, scol+dcol).setValue(quotes[i].SharesOwned); dcol++; // SharesOwned
        sheet.getRange(row, scol+dcol).setValue(quotes[i].ShortRatio); dcol++; // ShortRatio
        sheet.getRange(row, scol+dcol).setValue(quotes[i].LastTradeTime); dcol++; // LastTradeTime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].TickerTrend); dcol++; // TickerTrend
        sheet.getRange(row, scol+dcol).setValue(quotes[i].OneyrTargetPrice); dcol++; // OneyrTargetPrice
        sheet.getRange(row, scol+dcol).setValue(quotes[i].Volume); dcol++; // Volume
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HoldingsValue); dcol++; // HoldingsValue
        sheet.getRange(row, scol+dcol).setValue(quotes[i].HoldingsValueRealtime); dcol++; // HoldingsValueRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].YearRange); dcol++; // YearRange
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DaysValueChange); dcol++; // DaysValueChange
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DaysValueChangeRealtime); dcol++; // DaysValueChangeRealtime
        sheet.getRange(row, scol+dcol).setValue(quotes[i].StockExchange); dcol++; // StockExchange
        sheet.getRange(row, scol+dcol).setValue(quotes[i].DividendYield); dcol++; // DividendYield
        sheet.getRange(row, scol+dcol).setValue(quotes[i].PercentChange); dcol++; // PercentChange
    }//*/
}

String.prototype.capitalize = function() {
    return this.replace(/(?:^|\s)\S/g, function(a) { return a.toUpperCase(); });
};
