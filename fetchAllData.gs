function fetchAllData(){
    updateDollarPrice_()
}

function fetchDollarPrice_() {
    var url = "https://cdn.jsdelivr.net/gh/fawazahmed0/currency-api@1/latest/currencies/eur/usd.json";
    var response = UrlFetchApp.fetch(
        url, {'muteHttpExceptions':true}
    );

    var json = JSON.parse(response.getContentText());
    return json["usd"];
}

function updateDollarPrice_() {
    var dollarPrice = fetchDollarPrice_();
    var activeSheet = SpreadsheetApp.getActiveSheet();
    activeSheet.getRange("'Configuración'!B10").setValue(dollarPrice);
}
