function onOpen() {
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('Quick formats')
        .addItem('Format row header', 'formatRowHeader')
        .addItem('Format column header', 'formatColumnHeader')
        .addItem('Format dataset', 'formatDataset')
        .addSeparator()
        .addSubMenu(ui.createMenu('Create character sheet')
            .addItem('Episode IV', 'createPeopleSheetIV')
            .addItem('Episode V', 'createPeopleSheetV')
            .addItem('Episode VI', 'createPeopleSheetVI')
        )
        .addToUi();
}

function formatRowHeader() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());

    headerRange
        .setFontWeight('bold')
        .setFontColor('#ffffff')
        .setBackground('#007272')
        .setBorder(
            true, true, true, true, null, null,
            null,
            SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function formatColumnHeader() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var numberRows = sheet.getDataRange().getLastColumn() - 1;
    var columnHeaderRange = sheet.getRange(2,1,numberRows,1);

    columnHeaderRange
        .setFontWeight('bold')
        .setFontStyle('italic')
        .setBorder(
            true, true, true, true, null, null,
            null,
            SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    hyperlinkColumnHeaders_(columnHeaderRange, numberRows);
}

function formatDataset() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var fullDataRange = sheet.getDataRange();

    // Apply row banding to the data, excluding the header
    // row and column. Only apply the banding if the range
    // doesn't already have banding set.
    var noHeadersRange = fullDataRange.offset(
        1, 1,
        fullDataRange.getNumRows() - 1,
        fullDataRange.getNumColumns() - 1
    );

    if (!noHeadersRange.getBandings()[0]) {
        // The range doesn't already have banding, so it's
        // safe to apply it.
        noHeadersRange.applyRowBanding(
            SpreadsheetApp.BandingTheme.LIGHT_GREY,
            false, false
        );
    }

    formatDates_(columnIndexOf_('release_date'));

    fullDataRange.setBorder(
        true, true, true, true, null, null,
        null,
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.autoResizeColumns(1, fullDataRange.getNumColumns());
    sheet.autoResizeRows(1, fullDataRange.getNumRows());
}

/**
 * Helper function that hyperlinks the column header with the
 * 'url' column contents. The function then removes the column.
 *
 * @param {object} headerRange The range of the column header
 *   to update.
 * @param {number} numberRows The size of the column header.
 */
function hyperlinkColumnHeaders_(headerRange, numberRows) {
    var headerColumnIndex = 1;
    var urlColumnIndex = columnIndexOf_('url');

    if (urlColumnIndex === -1) {
        return;
    }

    var urlRange = headerRange.offset(0, urlColumnIndex - headerColumnIndex);
    var headerValues = headerRange.getValues();
    var urlValues = urlRange.getValues();

    for (var row = 0; row <= numberRows; row++) {
        let column = 0;
        headerValues[row][column] = '=HYPERLINK("' + urlValues[row] + '","' + headerValues[row] + '")';
    }
    headerRange.setValues(headerValues);

    SpreadsheetApp.getActiveSheet().deleteColumn(urlColumnIndex);
}

/**
 * Helper method that applies a
 * "Month Day, Year (Day of Week)" date format to the
 * indicated column in the active sheet.
 *
 * @param {number} columnIndexOf The index of the column
 *   to format.
 */
function formatDates_(columnIndexOf) {
    if (columnIndexOf < 0){
        return;
    }
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(2, columnIndexOf, sheet.getLastRow() - 1, 1)
        .setNumberFormat("mmmm dd, yyyy (dddd)");
}

/**
 * Helper function that goes through the headers of all columns
 * and returns the index of the column with the specified name
 * in row 1. If a column with that name does not exist,
 * this function returns -1. If multiple columns have the same
 * name in row 1, the index of the first one discovered is
 * returned.
 *
 * @param {string} columnName The name to find in the column
 *   headers.
 * @return The index of that column in the active sheet,
 *   or -1 if the name isn't found.
 */
function columnIndexOf_(columnName) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var columnHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var columnNames = columnHeaders.getValues();

    for (var column = 0; column <= columnNames[0].length; column++) {
        if (columnNames[0][column-1] === columnName){
            return column
        }
    }

    return -1;
}
