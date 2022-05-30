function renameSpreadSheet() {
    var mySS = SpreadsheetApp.getActiveSpreadsheet();
    mySS.rename("2017 Avocado Prices in Portland, Seattle");
}

function duplicateSheet() {
    var mySS = SpreadsheetApp.getActiveSpreadsheet();
    var duplicateSheet = mySS.duplicateActiveSheet();
}

function duplicateAndOrganizeActiveSheet(){
    var mySS = SpreadsheetApp.getActiveSpreadsheet();
    var duplicateSheet = mySS.duplicateActiveSheet();

    // Rename the new sheet
    duplicateSheet.setName("Sheet_" + duplicateSheet.getSheetId());

    // Format the new sheet
    duplicateSheet.autoResizeColumns(1,5);
    duplicateSheet.setFrozenRows(2);

    // Move column F to Column C
    var myRange = duplicateSheet.getRange("F2:F");
    myRange.moveTo(duplicateSheet.getRange("C2"));

    // Short all the data using column C (Price information)
    myRange = duplicateSheet.getRange("A3:D55");
    myRange.sort(3);
}

/// --------------------------------------------------------

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Book-list')
        .addItem('Actualizar datos', 'updateDollarValue')
        .addToUi();
}

function updateDollarValue() {
    SpreadsheetApp.getActiveSheet().getRange("A18").setValue("Hola mundo");
}

/// --------------------------------------------------------


/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Book-list')
        .addItem('Load Book-list', 'loadBookList')
        .addSeparator()
        .addItem(
            'Separate title/author at first comma', 'splitAtFirstComma')
        .addItem(
            'Separate title/author at last "by"', 'splitAtLastBy')
        .addSeparator()
        .addItem(
            'Fill in blank titles and author cells', 'fillInTheBlanks')
        .addToUi();
}

function loadBookList() {
    var sheet = SpreadsheetApp.getActiveSheet();

    var bookSpreadSheet = SpreadsheetApp.openById("1c0GvbVUDeBmhTpq_A3vJh2xsebtLuwGwpBYqcOBqGvo");

    var bookSheet = bookSpreadSheet.getSheetByName("codelab-book-list");
    var bookRange = bookSheet.getDataRange();
    var bookListValues = bookRange.getValues();

    // Add those values to the active sheet in the current
    // spreadsheet. This overwrites any values already there.
    sheet.getRange(1,1, bookRange.getHeight(), bookRange.getWidth()).setValues(bookListValues);

    sheet.setName("Book-list");
    sheet.autoResizeColumns(1,3);
}

function splitAtFirstComma() {
    var activeRange = SpreadsheetApp.getActiveRange();
    var titleAuthorRange = activeRange.offset(0,0,activeRange.getHeight(), activeRange.getWidth()+1);

    var titleAuthorValues = titleAuthorRange.getValues();

    for (var row = 0; row < titleAuthorValues.length; row++) {
        var indexOfFirstComma = titleAuthorValues[row][0].indexOf(", ");

        if (indexOfFirstComma >= 0){
            var titlesAndAuthors = titleAuthorValues[row][0];

            titleAuthorValues[row][0] = titlesAndAuthors.slice(indexOfFirstComma + 2);
            titleAuthorValues[row][1] = titlesAndAuthors.slice(0, indexOfFirstComma);
        }
    }

    titleAuthorRange.setValues(titleAuthorValues);
}

function splitAtLastBy() {
    var activeRange = SpreadsheetApp.getActiveRange();
    var titleAuthorRange = activeRange.offset(0,0,activeRange.getHeight(), activeRange.getWidth()+1);

    var titleAuthorValues = titleAuthorRange.getValues();

    for (var row = 0; row < titleAuthorValues.length; row++) {
        var indexOfLastBy = titleAuthorValues[row][0].lastIndexOf(" by ");

        if (indexOfLastBy >= 0){
            var titlesAndAuthors = titleAuthorValues[row][0];

            titleAuthorValues[row][0] = titlesAndAuthors.slice(0, indexOfLastBy);
            titleAuthorValues[row][1] = titlesAndAuthors.slice(indexOfLastBy + 4);
        }
    }

    titleAuthorRange.setValues(titleAuthorValues);
}


/*
Note: In Apps Script, function names ending in _ (an underscore) are considered private. Other scripts can't call these
functions when present in a library, or by clients during server-client communication. These are advanced topics though,
so if you know a function is only going to be used by the current script, it's best practice to end a function's name with _.
*/
/**
 * Helper function to retrieve book data from the Open Library
 * public API.
 *
 * @param {number} ISBN - The ISBN number of the book to find.
 * @return {object} The book's data, in JSON format.
 */
function fetchBookData_(ISBN){
    // Connect to the public API.
    var url = "https://openlibrary.org/api/books?bibkeys=ISBN:" + ISBN + "&jscmd=details&format=json";
    var response = UrlFetchApp.fetch(
        url, {'muteHttpExceptions':true}
    );

    // Make request to API and get response before this point.
    var json = response.getContentText();
    var bookData = JSON.parse(json);

    // Return only the data we're interested in.
    return bookData['ISBN:' + ISBN];
}

function fillInTheBlanks() {
    // Constants that identify the index of the title, author,
    // and ISBN columns (in the 2D bookValues array below).
    var TITLE_COLUMN = 0;
    var AUTHOR_COLUMN = 1;
    var ISBN_COLUMN = 2;

    var dataRange = SpreadsheetApp.getActiveSpreadsheet().getDataRange();
    var bookValues = dataRange.getValues();

    for (var row = 0; row < bookValues.length; row++) {
        var title = bookValues[row][TITLE_COLUMN];
        var author = bookValues[row][AUTHOR_COLUMN];
        var isbn = bookValues[row][ISBN_COLUMN];

        if (isbn != "" && (title === "" || author === "")) {
            var bookData = fetchBookData_(isbn);
            if (!bookData || !bookData.details){
                continue;
            }

            if (title === "" && bookData.details.title){
                bookValues[row][TITLE_COLUMN] = bookData.details.title
            }

            if (author ==="" && bookData.details.authors && bookData.details.authors[0].name){
                bookValues[row][AUTHOR_COLUMN] = bookData.details.authors[0].name;
            }
        }
    }
    dataRange.setValues(bookValues);
}