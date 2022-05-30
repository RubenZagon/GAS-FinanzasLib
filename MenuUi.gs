function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Book-list')
        .addItem('Fetch de datos', 'fetchAllData')
        .addToUi();
}