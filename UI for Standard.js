// Compiled using ts2gas 1.6.2 (TypeScript 3.6.4)
var exports = exports || {};
var module = module || { exports: exports };
var exports = exports || {};
var module = module || { exports: exports };
function startingPageforStandardMerge() {
    // send the UI to the user
    var app = HtmlService.createTemplateFromFile("index")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(app, "Mail Merge");
}
function helpPageForMerge() {
    // send the UI to the user
    var helpApp = HtmlService.createTemplateFromFile("help")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(helpApp, "Get Help");
}
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function checkEmailColumn() {
    var dataSheet = ss.getActiveSheet();
    var lastColumn = dataSheet.getLastColumn();
    var headers = dataSheet.getRange(1, 1, 1, lastColumn).getValues();
    return headers;
}
function startStandardMerge(email, kind, sendDrafts, recipientsHeader, mergeTitle, mergeConditions, customAttachment) {
    // stores the merge results in a variable to send back to the UI
    
    var result = merge(kind, email, sendDrafts, recipientsHeader, mergeTitle, mergeConditions, customAttachment);
    return result;
}
