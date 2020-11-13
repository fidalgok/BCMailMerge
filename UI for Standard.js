// Compiled using ts2gas 1.6.2 (TypeScript 3.6.4)
var exports = exports || {};
var module = module || { exports: exports };
var exports = exports || {};
var module = module || { exports: exports };
function setStartingPage(page) {

    const properties = PropertiesService.getScriptProperties();
    const updateProps = properties.setProperty('startPage', page);


}
function getStartPage() {
    const properties = PropertiesService.getScriptProperties();
    const startPage = properties.getProperty('startPage');
    return startPage;
}
function startingPageforStandardMerge() {
    // send the UI to the user
    setStartingPage('startPage')
    var app = HtmlService.createTemplateFromFile("index")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(app, "Mail Merge");
}
function configureMergeOptions() {

    // update the script properties to skip the welcome page
    setStartingPage('mergeOptions')
    // send the UI to the user
    // send the UI to the user
    var app = HtmlService.createTemplateFromFile("index")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(app, "Mail Merge");

}
function configureMergeConditions() {
    setStartingPage('conditions')
    // send the UI to the user
    // send the UI to the user
    var app = HtmlService.createTemplateFromFile("index")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(app, "Mail Merge");
}
function configureCustomAttachment() {
    setStartingPage('attachments')
    // send the UI to the user
    // send the UI to the user
    var app = HtmlService.createTemplateFromFile("index")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(app, "Mail Merge");
}
function configureMergePreview() {
    setStartingPage('preview')
    // send the UI to the user
    // send the UI to the user
    var app = HtmlService.createTemplateFromFile("index")
        .evaluate()
        .setWidth(960)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(app, "Mail Merge");
}
function reRunMerge() {
    setStartingPage('confirmation')
    // send the UI to the user
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
function startStandardMerge(email, kind, sendDrafts, recipientsHeader, mergeTitle, mergeConditions, customAttachment, currentDate) {
    // stores the merge results in a variable to send back to the UI

    var result = merge(kind, email, sendDrafts, recipientsHeader, mergeTitle, mergeConditions, customAttachment, currentDate);
    return result;
}
