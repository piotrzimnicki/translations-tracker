// const VALID_SHEET_NAME = 'Tracker';

// function isActiveSheetValid () {
//   const activeSheetName = SpreadsheetApp.getActiveSheet().getSheetName();
//   return activeSheetName === VALID_SHEET_NAME;
// }

// function isFilterActive() {
//   const sheets = SpreadsheetApp.getActive().getSheets();
//   const trackerSheet = sheets[1];
//   const filterC1 = trackerSheet.getRange("C1").getDisplayValue();
//   const filterD1 = trackerSheet.getRange("D1").getDisplayValue();

//   return filterC1 !== "Row" || filterD1 !== "Row"
// }

// function getParsedData() {
//   const sheets = SpreadsheetApp.getActive().getSheets();

//   const trackerSheet = sheets[1];
//   const trackerSheetMaxRows = trackerSheet.getMaxRows();

//   const trackerSheetRawData = trackerSheet.getRange(`A5:G${trackerSheetMaxRows}`).getDisplayValues();

//   return trackerSheetRawData.reduce((prev, curr, index) => {
//     const isExactMatch = curr[2] !== "-" && curr[2] !== "not found";

//     if(!isExactMatch) return prev;
 
//     const currentRow = index + 5;
//     const exactMatchCell = curr[2].split(",").map((el) => "H" + el.trim());

//     const isAlreadyTranslated = curr[6] !== "-";
//     const alreadyTranslatedCell = "H" + curr[6]; 

//     const isTranslationAdded = !!curr[1].length;
//     const addedTranslation = curr[1];

//     if(isTranslationAdded) {
//       if(prev.added) {
//         return {...prev, added: {...prev.added, [currentRow]: {to: exactMatchCell, text: addedTranslation}}}
//       } else {
//         return {...prev, added: {[currentRow]: {to: exactMatchCell, text: addedTranslation}}}
//       }
//     }
//     if(isAlreadyTranslated) {
//       if(prev.existing) {
//         return {...prev, existing: {...prev.existing, [currentRow]: {to: exactMatchCell, from: alreadyTranslatedCell}}}
//       } else {
//         return {...prev, existing: {[currentRow]: {to: exactMatchCell, from: alreadyTranslatedCell}}}
//       }
//     }

//     return prev;
//   }, {})
// }

// function addAlreadyAddedTranslations() {
//   const ui = SpreadsheetApp.getUi();
//   if(!isActiveSheetValid()) return ui.alert('Invalid sheet!\r\n Script works only on: ' + VALID_SHEET_NAME);
//   if(isFilterActive()) return ui.alert('Please set filter (C1/D1) to "Row"');

//   const destination = SpreadsheetApp.getActive().getSheets()[0];

//   const data = getParsedData().existing;
//   const dataToShow = data ? Object.keys(data).map((el) => `from ${data[el].from} to ${data[el].to}`) : null;

//   if(data) {
//     Object.keys(data).forEach((key) => {
//       const fromValue = destination.getRange(data[key].from).getValue();
//       data[key].to.forEach((el) => destination.getRange(el).setValue(fromValue));
//     })
//     return ui.alert("Changes:\r\n" + dataToShow.join("\r\n"))
//   } else {
//     return ui.alert('Nothing has changed')
//   }

// }

// function addNewlyAddedTranslations() {
//   const ui = SpreadsheetApp.getUi();
//   if(!isActiveSheetValid()) return ui.alert('Invalid sheet!\r\n Script works only on: ' + VALID_SHEET_NAME);
//   if(isFilterActive()) return ui.alert('Please set filter (C1/D1) to "Row"');

//   const destination = SpreadsheetApp.getActive().getSheets()[0];

//   const data = getParsedData().added;
//   const dataToShow = data ? Object.keys(data).map((el) => `${data[el].text} to ${data[el].to}`) : null;

//   if(data) {
//     Object.keys(data).forEach((key) => {
//       data[key].to.forEach((el) => destination.getRange(el).setValue(data[key].text));
//     })
//     return ui.alert("Changes:\r\n" + dataToShow.join("\r\n"))
//   } else {
//     return ui.alert('Nothing has changed')
//   }
// }



// function onOpen() {
//   SpreadsheetApp.getUi()
//       .createMenu("Dev")
//       .addItem('Add Existing Translations', 'addAlreadyAddedTranslations')
//       .addItem('Add New Translations', 'addNewlyAddedTranslations')
//       .addToUi();
// }
