type Added = {
  text: string;
  to: string[];
};
type Existing = {
  from: string;
  to: string[];
}

export type ParsedData = {
  added: Record<string, Added>;
  existing: Record<string, Existing>;
}

function getParsedData(): ParsedData {
  const sheets = SpreadsheetApp.getActive().getSheets();
  const trackerSheet = sheets[1];
  const trackerSheetMaxRows = trackerSheet.getMaxRows();

  const trackerSheetRawData = trackerSheet.getRange(`A5:G${trackerSheetMaxRows}`).getDisplayValues();

  return trackerSheetRawData.reduce((prev, curr, index) => {
    const isExactMatch = curr[2] !== "-" && curr[2] !== "not found";

    if(!isExactMatch) return prev;
 
    const currentRow = index + 5;
    const exactMatchCell = curr[2].split(",").map((el) => "H" + el.trim());

    const isAlreadyTranslated = curr[6] !== "-";
    const alreadyTranslatedCell = "H" + curr[6]; 

    const isTranslationAdded = !!curr[1].length;
    const addedTranslation = curr[1];

    if(isTranslationAdded) {
      if(prev.added) {
        return {...prev, added: {...prev.added, [currentRow]: {to: exactMatchCell, text: addedTranslation}}}
      } else {
        return {...prev, added: {[currentRow]: {to: exactMatchCell, text: addedTranslation}}}
      }
    }
    if(isAlreadyTranslated) {
      if(prev.existing) {
        return {...prev, existing: {...prev.existing, [currentRow]: {to: exactMatchCell, from: alreadyTranslatedCell}}}
      } else {
        return {...prev, existing: {[currentRow]: {to: exactMatchCell, from: alreadyTranslatedCell}}}
      }
    }

    return prev;
  }, {} as ParsedData)
}

function addAlreadyAddedTranslations() {
  const destination = SpreadsheetApp.getActive().getSheets()[0];

  const data = getParsedData().existing;

  if(data) {
    Object.keys(data).forEach((key) => {
      const fromValue = destination.getRange(data[key].from).getValue();
      data[key].to.forEach((el) => destination.getRange(el).setValue(fromValue));
    })
  }

}

function addNewlyAddedTranslations() {
  const destination = SpreadsheetApp.getActive().getSheets()[0];

  const data = getParsedData().added;

  if(data) {
    Object.keys(data).forEach((key) => {
      data[key].to.forEach((el) => destination.getRange(el).setValue(data[key].text));
    })
  }
}

function handleAddNewTranslations() {
  handleErrors();
  const changed = getNewChangesToShow(getParsedData().added);
  const ui = SpreadsheetApp.getUi();

  if(!changed.total) return ui.alert('Nothing to change');

  const modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
  modal.description = "Are you sure to add new translations?";

  modal.total = changed.total;
  modal.translations = changed.changes;
  modal.shouldShowDeleteRowsCheckBox = true;
  modal.modalType = "new";
  const evaluatedModal = modal.evaluate();
  ui.showModalDialog(evaluatedModal, 'Add new translations')
}

function handleAddExistingTranslations() {
  handleErrors();
  const changed = getExistingChangesToShow(getParsedData().existing);
  const ui = SpreadsheetApp.getUi();

  if(!changed.total) return ui.alert('Nothing to change');

  const modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
  modal.description = "Are you sure to add existing translations?";

  modal.total = changed.total;
  modal.translations = changed.changes;
  modal.shouldShowDeleteRowsCheckBox = true;
  modal.modalType = "existing";
  const evaluatedModal = modal.evaluate();
  ui.showModalDialog(evaluatedModal, 'Add existing translations')
}

function handleDeleteRows() {
  handleErrors();
  const rowsToDelete = getRowsToDelete();

  const ui = SpreadsheetApp.getUi();
  if(!rowsToDelete.length) return ui.alert('Nothing to delete');

  const modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
  modal.description = "Are you sure to delete rows?";

  modal.total = rowsToDelete.length;
  modal.translations = rowsToDelete;
  modal.shouldShowDeleteRowsCheckBox = false;
  modal.modalType = "deleteRows";
  const evaluatedModal = modal.evaluate();
  ui.showModalDialog(evaluatedModal, 'Delete rows')
}

function updateTranslationsStatus() {
  const maxRows = getMainSheet().getMaxRows();
  const data = getMainSheet().getRange(`G5:L${maxRows}`).getDisplayValues();
  const statuses = getMainSheet().getRange(`A5:A${maxRows}`).getDisplayValues();

  const result = data.reduce((prev,curr, index) => {
    const currentRow = index + 5;
    const en = curr[0].trim().normalize();
    const ar = curr[1].trim().normalize();
    const allEqual = curr.every(el => el === curr[0]);
    const allToBeDeleted = curr.every(el => el === '-');
    const currWithoutAr = curr.filter((_, index) => index !== 1);
    const equalWithoutAr = currWithoutAr.every(el => el === currWithoutAr[0]);
    const isDeleteStatus = statuses[index].includes('delete');
    const isTranslationMissing = curr.some(el => el === '');
    
    if(curr[0] === '' || isDeleteStatus) return prev;
    if (isTranslationMissing) return {...prev, [currentRow]: 'missing'}
    if(allToBeDeleted) return {...prev, [currentRow]: 'delete'}
    if(allEqual) return {...prev, [currentRow]: 'todo'}
    if(en !== ar && equalWithoutAr) return {...prev, [currentRow]: 'only-arabic'}

    return {...prev, [currentRow]: '-'}
  }, {})

  Object.keys(result).forEach((el) => {
    getMainSheet().getRange(`A${el}`).setValue(result[el]);
  })
}

function updateTranslationsForLoop() {
  const maxRows = getMainSheet().getMaxRows();
  let data = {};

  for(let i = 5; i <= maxRows; i++) {
    const currentTranslations = getMainSheet().getRange(`G${i}:L${i}`).getDisplayValues().pop();
    const currentStatus = getMainSheet().getRange(`A${i}`).getDisplayValues().pop();
    
    const en = currentTranslations[0].trim().normalize();
    const ar = currentTranslations[1].trim().normalize();
    const allEqual = currentTranslations.every(el => el === currentTranslations[0]);
    const allToBeDeleted = currentTranslations.every(el => el === '-');
    const currWithoutAr = currentTranslations.filter((_, index) => index !== 1);
    const equalWithoutAr = currWithoutAr.every(el => el === currWithoutAr[0]);
    const isDeleteStatus = currentStatus.includes('delete');
    const isTranslationMissing = currentTranslations.some(el => el === '');

    //if(currentTranslations[0] === '' || isDeleteStatus) return data;
    if (isTranslationMissing) getMainSheet().getRange(`A${i}`).setValue('missing');
    if(allToBeDeleted) getMainSheet().getRange(`A${i}`).setValue('delete');
    if(allEqual) getMainSheet().getRange(`A${i}`).setValue('todo');
    if(en !== ar && equalWithoutAr) getMainSheet().getRange(`A${i}`).setValue('only-arabic');
  };
}

function handleUpdateTranslationsStatus() {
  const ui = SpreadsheetApp.getUi();
  if (SpreadsheetApp.getActiveSheet().getSheetName() !== SETTINGS.sheets.main) {
    throw ui.alert('Invalid sheet!\r\n Script works only on: ' + SETTINGS.sheets.main);
};

  const modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
  modal.description = "Are you sure to update translations status?";

  modal.total = 0;
  modal.translations = [];
  modal.shouldShowDeleteRowsCheckBox = false;
  modal.modalType = "updateTranslationsStatus";
  const evaluatedModal = modal.evaluate();
  ui.showModalDialog(evaluatedModal, 'Update translations status')
}


function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu("Dev")
       .addSubMenu(SpreadsheetApp.getUi().createMenu('Main')
          .addItem('Update translations status', 'handleUpdateTranslationsStatus'))
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Tracker')
          .addItem('Add Existing Translations', 'handleAddExistingTranslations')
          .addItem('Add New Translations', 'handleAddNewTranslations')
          .addItem('Delete "not found" rows', 'handleDeleteRows'))
      .addToUi();
};
