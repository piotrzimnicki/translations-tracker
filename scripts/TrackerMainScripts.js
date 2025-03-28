// Compiled using tracker 1.0.0 (TypeScript 4.9.5)
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var exports = exports || {};
var module = module || { exports: exports };
function getParsedData() {
    var sheets = SpreadsheetApp.getActive().getSheets();
    var trackerSheet = sheets[1];
    var trackerSheetMaxRows = trackerSheet.getMaxRows();
    var trackerSheetRawData = trackerSheet.getRange("A5:G".concat(trackerSheetMaxRows)).getDisplayValues();
    return trackerSheetRawData.reduce(function (prev, curr, index) {
        var _a, _b, _c, _d;
        var isExactMatch = curr[2] !== "-" && curr[2] !== "not found";
        if (!isExactMatch)
            return prev;
        var currentRow = index + 5;
        var exactMatchCell = curr[2].split(",").map(function (el) { return "H" + el.trim(); });
        var isAlreadyTranslated = curr[6] !== "-";
        var alreadyTranslatedCell = "H" + curr[6];
        var isTranslationAdded = !!curr[1].length;
        var addedTranslation = curr[1];
        if (isTranslationAdded) {
            if (prev.added) {
                return __assign(__assign({}, prev), { added: __assign(__assign({}, prev.added), (_a = {}, _a[currentRow] = { to: exactMatchCell, text: addedTranslation }, _a)) });
            }
            else {
                return __assign(__assign({}, prev), { added: (_b = {}, _b[currentRow] = { to: exactMatchCell, text: addedTranslation }, _b) });
            }
        }
        if (isAlreadyTranslated) {
            if (prev.existing) {
                return __assign(__assign({}, prev), { existing: __assign(__assign({}, prev.existing), (_c = {}, _c[currentRow] = { to: exactMatchCell, from: alreadyTranslatedCell }, _c)) });
            }
            else {
                return __assign(__assign({}, prev), { existing: (_d = {}, _d[currentRow] = { to: exactMatchCell, from: alreadyTranslatedCell }, _d) });
            }
        }
        return prev;
    }, {});
}
function addAlreadyAddedTranslations() {
    var destination = SpreadsheetApp.getActive().getSheets()[0];
    var data = getParsedData().existing;
    if (data) {
        Object.keys(data).forEach(function (key) {
            var fromValue = destination.getRange(data[key].from).getValue();
            data[key].to.forEach(function (el) { return destination.getRange(el).setValue(fromValue); });
        });
    }
}
function addNewlyAddedTranslations() {
    var destination = SpreadsheetApp.getActive().getSheets()[0];
    var data = getParsedData().added;
    if (data) {
        Object.keys(data).forEach(function (key) {
            data[key].to.forEach(function (el) { return destination.getRange(el).setValue(data[key].text); });
        });
    }
}
function handleAddNewTranslations() {
    handleErrors();
    var changed = getNewChangesToShow(getParsedData().added);
    var ui = SpreadsheetApp.getUi();
    if (!changed.total)
        return ui.alert('Nothing to change');
    var modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
    modal.description = "Are you sure to add new translations?";
    modal.total = changed.total;
    modal.translations = changed.changes;
    modal.shouldShowDeleteRowsCheckBox = true;
    modal.modalType = "new";
    var evaluatedModal = modal.evaluate();
    ui.showModalDialog(evaluatedModal, 'Add new translations');
}
function handleAddExistingTranslations() {
    handleErrors();
    var changed = getExistingChangesToShow(getParsedData().existing);
    var ui = SpreadsheetApp.getUi();
    if (!changed.total)
        return ui.alert('Nothing to change');
    var modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
    modal.description = "Are you sure to add existing translations?";
    modal.total = changed.total;
    modal.translations = changed.changes;
    modal.shouldShowDeleteRowsCheckBox = true;
    modal.modalType = "existing";
    var evaluatedModal = modal.evaluate();
    ui.showModalDialog(evaluatedModal, 'Add existing translations');
}
function handleDeleteRows() {
    handleErrors();
    var rowsToDelete = getRowsToDelete();
    var ui = SpreadsheetApp.getUi();
    if (!rowsToDelete.length)
        return ui.alert('Nothing to delete');
    var modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
    modal.description = "Are you sure to delete rows?";
    modal.total = rowsToDelete.length;
    modal.translations = rowsToDelete;
    modal.shouldShowDeleteRowsCheckBox = false;
    modal.modalType = "deleteRows";
    var evaluatedModal = modal.evaluate();
    ui.showModalDialog(evaluatedModal, 'Delete rows');
}
function updateTranslationsStatus() {
    var maxRows = getMainSheet().getMaxRows();
    var data = getMainSheet().getRange("G5:L".concat(maxRows)).getDisplayValues();
    var statuses = getMainSheet().getRange("A5:A".concat(maxRows)).getDisplayValues();
    var result = data.reduce(function (prev, curr, index) {
        var _a, _b, _c, _d, _e;
        var currentRow = index + 5;
        var en = curr[0].trim().normalize();
        var ar = curr[1].trim().normalize();
        var allEqual = curr.every(function (el) { return el === curr[0]; });
        var allToBeDeleted = curr.every(function (el) { return el === '-'; });
        var currWithoutAr = curr.filter(function (_, index) { return index !== 1; });
        var equalWithoutAr = currWithoutAr.every(function (el) { return el === currWithoutAr[0]; });
        var isDeleteStatus = statuses[index].includes('delete');
        var isTranslationMissing = curr.some(function (el) { return el === ''; });
        if (curr[0] === '' || isDeleteStatus)
            return prev;
        if (isTranslationMissing)
            return __assign(__assign({}, prev), (_a = {}, _a[currentRow] = 'missing', _a));
        if (allToBeDeleted)
            return __assign(__assign({}, prev), (_b = {}, _b[currentRow] = 'delete', _b));
        if (allEqual)
            return __assign(__assign({}, prev), (_c = {}, _c[currentRow] = 'todo', _c));
        if (en !== ar && equalWithoutAr)
            return __assign(__assign({}, prev), (_d = {}, _d[currentRow] = 'only-arabic', _d));
        return __assign(__assign({}, prev), (_e = {}, _e[currentRow] = '-', _e));
    }, {});
    Object.keys(result).forEach(function (el) {
        getMainSheet().getRange("A".concat(el)).setValue(result[el]);
    });
}
function updateTranslationsForLoop() {
    var maxRows = getMainSheet().getMaxRows();
    var data = {};
    var _loop_1 = function (i) {
        var currentTranslations = getMainSheet().getRange("G".concat(i, ":L").concat(i)).getDisplayValues().pop();
        var currentStatus = getMainSheet().getRange("A".concat(i)).getDisplayValues().pop();
        var en = currentTranslations[0].trim().normalize();
        var ar = currentTranslations[1].trim().normalize();
        var allEqual = currentTranslations.every(function (el) { return el === currentTranslations[0]; });
        var allToBeDeleted = currentTranslations.every(function (el) { return el === '-'; });
        var currWithoutAr = currentTranslations.filter(function (_, index) { return index !== 1; });
        var equalWithoutAr = currWithoutAr.every(function (el) { return el === currWithoutAr[0]; });
        var isDeleteStatus = currentStatus.includes('delete');
        var isTranslationMissing = currentTranslations.some(function (el) { return el === ''; });
        //if(currentTranslations[0] === '' || isDeleteStatus) return data;
        if (isTranslationMissing)
            getMainSheet().getRange("A".concat(i)).setValue('missing');
        if (allToBeDeleted)
            getMainSheet().getRange("A".concat(i)).setValue('delete');
        if (allEqual)
            getMainSheet().getRange("A".concat(i)).setValue('todo');
        if (en !== ar && equalWithoutAr)
            getMainSheet().getRange("A".concat(i)).setValue('only-arabic');
    };
    for (var i = 5; i <= maxRows; i++) {
        _loop_1(i);
    }
    ;
}
function handleUpdateTranslationsStatus() {
    var ui = SpreadsheetApp.getUi();
    if (SpreadsheetApp.getActiveSheet().getSheetName() !== SETTINGS.sheets.main) {
        throw ui.alert('Invalid sheet!\r\n Script works only on: ' + SETTINGS.sheets.main);
    }
    ;
    var modal = HtmlService.createTemplateFromFile('TrackerModalTemplate');
    modal.description = "Are you sure to update translations status?";
    modal.total = 0;
    modal.translations = [];
    modal.shouldShowDeleteRowsCheckBox = false;
    modal.modalType = "updateTranslationsStatus";
    var evaluatedModal = modal.evaluate();
    ui.showModalDialog(evaluatedModal, 'Update translations status');
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
}
;
