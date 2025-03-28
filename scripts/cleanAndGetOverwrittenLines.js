// Compiled using undefined undefined (TypeScript 4.9.5)
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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
// Sheet config
var HEADER_LENGTH = 4;
var SHEET_NAME = 'Sheet1';
var COLUMN_OFFSET = 1;
var LANGUAGES_COUNT = 5;
// Defaults
var EMPTY_CELL_VALUE = '';
var getError = function (errorArray, errorMessage, successMessage) {
    if (errorArray.length) {
        Logger.log(errorMessage(errorArray.length));
        errorArray.map(Logger.log);
        return true;
    }
    Logger.log(successMessage);
    return false;
};
var getAreCellsProperlyFilled = function (data) {
    var errors = [];
    data.forEach(function (row, index) {
        // Skip header
        if (index < HEADER_LENGTH)
            return;
        // Get relevant cells for keys
        var keysRow = row.slice(COLUMN_OFFSET, LANGUAGES_COUNT + COLUMN_OFFSET);
        // Count the amount of filled cells
        var filledCells = keysRow.slice().filter(function (cell) { return cell !== ''; }).length;
        if (filledCells === 0) {
            // If no cells are filled, add to noCellsErrors error array
            errors.push("Row ".concat(index + 1 + HEADER_LENGTH, ": Empty cells"));
        }
        else if (filledCells > 1) {
            // If more than 1 cell are filled, add to multipleCellsErrors error array
            errors.push("Row ".concat(index + 1 + HEADER_LENGTH, ": Too many cells filled"));
        }
    });
    return !getError(errors, function (errorCount) { return "Found ".concat(errorCount, " rows with improperly filled cells"); }, 'No rows with improperly filled cells found :okok:');
};
var getAreCellsProperlySpacedBetweenRows = function (data) {
    var errors = [];
    // Tracking state of the position of the cell in the previous row - also allows us to skip the first check
    var previousKeyIndex = undefined;
    data.forEach(function (row, index, array) {
        // Skip header
        if (index < HEADER_LENGTH)
            return;
        // Get relevant cells for keys
        var keysRow = row.slice(COLUMN_OFFSET, LANGUAGES_COUNT + COLUMN_OFFSET);
        // Get the index of the key in keys row
        var currentKeyIndex = keysRow.findIndex(function (key) { return key !== EMPTY_CELL_VALUE; });
        if (previousKeyIndex && currentKeyIndex - previousKeyIndex > 1) {
            // If we have previousKeyIndex (we aren't in the first row) and the key is more to the right than only 1 cell (but not the other way around, since that's valid), add error to an errot array
            // example of invalid case:
            // PAGE_HEADER  |       |       |
            //              |       | TITLE |
            //
            // example of valid case:
            // PAGE_HEADER  |       |       |
            //              | TITLE |       |
            //
            // also valid case:
            //              |       | TITLE |
            // PAGE_FOOTER  |       |       |
            var previousRow = array[index - 1].slice(COLUMN_OFFSET, LANGUAGES_COUNT + COLUMN_OFFSET);
            errors.push("Row ".concat(index + 1 + HEADER_LENGTH, ": [").concat(keysRow, "] (previous: [").concat(previousRow, "])"));
        }
        // update previous key index
        previousKeyIndex = currentKeyIndex;
    });
    return !getError(errors, function (errorCount) { return "Found ".concat(errorCount, " rows incorrectly spaced cells"); }, 'No rows with improperly spaced cells found :okok:');
};
var getOverwrittenLines = function (data) {
    // Initialize an array that will hold values to construct key paths
    var breadcrumbs = Array(LANGUAGES_COUNT).fill(EMPTY_CELL_VALUE);
    // Initialize the dictionary
    var dictionary = {};
    data.forEach(function (row, rowIndex) {
        // Skip header
        if (rowIndex < HEADER_LENGTH)
            return;
        // Get relevant cells for keys
        var currentKeysRow = row.slice(COLUMN_OFFSET, LANGUAGES_COUNT + COLUMN_OFFSET);
        // Get relevant cells for translations
        var currentTranslationsRow = row.slice(6, 12);
        // Get the index (column) of the key in the row
        var currentKeyIndex = currentKeysRow.findIndex(function (key) { return key !== EMPTY_CELL_VALUE; });
        // Manage the state of what 'path' of keys were are on
        breadcrumbs = breadcrumbs.map(function (breadcrumb, breadcrumbIndex) {
            // if the saved key is to the left of the new key, keep it
            if (breadcrumbIndex < currentKeyIndex) {
                return breadcrumb;
            }
            // if the saved key is in the same place as the new key, overwrite it with new key
            if (breadcrumbIndex === currentKeyIndex) {
                return currentKeysRow[currentKeyIndex];
            }
            // if the saved key is to the right of the new key, replace it with empty value
            if (breadcrumbIndex > currentKeyIndex) {
                return EMPTY_CELL_VALUE;
            }
        });
        // get the translation cells 
        var areTranslationsPresent = currentTranslationsRow.filter(function (translation) { return translation !== EMPTY_CELL_VALUE; }).length === 0;
        // if there are no translations, then we already used it to update breadcrumbs and don't need to go any further
        if (areTranslationsPresent) {
            return;
        }
        // create string path from breadcrumbs
        var path = breadcrumbs.filter(function (breadcrumb) { return breadcrumb !== EMPTY_CELL_VALUE; }).join('.');
        // create or update the relevant dictionary entry
        if (dictionary[path]) {
            dictionary[path] = __assign(__assign({}, dictionary[path]), { rows: __spreadArray(__spreadArray([], dictionary[path].rows, true), [rowIndex + 1 + 4], false) });
        }
        else {
            dictionary[path] = { path: path, rows: [rowIndex + 1 + 4] };
        }
    });
    // log each occurrence of a key if it appears more than once
    Object.values(dictionary).map(function (entry) {
        if (entry.rows.length > 1) {
            Logger.log("Key ".concat(entry.path, " appears in lines ").concat(entry.rows.join(', '), "."));
        }
    });
};
var main = function () {
    var data = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).getDataRange().getValues().slice(HEADER_LENGTH);
    // var areCellsProperlyFilled = getAreCellsProperlyFilled(data);
    var areCellsProperlySpacedBetweenRows = getAreCellsProperlySpacedBetweenRows(data);
    if (areCellsProperlySpacedBetweenRows) {
        getOverwrittenLines(data);
    }
};
