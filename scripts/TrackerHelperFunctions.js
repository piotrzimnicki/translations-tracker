// Compiled using tracker 1.0.0 (TypeScript 4.9.5)
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
var exports = exports || {};
var module = module || { exports: exports };
//import {ParsedData} from './TrackerMainScripts'
var SETTINGS = {
    sheets: {
        main: 'Sheet1',
        tracker: 'Tracker'
    },
    emptyCell: '-',
    notFound: 'not found',
    getDataRange: function (maxRows) { return "A5:G".concat(maxRows); },
    filters: {
        first: {
            cell: 'C1',
            name: 'Row'
        },
        second: {
            cell: 'D1',
            name: 'Row'
        }
    },
    dataStartCells: {
        en: 'A5',
        ar: 'B5',
        rowsExact: 'C5',
        rowsPartial: 'D5',
        linkToFirstExactRow: 'E5',
        alreadyTranslatedRows: 'F5',
        alreadyTranslatedRowsFirstLink: 'G5'
    }
};
function getTrackerSheet() {
    return SpreadsheetApp.getActive().getSheetByName(SETTINGS.sheets.tracker);
}
function getMainSheet() {
    return SpreadsheetApp.getActive().getSheetByName(SETTINGS.sheets.main);
}
function isActiveSheet(sheet) {
    var activeSheetName = SpreadsheetApp.getActiveSheet().getSheetName();
    return activeSheetName === SETTINGS.sheets.tracker;
}
function isFilterActive() {
    var _a = SETTINGS.filters, first = _a.first, second = _a.second;
    var trackerSheet = getTrackerSheet();
    var firstFilter = trackerSheet.getRange(first.cell).getDisplayValue();
    var secondFilter = trackerSheet.getRange(second.cell).getDisplayValue();
    return firstFilter !== first.name || secondFilter !== second.name;
}
function handleErrors() {
    var _a = SETTINGS.filters, first = _a.first, second = _a.second;
    var ui = SpreadsheetApp.getUi();
    if (!isActiveSheet('tracker')) {
        throw ui.alert('Invalid sheet!\r\n Script works only on: ' + SETTINGS.sheets.tracker);
    }
    ;
    if (isFilterActive()) {
        throw ui.alert("Please set filters (".concat(first.cell, "/").concat(second.cell, ") to \"").concat(first.name, "/").concat(second.name, "\""));
    }
    return;
}
function getTotalFromObject(obj) {
    if (!obj)
        return 0;
    return Object.keys(obj).length;
}
function getExistingChangesToShow(obj) {
    if (!obj) {
        return {
            total: 0,
            changes: []
        };
    }
    return {
        total: Object.keys(obj).length,
        changes: Object.keys(obj).map(function (el) { return "from ".concat(obj[el].from, " to ").concat(obj[el].to); })
    };
}
function getNewChangesToShow(obj) {
    if (!obj) {
        return {
            total: 0,
            changes: []
        };
    }
    return {
        total: Object.keys(obj).length,
        changes: Object.keys(obj).map(function (el) { return "".concat(obj[el].text, " to ").concat(obj[el].to); })
    };
}
function getRowsToDelete() {
    var rowsExact = SETTINGS.dataStartCells.rowsExact, notFound = SETTINGS.notFound;
    var trackerSheet = getTrackerSheet();
    var trackerSheetMaxRows = trackerSheet.getMaxRows();
    var trackerSheetRawData = trackerSheet.getRange("".concat(rowsExact, ":C").concat(trackerSheetMaxRows)).getDisplayValues();
    return trackerSheetRawData.reduce(function (prev, curr, index) {
        var shouldBeDeleted = curr[0] === notFound;
        var currentIndex = index + 5;
        if (shouldBeDeleted)
            return __spreadArray(__spreadArray([], prev, true), [currentIndex], false);
        return prev;
    }, []);
}
var sortRowsIntoDeleteGroups = function (arr) {
    var result = [];
    var tempArr = [];
    for (var i = 0; i <= arr.length - 1; i++) {
        var current = arr[i];
        var next = arr[i + 1];
        var isNextNumberValid = next - current === 1;
        tempArr.push(current);
        if (!isNextNumberValid) {
            result.push(tempArr);
            tempArr = [];
        }
    }
    result.sort(function (a, b) { return b[b.length - 1] - a[a.length - 1]; });
    return result.reduce(function (prev, curr) {
        return __spreadArray(__spreadArray([], prev, true), [[curr[0], curr.length]], false);
    }, []);
};
function deleteNotFoundRows() {
    var trackerSheet = getTrackerSheet();
    var rowsToDelete = sortRowsIntoDeleteGroups(getRowsToDelete());
    for (var i = 0; i <= rowsToDelete.length - 1; i++) {
        trackerSheet.deleteRows(rowsToDelete[i][0], rowsToDelete[i][1]);
    }
}
