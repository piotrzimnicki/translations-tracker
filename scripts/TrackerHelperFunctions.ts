import {ParsedData} from './TrackerMainScripts'

const SETTINGS = {
  sheets: {
    main: 'Sheet1',
    tracker: 'Tracker',
  },
  emptyCell: '-',
  notFound: 'not found',
  getDataRange: (maxRows: number) => `A5:G${maxRows}`,
  filters: {
    first: {
      cell: 'C1',
      default: 'Row'
    },
    second: {
      cell: 'D1',
      default: 'Row',
    }
  },
  dataStartCells: {
    en: 'A5',
    ar: 'B5',
    rowsExact: 'C5',
    rowsPartial: 'D5',
    linkToFirstExactRow: 'E5',
    alreadyTranslatedRows: 'F5',
    alreadyTranslatedRowsFirstLink: 'G5', 
  }
}

function getTrackerSheet() {
  return SpreadsheetApp.getActive().getSheetByName(SETTINGS.sheets.tracker);
}

function getMainSheet() {
  return SpreadsheetApp.getActive().getSheetByName(SETTINGS.sheets.main);
}


function isActiveSheet (sheet: keyof typeof SETTINGS.sheets) {
  const activeSheetName = SpreadsheetApp.getActiveSheet().getSheetName();
  return activeSheetName === sheet;
}

function isFilterActive() {
  const { filters: {first, second}} = SETTINGS;
  const trackerSheet = getTrackerSheet();
  const firstFilter = trackerSheet.getRange(first.cell).getDisplayValue();
  const secondFilter = trackerSheet.getRange(second.cell).getDisplayValue();

  return firstFilter !== first.default || secondFilter !== second.default
}

function handleErrors() {
  const { filters: {first, second}} = SETTINGS;
  const ui = SpreadsheetApp.getUi();
  if(!isActiveSheet('tracker')) {
    throw ui.alert('Invalid sheet!\r\n Script works only on: ' + SETTINGS.sheets.tracker); 
  };
  if(isFilterActive()) {
    throw ui.alert(`Please set filters (${first.cell}/${second.cell}) to "${first.default}/${second.default}"`);
  }
  return;
}

function getTotalFromObject(obj: Record<string, unknown>) {
  if(!obj) return 0;
  return Object.keys(obj).length;
}

function getExistingChangesToShow(obj: ParsedData['existing']) {
  if(!obj) {
    return {
      total: 0,
      changes: [],
    }
  }
  return {
      total: Object.keys(obj).length,
      changes: Object.keys(obj).map((el) => `from ${obj[el].from} to ${obj[el].to}`),
    }
}

function getNewChangesToShow(obj: ParsedData['added']) {
  if(!obj) {
    return {
      total: 0,
      changes: [],
    }
  }
  return {
      total: Object.keys(obj).length,
      changes: Object.keys(obj).map((el) => `${obj[el].text} to ${obj[el].to}`),
    }
}

function getRowsToDelete() {
  const {dataStartCells: {rowsExact}, notFound } = SETTINGS;
  const trackerSheet = getTrackerSheet();
  const trackerSheetMaxRows = trackerSheet.getMaxRows();
  const trackerSheetRawData = trackerSheet.getRange(`${rowsExact}:C${trackerSheetMaxRows}`).getDisplayValues(); 

  return trackerSheetRawData.reduce((prev,curr,index) => {
    const shouldBeDeleted = curr[0] === notFound;
    const currentIndex = index + 5;

    if(shouldBeDeleted) return [...prev, currentIndex];

    return prev;
    
  }, [] as number[]);
}

const sortRowsIntoDeleteGroups = (arr: number[]) => {
  let result: number[][] = [];
  let tempArr: number[] = [];
 
  for(let i = 0; i <= arr.length - 1; i++) {  
    const current = arr[i];
    const next = arr[i+1];
    const isNextNumberValid = next - current === 1;
    
    tempArr.push(current);
    if(!isNextNumberValid) {
      result.push(tempArr);
      tempArr = [];
    }
  }
  
  result.sort(function (a, b) { return b[b.length - 1] - a[a.length - 1]; });
  
  return result.reduce((prev, curr) => {
    return [...prev, [curr[0], curr.length]]
  }, [] as number[][]);

}

function deleteNotFoundRows() {
  const trackerSheet = getTrackerSheet()
  const rowsToDelete = sortRowsIntoDeleteGroups(getRowsToDelete());

  for(let i = 0; i <= rowsToDelete.length - 1; i++) {
    trackerSheet.deleteRows(rowsToDelete[i][0], rowsToDelete[i][1]);
  }
}