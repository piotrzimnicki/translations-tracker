function myFunction() {
  const dictionary = {};
  const data = SpreadsheetApp.getActive().getSheetByName('Sheet1').getDataRange().getValues();  
  data.forEach((row, index) => {
    if(index < 4) return;
    const firstKey = row[1];
    if(!firstKey) return;
    if(dictionary[firstKey]) {
      dictionary[firstKey].push(index + 1);
    } else {
      dictionary[firstKey] = [index + 1];
    }
  })
  Object.entries(dictionary)
  .map((e) => e).sort(([_a, a], [_b, b]) => b.length - a.length)
  .map(([key, value]) => {
    if(value.length === 1) return;
    Logger.log(`${key}: ${value.join(',')}`);
  })
}
