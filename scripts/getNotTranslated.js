function getNoTranslated() {

  const languageIndex = {
    0: "en",
    1: 'ar',
    2: 'bn',
    3: 'hi',
    4: 'fil',
    5: 'ur',
  }


  const sheets = SpreadsheetApp.getActive().getSheets();
  const mainSheet = sheets[0];
  const translationSheet = sheets[1];
  const mainSheetMaxRows = mainSheet.getMaxRows() - 5;

  const allRawData = mainSheet.getRange(`G5:L300`).getDisplayValues();
  
  const missing = [];

  allRawData.forEach((el,index) => {
    const row = index + 5
    const en = el[0];

    if(!en.length) return;

    const currentTranslationStatus = [];

    el.forEach((singleCell, index) => {
      if(index > 0) {
        if(singleCell === en) {
          currentTranslationStatus.push('❌')
        } else {
          currentTranslationStatus.push('✅')
        }
      }
    }
    )
    if(currentTranslationStatus.filter((el) => el !== '✅').length) {
      missing.push([[en], ...currentTranslationStatus, row]);
    }
  })
  Logger.log(missing)
  //translationSheet.getRange(`S5:Y${missing.length+4}`).setValues(missing)
}

