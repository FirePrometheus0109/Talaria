function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu("Talaria")
    .addItem("Sync Capacity Tracker Data", 'capacitytrackerFunction')
    .addItem("Sync Capacity Tracker V2 Data", 'capacitytrackerv2Function')
    .addItem("Sync Capacity Tracker V2 - Monday Data", 'capacitytrackerv2MondayFunction')
    .addToUi()
}

function onInstall() {
  onOpen();
}

function getColumnIndexes(data, startDate, endDate) {
  let columnIndexes = []
  for (let col in data) {
    let vl = data[col]
    if (isValidDate(vl)) {
      let cellDate = new Date(vl)
      if (cellDate >= startDate && cellDate <= endDate) {
        columnIndexes.push(col)
      }
    }
  }
  return columnIndexes
}

function isValidDate(value) {
  var date = new Date(value);
  return !isNaN(date.getTime());
}

function readExternalSheet(spreadsheetID, sheetName) {
  let spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  let sheet = spreadsheet.getSheetByName(sheetName);
  let sheetData = sheet.getDataRange().getValues();
  return sheetData
}

function syncCapacityData(index) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capacity Tracker");
  const masterSheetData = readExternalSheet('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo', 'Current Quarter')

  let startDate = new Date(activeSheet.getRange("A1").getValue());
  let endDate = new Date(activeSheet.getRange("B1").getValue());

  let job = activeSheet.getRange(`C${index}`).getValue();
  let columnIndexes = getColumnIndexes(masterSheetData[2], startDate, endDate)
  let firstJobData = masterSheetData.filter(item => item[0] === job)

  var allExtractedData = [];
  for (var i = 0; i < firstJobData.length; i++) {
    var innerArray = firstJobData[i];
    var extractedData = columnIndexes.map((index) => {
      return innerArray[index];
    });
    allExtractedData.push(extractedData);
  }

  let flattenedArray = allExtractedData.flat();
  return flattenedArray
}

function syncCapacityv2Data(index, startDate, endDate) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
  const masterSheetData = readExternalSheet('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo', 'Current Quarter')

  let job = activeSheet.getRange(`D${index}`).getValue();
  let columnIndexes = getColumnIndexes(masterSheetData[2], startDate, endDate);

  let firstJobData = masterSheetData.filter(item => item[0] === job && (item[3] === 'Full-Time' || item[3] === 'Part-Time'))

  var allExtractedData = [];
  for (var i = 0; i < firstJobData.length; i++) {
    var innerArray = firstJobData[i];
    var extractedData = columnIndexes.map((index) => {
      return innerArray[index];
    });
    allExtractedData.push(extractedData);
  }

  let flattenedArray = allExtractedData.flat();
  return flattenedArray
}

function syncCapacityv2_Weekday_Data(index, startDate, endDate, weekday) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
  const masterSheetData = readExternalSheet('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo', 'Current Quarter')

  let job = activeSheet.getRange(`D${index}`).getValue();
  let columnIndexes = getColumnIndexes(masterSheetData[2], startDate, endDate).filter(colIndex => {
    return masterSheetData[1][colIndex] === `${weekday}`;
  });

  let firstJobData = masterSheetData.filter(item => item[0] === job && (item[3] === 'Full-Time' || item[3] === 'Part-Time'))

  var allExtractedData = [];
  for (var i = 0; i < firstJobData.length; i++) {
    var innerArray = firstJobData[i];
    var extractedData = columnIndexes.map((index) => {
      return innerArray[index];
    });
    allExtractedData.push(extractedData);
  }

  let flattenedArray = allExtractedData.flat();
  return flattenedArray
}



function capacitytrackerFunction() {
  for (let i = 3; i < 25; i++) {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capacity Tracker");
    let job = activeSheet.getRange(`C${i}`).getValue();

    if (job) {

      let flattenedArray = syncCapacityData(i);
      let nonEmptyValues = flattenedArray.filter(value => value.trim() !== '' && !value.trim().toLowerCase().includes('available'));
      let uniqueValues = [...new Set(nonEmptyValues)];

      let nonCashOrVaultCount = uniqueValues.filter(value => !value.trim().toLowerCase().includes('cash') && !value.trim().toLowerCase().includes('vault')).length;
      let cashCount = uniqueValues.filter(value => value.trim().toLowerCase().startsWith('cash')).length;
      let nonCashOrChangeCount = nonEmptyValues.filter(value => value.trim().toLowerCase().includes('vault') || value.trim().toLowerCase().includes('change')).length;

      activeSheet.getRange(`G${i}`).setValue(nonCashOrVaultCount);
      activeSheet.getRange(`H${i}`).setValue(cashCount);
      activeSheet.getRange(`I${i}`).setValue(nonCashOrChangeCount);
    }

  }
  const ui = SpreadsheetApp.getUi()
  ui.alert("Success", `Sheet updated successfully!`, ui.ButtonSet.OK);
}

function capacitytrackerv2Function() {
  for (let i = 4; i < 25; i++) {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
    let job = activeSheet.getRange(`D${i}`).getValue();

    let startDate = new Date(activeSheet.getRange("A2").getValue());
    let endDate = new Date(activeSheet.getRange("B2").getValue());

    if (job) {

      let nonCashOrAvailableCount = 0;
      let cashCount = 0;
      let vaultCount = 0;

      for (let currentDate = new Date(startDate); currentDate <= endDate; currentDate = new Date(currentDate.getTime() + 7 * 24 * 60 * 60 * 1000)) {

        let flattenedArray = syncCapacityv2Data(i, currentDate, new Date(currentDate.getTime() + 6 * 24 * 60 * 60 * 1000));
        let nonEmptyValues = flattenedArray.filter(value => value.trim() !== '');
        let uniqueValues = [...new Set(nonEmptyValues)];

        nonCashOrAvailableCount += uniqueValues.filter(value => !value.trim().toLowerCase().includes('cash') && !value.trim().toLowerCase().includes('available')).length;
        cashCount += uniqueValues.filter(value => value.trim().toLowerCase().includes('cash')).length;
        vaultCount += nonEmptyValues.filter(value => value.trim().toLowerCase().includes('vault')).length;
      }
      activeSheet.getRange(`J${i}`).setValue(nonCashOrAvailableCount / 4);
      activeSheet.getRange(`K${i}`).setValue(cashCount / 4);
      activeSheet.getRange(`L${i}`).setValue(vaultCount / 4);
    }
  }
  const ui = SpreadsheetApp.getUi()
  ui.alert("Success", `Sheet updated successfully!`, ui.ButtonSet.OK);
}

function capacitytrackerv2MondayFunction() {
  for (let i = 4; i < 25; i++) {
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
    let job = activeSheet.getRange(`D${i}`).getValue();

    let startDate = new Date(activeSheet.getRange("A2").getValue());
    let endDate = new Date(activeSheet.getRange("B2").getValue());

    if (job) {

      let nonCashOrVaultCount = 0;
      let cashCount = 0;
      let vaultCount = 0;

      for (let currentDate = new Date(startDate); currentDate <= endDate; currentDate = new Date(currentDate.getTime() + 7 * 24 * 60 * 60 * 1000)) {

        let flattenedArray = syncCapacityv2MondayData(i, currentDate, new Date(currentDate.getTime() + 6 * 24 * 60 * 60 * 1000));
        let nonEmptyValues = flattenedArray.filter(value => value.trim() !== '');
        let uniqueValues = [...new Set(nonEmptyValues)];

        nonCashOrVaultCount += uniqueValues.filter(value => !value.trim().toLowerCase().includes('cash') && !value.trim().toLowerCase().includes('vault')).length;
        cashCount += uniqueValues.filter(value => value.trim().toLowerCase().includes('cash')).length;
        vaultCount += nonEmptyValues.filter(value => value.trim().toLowerCase().includes('vault')).length;
      }
      activeSheet.getRange(`M${i}`).setValue(nonCashOrVaultCount / 4);
      activeSheet.getRange(`N${i}`).setValue(cashCount / 4);
      activeSheet.getRange(`O${i}`).setValue(vaultCount / 4);
    }
  }
  const ui = SpreadsheetApp.getUi()
  ui.alert("Success", `Sheet updated successfully!`, ui.ButtonSet.OK);
}