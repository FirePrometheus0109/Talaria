const dayMapping = {
  'MON': ['M', 'N', 'O'],
  'TUE': ['Q', 'R', 'S'],
  'WED': ['U', 'V', 'W'],
  'THU': ['Y', 'Z', 'AA'],
  'FRI': ['AC', 'AD', 'AE'],
  'SAT': ['AG', 'AH', 'AI'],
  'SUN': ['AK', 'AL', 'AM']
};

function setStartAndEndDates() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
  const currentDate = new Date();
  const currentDayOfWeek = currentDate.getDay();
  let startDate, endDate;

  if (currentDayOfWeek === 1) {
    startDate = new Date();
    startDate.setDate(currentDate.getDate() - 28);
    endDate = new Date();
    endDate.setDate(currentDate.getDate() - 1);
  }

  if (startDate && endDate){
    activeSheet.getRange("A2").setValue(startDate);
    activeSheet.getRange("B2").setValue(endDate);
    const ui = SpreadsheetApp.getUi()
    ui.alert("Success", `Set A2 and B2`, ui.ButtonSet.OK);
  }
}

function onOpen() {
  let ui = SpreadsheetApp.getUi()
  ui.createMenu("Talaria")
    .addItem("Sync Capacity Tracker Data", 'capacitytrackerFunction')
    .addItem("Sync Capacity Tracker V2 Data", 'capacitytrackerv2Function')
    .addItem("Sync Capacity Tracker V2 - Daily Avg Data", 'capacitytrackerv2_daily_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Mon Avg Data", 'capacitytrackerv2_MON_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Tue Avg Data", 'capacitytrackerv2_TUE_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Wed Avg Data", 'capacitytrackerv2_WED_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Thu Avg Data", 'capacitytrackerv2_THU_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Fri Avg Data", 'capacitytrackerv2_FRI_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Sat Avg Data", 'capacitytrackerv2_SAT_avg_Function')
    // .addItem("Sync Capacity Tracker V2 - Sun Avg Data", 'capacitytrackerv2_SUN_avg_Function')
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

function capacitytrackerFunction() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Capacity Tracker");
  const masterSheetData = readExternalSheet('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo', 'Current Quarter')

  let startDate = new Date(activeSheet.getRange("A1").getValue());
  let endDate = new Date(activeSheet.getRange("B1").getValue());

  for (let i = 3; i < 25; i++) {
    let job = activeSheet.getRange(`C${i}`).getValue();

    if (job) {
      let nonCashOrVaultCount = 0;
      let cashCount = 0;
      let nonVaultOrChangeCount = 0;
      let firstJobData = masterSheetData.filter(item => item[0] === job)
      let columnIndexes = getColumnIndexes(masterSheetData[2], startDate, endDate)

      var allExtractedData = [];
      for (var j = 0; j < firstJobData.length; j++) {
        var innerArray = firstJobData[j];
        var extractedData = columnIndexes.map((index) => {
          return innerArray[index];
        });
        allExtractedData.push(extractedData);
      }

      let flattenedArray = allExtractedData.flat();
      let nonEmptyValues = flattenedArray.filter(value => value.trim() !== '' && !value.trim().toLowerCase().includes('available'));
      let uniqueValues = [...new Set(nonEmptyValues)];

      nonCashOrVaultCount = uniqueValues.filter(value => !value.trim().toLowerCase().includes('cash') && !value.trim().toLowerCase().includes('vault')).length;
      cashCount = uniqueValues.filter(value => value.trim().toLowerCase().startsWith('cash')).length;
      nonVaultOrChangeCount = nonEmptyValues.filter(value => value.trim().toLowerCase().includes('vault') || value.trim().toLowerCase().includes('change')).length;

      activeSheet.getRange(`G${i}`).setValue(nonCashOrVaultCount);
      activeSheet.getRange(`H${i}`).setValue(cashCount);
      activeSheet.getRange(`I${i}`).setValue(nonVaultOrChangeCount);
    }

  }
  const ui = SpreadsheetApp.getUi()
  ui.alert("Success", `Sheet updated successfully!`, ui.ButtonSet.OK);
}

function capacitytrackerv2Function() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
  const masterSheetData = readExternalSheet('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo', 'Current Quarter')
  let startDate = new Date(activeSheet.getRange("A2").getValue());
  let endDate = new Date(activeSheet.getRange("B2").getValue());

  for (let i = 4; i < 25; i++) {
    let job = activeSheet.getRange(`D${i}`).getValue();
    if (job) {
      let nonCashOrAvailableCount = 0;
      let cashCount = 0;
      let vaultCount = 0;
      let firstJobData = masterSheetData.filter(item => item[0] === job && (item[3] === 'Full-Time' || item[3] === 'Part-Time'));

      for (let currentDate = new Date(startDate); currentDate <= endDate; currentDate = new Date(currentDate.getTime() + 7 * 24 * 60 * 60 * 1000)) {
        let columnIndexes = getColumnIndexes(masterSheetData[2], currentDate, new Date(currentDate.getTime() + 6 * 24 * 60 * 60 * 1000));

        var allExtractedData = [];
        for (var j = 0; j < firstJobData.length; j++) {
          var innerArray = firstJobData[j];
          var extractedData = columnIndexes.map((index) => {
            return innerArray[index];
          });
          allExtractedData.push(extractedData);
        }

        let flattenedArray = allExtractedData.flat();

        let nonEmptyValues = flattenedArray.filter(value => value.trim() !== '' && !value.trim().toLowerCase().includes('available'));
        let uniqueValues = [...new Set(nonEmptyValues)];

        nonCashOrAvailableCount += uniqueValues.filter(value => !value.trim().toLowerCase().includes('cash')).length;
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

function capacitytrackerv2_Weekday_Function(weekday) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fixed Capacity Tracker V2 - Fixed Time Period");
  const masterSheetData = readExternalSheet('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo', 'Current Quarter')
  let startDate = new Date(activeSheet.getRange("A2").getValue());
  let endDate = new Date(activeSheet.getRange("B2").getValue());
  for (let i = 4; i < 25; i++) {
    let job = activeSheet.getRange(`D${i}`).getValue();
    if (job) {
      let nonCashOrVaultCount = 0;
      let cashCount = 0;
      let vaultCount = 0;
      let firstJobData = masterSheetData.filter(item => item[0] === job && (item[3] === 'Full-Time' || item[3] === 'Part-Time'));

      for (let currentDate = new Date(startDate); currentDate <= endDate; currentDate = new Date(currentDate.getTime() + 7 * 24 * 60 * 60 * 1000)) {
        let columnIndexes = getColumnIndexes(masterSheetData[2], currentDate, new Date(currentDate.getTime() + 6 * 24 * 60 * 60 * 1000)).filter(colIndex => {
          return masterSheetData[1][colIndex] === `${weekday}`;
        });

        var allExtractedData = [];
        for (var j = 0; j < firstJobData.length; j++) {
          var innerArray = firstJobData[j];
          var extractedData = columnIndexes.map((index) => {
            return innerArray[index];
          });
          allExtractedData.push(extractedData);
        }

        let flattenedArray = allExtractedData.flat();

        let nonEmptyValues = flattenedArray.filter(value => value.trim() !== '' && !value.trim().toLowerCase().includes('available'));
        let uniqueValues = [...new Set(nonEmptyValues)];

        nonCashOrVaultCount += uniqueValues.filter(value => !value.trim().toLowerCase().includes('cash') && !value.trim().toLowerCase().includes('vault') && !value.trim().toLowerCase().includes('off')).length;
        cashCount += uniqueValues.filter(value => value.trim().toLowerCase().includes('cash')).length;
        vaultCount += nonEmptyValues.filter(value => value.trim().toLowerCase().includes('vault')).length;
      }

      let [nonCashOrVaultCell, cashCell, vaultCell] = dayMapping[weekday];
      activeSheet.getRange(`${nonCashOrVaultCell}${i}`).setValue(nonCashOrVaultCount / 4);
      activeSheet.getRange(`${cashCell}${i}`).setValue(cashCount / 4);
      activeSheet.getRange(`${vaultCell}${i}`).setValue(vaultCount / 4);
    }
  }
  if(weekday === 'SUN'){
    const ui = SpreadsheetApp.getUi()
    ui.alert("Success", `Sheet updated successfully!`, ui.ButtonSet.OK);
  }
}

function capacitytrackerv2_daily_avg_Function() {
  capacitytrackerv2_Weekday_Function('MON');
  capacitytrackerv2_Weekday_Function('TUE');
  capacitytrackerv2_Weekday_Function('WED');
  capacitytrackerv2_Weekday_Function('THU');
  capacitytrackerv2_Weekday_Function('FRI');
  capacitytrackerv2_Weekday_Function('SAT');
  capacitytrackerv2_Weekday_Function('SUN');
}

// function capacitytrackerv2_MON_avg_Function() {
//   capacitytrackerv2_Weekday_Function('MON');
// }

// function capacitytrackerv2_TUE_avg_Function() {
//   capacitytrackerv2_Weekday_Function('TUE');
// }

// function capacitytrackerv2_WED_avg_Function() {
//   capacitytrackerv2_Weekday_Function('WED');
// }

// function capacitytrackerv2_THU_avg_Function() {
//   capacitytrackerv2_Weekday_Function('THU');
// }

// function capacitytrackerv2_FRI_avg_Function() {
//   capacitytrackerv2_Weekday_Function('FRI');
// }

// function capacitytrackerv2_SAT_avg_Function() {
//   capacitytrackerv2_Weekday_Function('SAT');
// }

// function capacitytrackerv2_SUN_avg_Function() {
//   capacitytrackerv2_Weekday_Function('SUN');
// }