const routeTypeMapping = {
  'Cash': ['C', 'D', 'E'],
  'Product': ['H', 'I', 'J'],
  'Gold Apple Product': ['M', 'N', 'O'],
  'Admin': ['R', 'S', 'T'],
};

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Talaria")
    .addItem("Sync All data", 'all')
    .addToUi()
}

function onInstall() {
  onOpen();
}

function readExternalSheet(sheetName) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let sheetData = sheet.getDataRange().getValues();
  return sheetData
}

function getDates(sheet) {
  let startDate = new Date(sheet.getRange("C1").getValue());
  let endDate = new Date(sheet.getRange("C2").getValue());
  return {startDate, endDate};
}

function alertSuccess() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Success", `Sheet updated successfully!`, ui.ButtonSet.OK);
}

function routesbystateFunction() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const inputRoutesSheetData = readExternalSheet('Input Routes');
  const {startDate, endDate} = getDates(activeSheet);

  for (let i = 37; i < 54; i++) {
    let stateName = activeSheet.getRange(`L${i}`).getValue();
    if (stateName) {
      let totalRoutes = inputRoutesSheetData.filter(row => {
        let date = new Date(row[2]);
        let state = row[5];
        return date >= startDate && date <= endDate && state === stateName;
      }).length;
      activeSheet.getRange(`M${i}`).setValue(totalRoutes);
    }
  }
  alertSuccess();
}

function routesTypeFunction(routeType) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const inputRoutesSheetData = readExternalSheet('Input Routes');
  const {startDate, endDate} = getDates(activeSheet);

  for (let i = 10; i < 31; i++) {
    let hubName = activeSheet.getRange(`B${i}`).getValue();
    if (hubName) {
      let [totalCosts, totalInvoice, rowCount] = inputRoutesSheetData.filter(row => {
        let date = new Date(row[2]);
        let hub = row[3];
        let route_Type = row[12];
        return date >= startDate && date <= endDate && hub === hubName && route_Type === routeType;
      }).reduce((acc, row) => {
        if(typeof row[28] === 'number') acc[0] += Number(row[28]);
        if(typeof row[29] === 'number') acc[1] += Number(row[29]);
        acc[2]++;
        return acc;
      }, [0, 0, 0]);

      let [totalCostsCell, totalInvoiceCell, cogsCell] = routeTypeMapping[routeType];
      activeSheet.getRange(`${totalCostsCell}${i}`).setValue(totalCosts);
      activeSheet.getRange(`${totalInvoiceCell}${i}`).setValue(totalInvoice);
      activeSheet.getRange(`${cogsCell}${i}`).setValue(rowCount === 0 ? '0.00%' : `${(totalCosts / totalInvoice * 100).toFixed(2)}%`);
    }
  }
}

function customer_profitability_Function() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const inputRoutesSheetData = readExternalSheet('Input Routes');
  const {startDate, endDate} = getDates(activeSheet);

  for (let i = 37; i < 87; i++) {
    let customerName = activeSheet.getRange(`B${i}`).getValue();
    if (customerName) {
      let [totalCosts, totalInvoice, rowCount] = inputRoutesSheetData.filter(row => {
        let date = new Date(row[2]);
        let customer_name = row[4];
        return date >= startDate && date <= endDate && customer_name === customerName;
      }).reduce((acc, row) => {
        if(typeof row[28] === 'number') acc[0] += Number(row[28]);
        if(typeof row[29] === 'number') acc[1] += Number(row[29]);
        acc[2]++;
        return acc;
      }, [0, 0, 0]);

      activeSheet.getRange(`C${i}`).setValue(totalCosts);
      activeSheet.getRange(`D${i}`).setValue(totalInvoice);
      activeSheet.getRange(`E${i}`).setValue(rowCount === 0 ? '0.00%' : `${(totalCosts / totalInvoice * 100).toFixed(2)}%`);
    }
  }
}

function all(){
  ['Cash', 'Product', 'Gold Apple Product', 'Admin'].forEach(routesTypeFunction);
  customer_profitability_Function();
  routesbystateFunction();
}