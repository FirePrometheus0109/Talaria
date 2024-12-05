function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  if (editedRow < 3 || editedCol < 3 || editedCol >= 150 ) return;

  sheet.getRange(editedRow, 1, 1, 2).setValues([[new Date(), Session.getActiveUser().getEmail()]]);

  if (editedCol == 7 && editedRow >= 3 && e.range.getValue() !== "") {
    sheet.getRange(editedRow, 16).setValue("Terminated");
  }

}

function onEditInstall(e) {
  const masterSheet = SpreadsheetApp.openById('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo').getSheetByName('Current Quarter');
  var sheet = e.source.getActiveSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  if (editedRow < 345 || editedCol < 3 || editedCol >= 150 ) return;

  var mastersheetCol = 236;
  var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameValue = rowData[4]; // Column 15
  var hubValue = rowData[14]; // Column 15
  var timeValue = rowData[15]; // Column 16
  var sheet_mastersheetID = rowData[mastersheetCol - 1]; // Column 27

  switch (editedCol) {
    case 15: // HubName
      if(sheet_mastersheetID){
        masterSheet.deleteRow(sheet_mastersheetID);
      }
      var masterSheetData = masterSheet.getDataRange().getValues();
      var index = masterSheetData.reverse().findIndex(row => row[0] == hubValue && (row[3] == 'Full-Time' || row[3] == 'Part-Time'));
      var new_mastersheedId = index >= 0 ? masterSheetData.length - index : -1;
      if (new_mastersheedId > 0) {
        sheet.getRange(editedRow, mastersheetCol).setValue(new_mastersheedId + 1);
        masterSheet.insertRowAfter(new_mastersheedId);
        masterSheet.getRange(new_mastersheedId + 1, 1, 1, masterSheet.getLastColumn()).setBackground('white').setFontColor('black');
        var status = timeValue.trim().toLowerCase().includes('full-time') ? 'Full-Time' : timeValue.trim().toLowerCase().includes('part-time') ? 'Part-Time' : '';
        var fullName = nameValue? nameValue : '';
        masterSheet.getRange(new_mastersheedId + 1, 4).setValue(status);
        masterSheet.getRange(new_mastersheedId + 1, 2).setValue(fullName);
        masterSheet.getRange(new_mastersheedId + 1, 1).setValue(hubValue)
      }
      break;
    case 5: // Full-Name
      if(!sheet_mastersheetID) return;
      masterSheet.getRange(sheet_mastersheetID, 2).setValue(nameValue);
      break;
    case 16: //Full-time or Part-time
      if(!sheet_mastersheetID) return;
      var status = timeValue.trim().toLowerCase().includes('full-time') ? 'Full-Time' : timeValue.trim().toLowerCase().includes('part-time') ? 'Part-Time' : '';
      masterSheet.getRange(sheet_mastersheetID, 4).setValue(status);
      break;
    default:
      break;
  }
}