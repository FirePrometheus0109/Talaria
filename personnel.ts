function onEditInstall(e) {
  const masterSheet = SpreadsheetApp.openById('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo').getSheetByName('Current Quarter');
  var sheet = e.source.getActiveSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  if (editedRow < 3 || editedCol < 3) return;

  if (editedCol == 7 && editedRow >= 3 && e.range.getValue() !== "") {
    sheet.getRange(editedRow, 16).setValue("Terminated");
  }

  sheet.getRange(editedRow, 1, 1, 2).setValues([[new Date(), Session.getActiveUser().getEmail()]]);

  var mastersheetCol = 27;
  var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameValue = rowData[4]; // Column 15
  var hubValue = rowData[14]; // Column 15
  var timeValue = rowData[15]; // Column 16
  var sheet_mastersheetID = rowData[mastersheetCol - 1]; // Column 27

  if (!sheet_mastersheetID && editedRow >= 343){
    var lastRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(lastRow, 1, 1, masterSheet.getLastColumn()).setBackground('white').setFontColor('black');
    masterSheet.getRange(lastRow, 2).setValue(nameValue);
    sheet.getRange(editedRow, mastersheetCol).setValue(lastRow);
  } else {
    switch (editedCol) {
      case 5: // Full-Name
        masterSheet.getRange(sheet_mastersheetID, 2).setValue(nameValue);
        break;
      case 15: // HubName
        var masterSheetData = masterSheet.getDataRange().getValues();
        var index = masterSheetData.reverse().findIndex(row => row[0] == hubValue && (row[3] == 'Full-Time' || row[3] == 'Part-Time'));
        var new_mastersheedId = index >= 0 ? masterSheetData.length - index : -1;

        if (new_mastersheedId > 0) {
          masterSheet.insertRowAfter(new_mastersheedId);
          masterSheet.getRange(new_mastersheedId + 1, 1, 1, masterSheet.getLastColumn()).setBackground('white').setFontColor('black');
          var sourceRow = masterSheet.getRange(`${sheet_mastersheetID + 1}:${sheet_mastersheetID + 1}`);
          var targetRow = masterSheet.getRange(`${new_mastersheedId + 1}:${new_mastersheedId + 1}`);
          sheet.getRange(editedRow, mastersheetCol).setValue(new_mastersheedId + 1);
          targetRow.setValues(sourceRow.getValues());
          masterSheet.getRange(new_mastersheedId + 1, 1).setValue(hubValue);
          masterSheet.deleteRow(sheet_mastersheetID + 1);

          var status = timeValue.trim().toLowerCase().includes('full-time') ? 'Full-Time' : timeValue.trim().toLowerCase().includes('part-time') ? 'Part-Time' : '';
          masterSheet.getRange(new_mastersheedId + 1, 4).setValue(status);
        }
        break;
      case 16: //Full-time or Part-time
        var status = timeValue.trim().toLowerCase().includes('full-time') ? 'Full-Time' : timeValue.trim().toLowerCase().includes('part-time') ? 'Part-Time' : '';
        masterSheet.getRange(sheet_mastersheetID, 4).setValue(status);
        break;
      default:
        break;
    }
  }
}