function onEditInstall(e) {
  const masterSheet = SpreadsheetApp.openById('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo').getSheetByName('Current Quarter');

  var sheet = e.source.getActiveSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  if (editedRow < 3 || editedCol < 3) return;

  if (editedCol == 7 && editedRow >= 3) {
    var cellValue = e.range.getValue();
    if (cellValue !== "") {
      var statusCol = 16;
      var terminatedValue = "Terminated";

      sheet.getRange(editedRow, statusCol).setValue(terminatedValue);
    }
  }

  var timeStampCol = 1;
  var userCol = 2;
  var dateTime = new Date();
  var userEmail = Session.getActiveUser().getEmail();

  sheet.getRange(editedRow, timeStampCol).setValue(dateTime);
  sheet.getRange(editedRow, userCol).setValue(userEmail);

  var mastersheetCol = 27;
  var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameValue = rowData[4]; // Column 15
  var hubValue = rowData[14]; // Column 15
  var timeValue = rowData[15]; // Column 16
  var sheet_mastersheetID = rowData[mastersheetCol - 1]; // Column 27

  
  if (!sheet_mastersheetID && editedRow >= 343){
    var lastRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(lastRow, 1, 1, masterSheet.getLastColumn()).setBackground('white');
    masterSheet.getRange(lastRow, 1, 1, masterSheet.getLastColumn()).setFontColor('black');
    masterSheet.getRange(lastRow, 2).setValue(nameValue);
    sheet.getRange(editedRow, mastersheetCol).setValue(lastRow);
  } else {
    switch (editedCol) {
      case 5: // Full-Name
        masterSheet.getRange(sheet_mastersheetID, 2).setValue(nameValue);
      case 15: // HubName
        var masterSheetData = masterSheet.getDataRange().getValues();
        var new_mastersheedId = -1;
        for(var i = masterSheetData.length - 1; i >= 0; i--) {
          if(masterSheetData[i][0] == hubValue && (masterSheetData[i][3] == 'Full-Time' || masterSheetData[i][3] == 'Part-Time')) {
            new_mastersheedId = i + 1;
            break;
          }
        }
        if (new_mastersheedId > 0) {
          masterSheet.insertRowAfter(new_mastersheedId);
          masterSheet.getRange(new_mastersheedId + 1, 1, 1, masterSheet.getLastColumn()).setBackground('white');
          masterSheet.getRange(new_mastersheedId + 1, 1, 1, masterSheet.getLastColumn()).setFontColor('black');
          var sourceRow = masterSheet.getRange(`${sheet_mastersheetID + 1}:${sheet_mastersheetID + 1}`);
          var targetRow = masterSheet.getRange(`${new_mastersheedId + 1}:${new_mastersheedId + 1}`);
          sheet.getRange(editedRow, mastersheetCol).setValue(new_mastersheedId + 1);
          targetRow.setValues(sourceRow.getValues());
          masterSheet.getRange(new_mastersheedId + 1, 1).setValue(hubValue);
          masterSheet.deleteRow(sheet_mastersheetID + 1);

          if (timeValue.trim().toLowerCase().includes('full-time')) {
            masterSheet.getRange(new_mastersheedId + 1, 4).setValue('Full-Time');
          } else if(timeValue.trim().toLowerCase().includes('part-time')) {
            masterSheet.getRange(new_mastersheedId + 1, 4).setValue('Part-Time');
          } else if(timeValue.trim().toLowerCase().includes('Salaried')){
            masterSheet.getRange(new_mastersheedId + 1, 4).setValue('');
          }
        }
      case 16: //Full-time or Part-time
        if (timeValue.trim().toLowerCase().includes('full-time')) {
          masterSheet.getRange(sheet_mastersheetID, 4).setValue('Full-Time');
        } else if(timeValue.trim().toLowerCase().includes('part-time')) {
          masterSheet.getRange(sheet_mastersheetID, 4).setValue('Part-Time');
        } else if(timeValue.trim().toLowerCase().includes('Salaried')){
          masterSheet.getRange(new_mastersheedId + 1, 4).setValue('');
        }
      default:
        break;
    }
  }
}

