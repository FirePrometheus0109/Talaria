function onEditInstall(e) {
  const ui = SpreadsheetApp.getUi()
  const masterSheet = SpreadsheetApp.openById('1bEDr0RbmwtFCmIrZ623B-A29SU3EPAK-da-nprCetDo').getSheetByName('Current Quarter');
  var is_new = false;

  var sheet = e.source.getActiveSheet();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  if (editedRow < 3 || editedCol < 3 || editedCol >= 150 ) return;

  if (!sheet.getRange(editedRow, 1).getValue() && !sheet.getRange(editedRow, 2).getValue()) {
    is_new = true;
  }

  sheet.getRange(editedRow, 1, 1, 2).setValues([[new Date(), Session.getActiveUser().getEmail()]]);

  if (editedCol == 7 && editedRow >= 3 && e.range.getValue() !== "") {
    sheet.getRange(editedRow, 16).setValue("Terminated");
  }

  var mastersheetCol = 236;
  var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nameValue = rowData[4]; // Column 15
  var hubValue = rowData[14]; // Column 15
  var timeValue = rowData[15]; // Column 16
  var sheet_mastersheetID = rowData[mastersheetCol - 1]; // Column 27

  if (!sheet_mastersheetID && is_new){
    var lastRow = masterSheet.getLastRow() + 1;
    var status = timeValue.trim().toLowerCase().includes('full-time') ? 'Full-Time' : timeValue.trim().toLowerCase().includes('part-time') ? 'Part-Time' : '';
    sheet.getRange(editedRow, mastersheetCol).setValue(lastRow);
    masterSheet.getRange(lastRow, 1, 1, masterSheet.getLastColumn()).setBackground('white').setFontColor('black');
    masterSheet.getRange(lastRow, 2).setValue(nameValue);
    masterSheet.getRange(lastRow, 4).setValue(status);
    ui.alert("Success", `Stored in row${lastRow} from Talaria mastersheet`, ui.ButtonSet.OK);
  } else if(sheet_mastersheetID && !is_new) {
    switch (editedCol) {
      case 5: // Full-Name
        masterSheet.getRange(sheet_mastersheetID, 2).setValue(nameValue);
        ui.alert("Success", `Updated Full Name from Talaria mastersheet`, ui.ButtonSet.OK);
        break;
      case 15: // HubName
        var masterSheetData = masterSheet.getDataRange().getValues();
        var index = masterSheetData.reverse().findIndex(row => row[0] == hubValue && (row[3] == 'Full-Time' || row[3] == 'Part-Time'));
        var new_mastersheedId = index >= 0 ? masterSheetData.length - index : -1;

        if (new_mastersheedId > 0) {
          if (new_mastersheedId < sheet_mastersheetID) {
            masterSheet.insertRowAfter(new_mastersheedId);
            masterSheet.getRange(new_mastersheedId + 1, 1, 1, masterSheet.getLastColumn()).setBackground('white').setFontColor('black');
            var sourceRow = masterSheet.getRange(`${sheet_mastersheetID + 1}:${sheet_mastersheetID + 1}`);
            var targetRow = masterSheet.getRange(`${new_mastersheedId + 1}:${new_mastersheedId + 1}`);
            //////Sync All mastersheet ID
            var columnValues = sheet.getRange(1, 236, sheet.getLastRow()).getValues();
            var flatColumnValues = [].concat.apply([], columnValues);
            for(var i = 0; i < flatColumnValues.length; i++) {
              if(flatColumnValues[i] >= new_mastersheedId + 1 && flatColumnValues[i]<sheet_mastersheetID) {
                flatColumnValues[i]++;
                sheet.getRange(i+1, 236).setValue(flatColumnValues[i]);
              }
            }
            sheet.getRange(editedRow, mastersheetCol).setValue(new_mastersheedId + 1);
            targetRow.setValues(sourceRow.getValues());
            masterSheet.deleteRow(sheet_mastersheetID + 1);
            masterSheet.getRange(new_mastersheedId + 1, 1).setValue(hubValue);
            ui.alert("Success", `Stored in row${new_mastersheedId + 1} from Talaria mastersheet`, ui.ButtonSet.OK);
          } else {
            masterSheet.insertRowAfter(new_mastersheedId);
            masterSheet.getRange(new_mastersheedId + 1, 1, 1, masterSheet.getLastColumn()).setBackground('white').setFontColor('black');
            var sourceRow = masterSheet.getRange(`${sheet_mastersheetID}:${sheet_mastersheetID}`);
            var targetRow = masterSheet.getRange(`${new_mastersheedId + 1}:${new_mastersheedId + 1}`);
            //////Sync All mastersheet ID
            var columnValues = sheet.getRange(1, 236, sheet.getLastRow()).getValues();
            var flatColumnValues = [].concat.apply([], columnValues);
            for(var i = 0; i < flatColumnValues.length; i++) {
              if(flatColumnValues[i] < new_mastersheedId && flatColumnValues[i]>=sheet_mastersheetID) {
                flatColumnValues[i]--;
                sheet.getRange(i+1, 236).setValue(flatColumnValues[i]);
              }
            }
            sheet.getRange(editedRow, mastersheetCol).setValue(new_mastersheedId);
            targetRow.setValues(sourceRow.getValues());
            masterSheet.deleteRow(sheet_mastersheetID);
            masterSheet.getRange(new_mastersheedId, 1).setValue(hubValue);
            ui.alert("Success", `Stored in row${new_mastersheedId} from Talaria mastersheet`, ui.ButtonSet.OK);
          }
        }
        break;
      case 16: //Full-time or Part-time
        var status = timeValue.trim().toLowerCase().includes('full-time') ? 'Full-Time' : timeValue.trim().toLowerCase().includes('part-time') ? 'Part-Time' : '';
        masterSheet.getRange(sheet_mastersheetID, 4).setValue(status);
        ui.alert("Success", `Updated Role Type from Talaria mastersheet`, ui.ButtonSet.OK);
        break;
      default:
        break;
    }
  }
}