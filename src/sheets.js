function appendShippingToSheet_(shippingData, config) {
  const sheet = getRequiredSheetByName_(config.masterSheetName);
  insertRequestRow_(sheet, toShippingRow_(shippingData));
}

function appendSupplyOrderToSheet_(orderData, config) {
  const sheet = getRequiredSheetByName_(config.supplySheetName);
  insertRequestRow_(sheet, toSupplyOrderRow_(orderData));
}

function insertRequestRow_(sheet, rowValues) {
  sheet.insertRowsBefore(2, 1);

  const rowRange = sheet.getRange(2, 1, 1, rowValues.length);
  rowRange.setValues([rowValues]);

  const checkboxCell = sheet.getRange(2, rowValues.length);
  checkboxCell.insertCheckboxes();
  checkboxCell.setValue(false);
}

function getRequiredSheetByName_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Active spreadsheet is not available. Use a bound script.');
  }

  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  return sheet;
}
