function getAppConfig_() {
  return {
    liffId: getRequiredScriptProperty_('LIFF_ID'),
    lineToken: getRequiredScriptProperty_('LINE_TOKEN'),
    masterSheetName: getRequiredScriptProperty_('MASTER_SHEET_NAME'),
    supplySheetName: getRequiredScriptProperty_('SUPPLY_SHEET_NAME'),
    targetUserId: getRequiredScriptProperty_('TARGET_USER_ID'),
    productRequestSheetName: getOptionalScriptProperty_('PRODUCT_REQUEST_SHEET_NAME', '商品希望'),
    productRequestImageFolderId: getOptionalScriptProperty_('PRODUCT_REQUEST_IMAGE_FOLDER_ID', '')
  };
}

function authorizeRequiredScopes() {
  getAppConfig_();
  SpreadsheetApp.getActiveSpreadsheet();
  const rootFolder = DriveApp.createFolder('liff-supply-shipping-auth-check');
  const childFolder = rootFolder.createFolder('child');
  const file = childFolder.createFile(
    Utilities.newBlob('authorization check', 'text/plain', 'auth-check.txt')
  );
  DriveApp.getFoldersByName(rootFolder.getName()).hasNext();
  DriveApp.getFolderById(rootFolder.getId()).getName();
  file.getUrl();
  rootFolder.setTrashed(true);
  Drive.Files.list({
    q: "'" + rootFolder.getId() + "' in parents",
    maxResults: 1
  });
  Drive.Files.remove(rootFolder.getId());
  ScriptApp.getProjectTriggers();
  return 'OK';
}

function getRequiredScriptProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Missing script property: ' + key);
  }
  return value;
}

function getOptionalScriptProperty_(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (value === null || value === undefined || String(value).trim() === '') {
    return fallback;
  }
  return String(value);
}

function appendShippingToSheet_(shippingData, config) {
  const sheet = getRequiredSheetByName_(config.masterSheetName);
  insertRequestRow_(sheet, toShippingRow_(shippingData));
}

function appendSupplyOrderToSheet_(orderData, config) {
  const sheet = getRequiredSheetByName_(config.supplySheetName);
  insertRequestRow_(sheet, toSupplyOrderRow_(orderData));
}

function appendProductRequestToSheet_(requestData, config) {
  const sheet = getRequiredSheetByName_(config.productRequestSheetName);
  const row = insertRequestRow_(sheet, toProductRequestRow_(requestData));
  applyProductRequestRowLinks_(sheet, row, requestData.imageFolderUrl, requestData.imageUrls || []);
}

const DASHBOARD_SHEET_NAME = 'ダッシュボード';
const REQUESTER_ORDER_SHEET_NAME = '依頼者順';
const DASHBOARD_HEADER_ROW = 8;
const DASHBOARD_DATA_START_ROW = 9;
const DASHBOARD_COLUMNS = 10;
const DASHBOARD_META_START_COLUMN = 11;
const DASHBOARD_META_COLUMNS = 2;
const DASHBOARD_DONE_COLUMN = 10;
const DASHBOARD_RENDER_MAX_ROWS = 300;
const DASHBOARD_DATE_STORE_COLUMN = 13; // M
const DASHBOARD_DATE_STORE_START_ROW = 3;
const DASHBOARD_DATE_STORE_MAX_ROWS = 300;
const DASHBOARD_BUTTON_PROTECTION_DESCRIPTION = 'Dashboard calendar button/display warning';
const DASHBOARD_DATE_STORE_PROTECTION_DESCRIPTION = 'Dashboard selected date store warning';
const DASHBOARD_STATUS_ALL = '全件';
const DASHBOARD_STATUS_PENDING = '未完了のみ';
const DASHBOARD_STATUS_DONE = '完了のみ';
const DASHBOARD_SOURCE_SHIPPING = '送り依頼';
const DASHBOARD_SOURCE_SUPPLY = '備品注文';
const DASHBOARD_SOURCE_PRODUCT = '商品希望';
const DASHBOARD_TITLE_COLOR = '#12355B';
const DASHBOARD_HEADER_COLOR = '#1F4E79';
const DASHBOARD_LABEL_COLOR = '#E8EEF7';
const DASHBOARD_SUMMARY_COLOR = '#F7F7F7';
const SHIPPING_DONE_COLUMN = 10;
const SUPPLY_DONE_COLUMN = 8;
const PRODUCT_DONE_COLUMN = 10;
const SOURCE_READ_MAX_ROWS = 100;
const HISTORY_DEFAULT_LIMIT = 10;
const HISTORY_MAX_LIMIT = 100;
const PRODUCT_IMAGE_TRASH_DAYS = 14;
const PRODUCT_IMAGE_DELETE_DAYS = 30;

function setupDashboard() {
  const config = getAppConfig_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Active spreadsheet is not available. Use a bound script.');
  }

  const dashboardSheet = getOrCreateSheet_(spreadsheet, DASHBOARD_SHEET_NAME);
  const requesterOrderSheet = getOrCreateSheet_(spreadsheet, REQUESTER_ORDER_SHEET_NAME);

  setupDashboardLayout_(dashboardSheet);
  ensureRequesterOrderSheet_(requesterOrderSheet);
  supplementRequesterOrder();
  refreshDashboard();
}

function refreshDashboard() {
  const config = getAppConfig_();
  const dashboardSheet = getRequiredSheetByName_(DASHBOARD_SHEET_NAME);
  ensureDashboardStaticLayout_(dashboardSheet);
  const targetDates = getDashboardTargetDates_(dashboardSheet);
  const statusFilter = getDashboardStatusFilter_(dashboardSheet);
  const requesterOrder = readRequesterOrder_();
  const mergedRows = buildDashboardRows_(config, targetDates, statusFilter, requesterOrder.orderMap);

  writeDashboardRows_(dashboardSheet, mergedRows.rows);
  writeDashboardSummary_(dashboardSheet, mergedRows.summary);
}

function syncDashboardDoneToSource(e) {
  if (!e || !e.range) {
    return;
  }

  const dashboardSheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (
    dashboardSheet.getName() !== DASHBOARD_SHEET_NAME ||
    row < DASHBOARD_DATA_START_ROW ||
    col !== DASHBOARD_DONE_COLUMN
  ) {
    return;
  }

  const metaValues = dashboardSheet.getRange(row, DASHBOARD_META_START_COLUMN, 1, DASHBOARD_META_COLUMNS).getValues()[0];
  const sourceSheetName = normalizeTextKey_(metaValues[0]);
  const sourceRow = Number(metaValues[1]);
  if (!sourceSheetName || !Number.isFinite(sourceRow) || sourceRow < 2) {
    throw new Error('ダッシュボードの元行情報が見つかりません。更新メニューで再描画してください。');
  }

  const sourceSheet = getRequiredSheetByName_(sourceSheetName);
  const doneColumn = getDoneColumnFromHeader_(sourceSheet);
  const done = toDoneBoolean_(e.range.getValue());
  sourceSheet.getRange(sourceRow, doneColumn).setValue(done);
  refreshDashboard();
}

function getDoneColumnFromHeader_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    throw new Error('済列が見つかりません: ' + sheet.getName());
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (let index = headers.length - 1; index >= 0; index -= 1) {
    if (normalizeTextKey_(headers[index]) === '済') {
      return index + 1;
    }
  }
  throw new Error('済列が見つかりません: ' + sheet.getName());
}

function supplementRequesterOrder() {
  const config = getAppConfig_();
  const orderSheet = getRequiredSheetByName_(REQUESTER_ORDER_SHEET_NAME);
  const requesterOrder = readRequesterOrder_();
  const known = requesterOrder.orderMap;

  const shippingRows = readShippingRows_(config.masterSheetName);
  const supplyRows = readSupplyRows_(config.supplySheetName);
  const productRows = readProductRequestRows_(config.productRequestSheetName);
  const namesToAppend = [];

  shippingRows.concat(supplyRows).concat(productRows).forEach(function (row) {
    const name = normalizeTextKey_(row.requester);
    if (!name) {
      return;
    }
    if (!known.hasOwnProperty(name)) {
      known[name] = requesterOrder.names.length + namesToAppend.length + 1;
      namesToAppend.push(name);
    }
  });

  if (namesToAppend.length === 0) {
    return;
  }

  const startRow = Math.max(orderSheet.getLastRow() + 1, 2);
  const values = namesToAppend.map(function (name, index) {
    return [startRow + index - 1, name, ''];
  });
  orderSheet.getRange(startRow, 1, values.length, 3).setValues(values);
}

function insertRequestRow_(sheet, rowValues) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Another request is updating the sheet. Please try again.');
  }

  try {
    sheet.insertRowsBefore(2, 1);

    const rowRange = sheet.getRange(2, 1, 1, rowValues.length);
    rowRange.removeCheckboxes();
    rowRange.clearDataValidations();
    rowRange.setValues([rowValues]);

    const checkboxCell = sheet.getRange(2, rowValues.length);
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);
    return 2;
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  const existing = spreadsheet.getSheetByName(sheetName);
  if (existing) {
    return existing;
  }
  return spreadsheet.insertSheet(sheetName);
}

function setupDashboardLayout_(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);

  sheet.getRange('A1').setValue('ダッシュボード');
  sheet.getRange('A3').setValue('日付を選択');
  sheet.getRange('A4').setValue('表示対象');
  sheet.getRange('D3').setValue('対象日件数');
  sheet.getRange('D4').setValue('送り依頼');
  sheet.getRange('D5').setValue('備品注文');
  sheet.getRange('D6').setValue('商品希望');
  sheet.getRange('D7').setValue('未完了');

  sheet.getRange('A1:J1')
    .setBackground(DASHBOARD_TITLE_COLOR)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16);
  sheet.setRowHeight(1, 34);

  sheet.getRange('A3:A4')
    .setBackground(DASHBOARD_LABEL_COLOR)
    .setFontColor(DASHBOARD_TITLE_COLOR);
  sheet.getRange('D3:E7')
    .setBackground(DASHBOARD_SUMMARY_COLOR)
    .setFontColor(DASHBOARD_TITLE_COLOR);
  sheet.getRange('A3:A4').setFontWeight('bold');
  sheet.getRange('D3:D7').setFontWeight('bold');
  sheet.getRange('A3')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#AAB7C8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setNote('図形ボタンをこのセルの上に配置し、showDashboardCalendar を割り当てます。');

  const today = normalizeDateKey_(new Date());
  sheet.getRange('B3').setValue(formatDashboardDateSummary_([today]));
  sheet.getRange('B3').setNumberFormat('@');
  sheet.getRange('B3')
    .setHorizontalAlignment('left')
    .setNote('A3 をクリックして希望着日を複数選択');
  sheet.getRange('B4').setValue(DASHBOARD_STATUS_ALL);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([DASHBOARD_STATUS_ALL, DASHBOARD_STATUS_PENDING, DASHBOARD_STATUS_DONE], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B4').setDataValidation(statusRule);

  const headers = [getDashboardHeaders_()];
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, DASHBOARD_COLUMNS).setValues(headers);
  formatDashboardTableHeader_(sheet);

  writeDashboardStoredDates_(sheet, [today]);
  sheet.hideColumns(DASHBOARD_META_START_COLUMN, DASHBOARD_META_COLUMNS);
  sheet.hideColumns(DASHBOARD_DATE_STORE_COLUMN);
  setupDashboardEditWarnings_(sheet);
  sheet.setFrozenRows(DASHBOARD_HEADER_ROW);
  setDashboardColumnWidths_(sheet);
}

function getDashboardHeaders_() {
  return [
    '種別',
    '依頼者',
    '希望着日',
    '内容',
    '運送会社',
    '最低CT',
    '最高CT',
    '補足/画像',
    '確認',
    '済'
  ];
}

function ensureDashboardStaticLayout_(sheet) {
  sheet.getRange('D3:D7').setValues([
    ['対象日件数'],
    ['送り依頼'],
    ['備品注文'],
    ['商品希望'],
    ['未完了']
  ]);
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, DASHBOARD_COLUMNS).setValues([getDashboardHeaders_()]);
  sheet.getRange(DASHBOARD_HEADER_ROW, DASHBOARD_META_START_COLUMN, 1, DASHBOARD_META_COLUMNS).clearContent();
  formatDashboardTableHeader_(sheet);
  setDashboardColumnWidths_(sheet);
  sheet.hideColumns(DASHBOARD_META_START_COLUMN, DASHBOARD_META_COLUMNS);
  sheet.hideColumns(DASHBOARD_DATE_STORE_COLUMN);
}

function ensureRequesterOrderSheet_(sheet) {
  if (sheet.getLastRow() >= 1 && String(sheet.getRange(1, 1).getValue()).trim() === '並び順') {
    return;
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([['並び順', '依頼者', 'メモ']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
}

function getDashboardTargetDates_(dashboardSheet) {
  const values = dashboardSheet
    .getRange(DASHBOARD_DATE_STORE_START_ROW, DASHBOARD_DATE_STORE_COLUMN, DASHBOARD_DATE_STORE_MAX_ROWS, 1)
    .getValues()
    .map(function (row) { return normalizeDashboardStoredDateKey_(row[0]); })
    .filter(function (value) { return Boolean(value); });

  const unique = uniqueDateKeys_(values);
  if (unique.length > 0) {
    return unique;
  }

  const fallbackDates = getDashboardFallbackTargetDates_(dashboardSheet);
  if (fallbackDates.length > 0) {
    writeDashboardStoredDates_(dashboardSheet, fallbackDates);
    dashboardSheet.getRange('B3').setValue(formatDashboardDateSummary_(fallbackDates));
    return fallbackDates;
  }

  const today = normalizeDateKey_(new Date());
  writeDashboardStoredDates_(dashboardSheet, [today]);
  dashboardSheet.getRange('B3').setValue(formatDashboardDateSummary_([today]));
  return [today];
}

function uniqueDateKeys_(dateKeys) {
  const unique = [];
  const seen = {};
  dateKeys.forEach(function (value) {
    if (!seen.hasOwnProperty(value)) {
      seen[value] = true;
      unique.push(value);
    }
  });
  unique.sort();
  return unique;
}

function getDashboardFallbackTargetDates_(dashboardSheet) {
  const displayText = normalizeTextKey_(dashboardSheet.getRange('B3').getValue());
  const dateMatches = displayText.match(/\d{4}-\d{2}-\d{2}/g) || [];
  if (dateMatches.length >= 2 && displayText.indexOf('~') !== -1) {
    return expandDateRange_(dateMatches[0], dateMatches[1]);
  }
  return uniqueDateKeys_(dateMatches.map(function (dateKey) {
    return normalizeDateKey_(dateKey);
  }).filter(function (dateKey) {
    return Boolean(dateKey);
  }));
}

function expandDateRange_(startDateKey, endDateKey) {
  const start = parseDateKey_(startDateKey);
  const end = parseDateKey_(endDateKey);
  if (!start || !end || start.getTime() > end.getTime()) {
    return uniqueDateKeys_([normalizeDateKey_(startDateKey), normalizeDateKey_(endDateKey)]);
  }

  const dates = [];
  const current = new Date(start.getTime());
  while (current.getTime() <= end.getTime() && dates.length < DASHBOARD_DATE_STORE_MAX_ROWS) {
    dates.push(normalizeDateKey_(current));
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function parseDateKey_(dateKey) {
  const match = String(dateKey).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  if (!Number.isFinite(date.getTime())) {
    return null;
  }
  return date;
}

function getDashboardStatusFilter_(dashboardSheet) {
  const status = normalizeTextKey_(dashboardSheet.getRange('B4').getValue());
  if (status === DASHBOARD_STATUS_PENDING || status === DASHBOARD_STATUS_DONE) {
    return status;
  }
  return DASHBOARD_STATUS_ALL;
}

function readRequesterOrder_() {
  const orderSheet = getRequiredSheetByName_(REQUESTER_ORDER_SHEET_NAME);
  const lastRow = orderSheet.getLastRow();
  const orderMap = {};
  const names = [];

  if (lastRow < 2) {
    return { orderMap: orderMap, names: names };
  }

  const values = orderSheet.getRange(2, 2, lastRow - 1, 1).getValues();
  values.forEach(function (row) {
    const name = normalizeTextKey_(row[0]);
    if (!name || orderMap.hasOwnProperty(name)) {
      return;
    }
    orderMap[name] = names.length + 1;
    names.push(name);
  });
  return { orderMap: orderMap, names: names };
}

function readShippingRows_(sheetName) {
  const sheet = getExistingSheetByName_(sheetName);
  if (!sheet) {
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const rowCount = getSourceReadRowCount_(lastRow);
  const startRow = getSourceReadStartRow_(lastRow);
  const values = sheet.getRange(startRow, 1, rowCount, SHIPPING_DONE_COLUMN).getValues();
  return values.map(function (row, index) {
    return {
      sourceType: DASHBOARD_SOURCE_SHIPPING,
      requestId: row[0],
      timestamp: row[1],
      userId: row[2],
      requester: row[3],
      carrier: row[4],
      arrivalDate: row[5],
      summary: row[6],
      minCt: row[7],
      maxCt: row[8],
      supplement: '',
      reviewStatus: '',
      done: row[9],
      sourceSheetName: sheetName,
      sourceRow: startRow + index
    };
  });
}

function normalizeDashboardStoredDateKey_(value) {
  if (value instanceof Date) {
    return normalizeDateKey_(value);
  }
  const text = normalizeTextKey_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return '';
  }
  return parseDateKey_(text) ? text : '';
}

function readSupplyRows_(sheetName) {
  const sheet = getExistingSheetByName_(sheetName);
  if (!sheet) {
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const rowCount = getSourceReadRowCount_(lastRow);
  const startRow = getSourceReadStartRow_(lastRow);
  const values = sheet.getRange(startRow, 1, rowCount, SUPPLY_DONE_COLUMN).getValues();
  return values.map(function (row, index) {
    const doneValue = row[7] !== '' ? row[7] : row[5];
    return {
      sourceType: DASHBOARD_SOURCE_SUPPLY,
      requestId: row[0],
      timestamp: row[1],
      userId: row[2],
      requester: row[3],
      carrier: '',
      arrivalDate: row[4],
      summary: row[5],
      minCt: '',
      maxCt: '',
      supplement: row[6],
      reviewStatus: '',
      done: doneValue,
      sourceSheetName: sheetName,
      sourceRow: startRow + index
    };
  });
}

function readProductRequestRows_(sheetName) {
  const sheet = getExistingSheetByName_(sheetName);
  if (!sheet) {
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const rowCount = getSourceReadRowCount_(lastRow);
  const startRow = getSourceReadStartRow_(lastRow);
  const values = sheet.getRange(startRow, 1, rowCount, PRODUCT_DONE_COLUMN).getValues();
  const folderRichValues = sheet.getRange(startRow, 8, rowCount, 1).getRichTextValues();
  const imageRichValues = sheet.getRange(startRow, 9, rowCount, 1).getRichTextValues();
  return values.map(function (row, index) {
    const folderLinks = getRichTextLinks_(folderRichValues[index][0]);
    const folderUrls = folderLinks.length > 0 ? folderLinks : extractUrls_(row[7]);
    const imageUrls = getRichTextLinks_(imageRichValues[index][0]);
    const fallbackImageUrls = splitMultilineValue_(row[8]);
    const dashboardImageUrls = imageUrls.length > 0 ? imageUrls : fallbackImageUrls;
    const dashboardFolderUrl = folderUrls[0] || '';
    return {
      sourceType: DASHBOARD_SOURCE_PRODUCT,
      requestId: row[0],
      timestamp: row[1],
      userId: row[2],
      requester: row[3],
      carrier: '',
      arrivalDate: row[4],
      summary: row[5],
      minCt: '',
      maxCt: '',
      supplement: dashboardFolderUrl ? 'フォルダを開く' : formatImageLinkLabels_(dashboardImageUrls),
      reviewStatus: '',
      done: row[9],
      sourceSheetName: sheetName,
      sourceRow: startRow + index,
      dashboardFolderUrl: dashboardFolderUrl,
      dashboardImageUrls: dashboardImageUrls
    };
  });
}

function getExistingSheetByName_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Active spreadsheet is not available. Use a bound script.');
  }
  return spreadsheet.getSheetByName(sheetName);
}

function formatProductRequestSheetLinks() {
  const config = getAppConfig_();
  const sheet = getRequiredSheetByName_(config.productRequestSheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { rows: 0 };
  }

  const rowCount = lastRow - 1;
  const folderValues = sheet.getRange(2, 8, rowCount, 1).getValues();
  const imageValues = sheet.getRange(2, 9, rowCount, 1).getValues();
  const folderRichValues = sheet.getRange(2, 8, rowCount, 1).getRichTextValues();
  const imageRichValues = sheet.getRange(2, 9, rowCount, 1).getRichTextValues();

  for (let index = 0; index < rowCount; index += 1) {
    const folderLinks = getRichTextLinks_(folderRichValues[index][0]);
    const folderUrl = folderLinks[0] || extractUrls_(folderValues[index][0])[0] || '';
    const imageLinks = getRichTextLinks_(imageRichValues[index][0]);
    const imageUrls = imageLinks.length > 0 ? imageLinks : extractUrls_(imageValues[index][0]);
    applyProductRequestRowLinks_(sheet, index + 2, folderUrl, imageUrls);
  }

  return { rows: rowCount };
}

function applyProductRequestRowLinks_(sheet, row, folderUrl, imageUrls) {
  const normalizedFolderUrl = toOptionalString_(folderUrl, '');
  const normalizedImageUrls = Array.isArray(imageUrls) ? imageUrls.filter(function (url) {
    return Boolean(toOptionalString_(url, ''));
  }) : [];

  const folderCell = sheet.getRange(row, 8);
  if (normalizedFolderUrl) {
    folderCell.setRichTextValue(buildSingleLinkRichText_('フォルダを開く', normalizedFolderUrl));
  } else {
    folderCell.clearContent();
  }

  const imageCell = sheet.getRange(row, 9);
  if (normalizedImageUrls.length > 0) {
    imageCell.setRichTextValue(buildImageLinksRichText_(normalizedImageUrls));
  } else {
    imageCell.clearContent();
  }
}

function buildSingleLinkRichText_(text, url) {
  return SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(0, text.length, url)
    .build();
}

function buildImageLinksRichText_(imageUrls) {
  const labels = imageUrls.map(function (_url, index) {
    return '画像' + String(index + 1);
  });
  const text = labels.join('\n');
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  let offset = 0;
  imageUrls.forEach(function (url, index) {
    const label = labels[index];
    builder.setLinkUrl(offset, offset + label.length, url);
    offset += label.length + 1;
  });
  return builder.build();
}

function formatImageLinkLabels_(imageUrls) {
  return imageUrls.map(function (_url, index) {
    return '画像' + String(index + 1);
  }).join('\n');
}

function splitMultilineValue_(value) {
  const urls = extractUrls_(value);
  if (urls.length > 0) {
    return urls;
  }
  return String(value || '').split(/\r?\n/).map(function (item) {
    return item.trim();
  }).filter(function (item) {
    return Boolean(item);
  });
}

function extractUrls_(value) {
  const text = String(value || '');
  return text.match(/https?:\/\/[^\s]+/g) || [];
}

function getRichTextLinks_(richTextValue) {
  if (!richTextValue) {
    return [];
  }
  const links = [];
  const runs = richTextValue.getRuns();
  runs.forEach(function (run) {
    const url = run.getLinkUrl();
    if (url && links.indexOf(url) === -1) {
      links.push(url);
    }
  });
  return links;
}

function getSourceReadRowCount_(lastRow) {
  return Math.min(lastRow - 1, SOURCE_READ_MAX_ROWS);
}

function getSourceReadStartRow_(lastRow) {
  return Math.max(2, lastRow - getSourceReadRowCount_(lastRow) + 1);
}

function buildDashboardRows_(config, targetDates, statusFilter, requesterOrderMap) {
  const shippingRows = readShippingRows_(config.masterSheetName);
  const supplyRows = readSupplyRows_(config.supplySheetName);
  const productRows = readProductRequestRows_(config.productRequestSheetName);
  const allRows = shippingRows.concat(supplyRows).concat(productRows);
  const targetMap = {};
  targetDates.forEach(function (dateKey) {
    targetMap[dateKey] = true;
  });
  const summary = {
    total: 0,
    shipping: 0,
    supply: 0,
    product: 0,
    pending: 0
  };

  const filteredRows = allRows.filter(function (row) {
    return targetMap.hasOwnProperty(normalizeDateKey_(row.arrivalDate));
  }).filter(function (row) {
    const done = toDoneBoolean_(row.done);
    if (statusFilter === DASHBOARD_STATUS_PENDING) {
      return !done;
    }
    if (statusFilter === DASHBOARD_STATUS_DONE) {
      return done;
    }
    return true;
  });

  filteredRows.forEach(function (row) {
    summary.total += 1;
    if (row.sourceType === DASHBOARD_SOURCE_SHIPPING) {
      summary.shipping += 1;
    } else if (row.sourceType === DASHBOARD_SOURCE_SUPPLY) {
      summary.supply += 1;
    } else if (row.sourceType === DASHBOARD_SOURCE_PRODUCT) {
      summary.product += 1;
    }
    if (!toDoneBoolean_(row.done)) {
      summary.pending += 1;
    }
  });

  filteredRows.sort(function (left, right) {
    const leftName = normalizeTextKey_(left.requester);
    const rightName = normalizeTextKey_(right.requester);
    const leftOrder = requesterOrderMap.hasOwnProperty(leftName) ? requesterOrderMap[leftName] : Number.MAX_SAFE_INTEGER;
    const rightOrder = requesterOrderMap.hasOwnProperty(rightName) ? requesterOrderMap[rightName] : Number.MAX_SAFE_INTEGER;

    if (leftOrder !== rightOrder) {
      return leftOrder - rightOrder;
    }
    if (leftName !== rightName) {
      return leftName.localeCompare(rightName, 'ja');
    }
    const leftSourceOrder = getDashboardSourceSortOrder_(left.sourceType);
    const rightSourceOrder = getDashboardSourceSortOrder_(right.sourceType);
    if (leftSourceOrder !== rightSourceOrder) {
      return leftSourceOrder - rightSourceOrder;
    }

    return toComparableTime_(left.timestamp) - toComparableTime_(right.timestamp);
  });

  const dashboardRows = filteredRows.map(function (row) {
    return [
      row.sourceType,
      toOptionalString_(row.requester, ''),
      normalizeDateKey_(row.arrivalDate),
      toOptionalString_(row.summary, ''),
      toOptionalString_(row.carrier, ''),
      toOptionalString_(row.minCt, ''),
      toOptionalString_(row.maxCt, ''),
      toOptionalString_(row.supplement, ''),
      toOptionalString_(row.reviewStatus, ''),
      toDoneBoolean_(row.done),
      row.sourceSheetName,
      row.sourceRow,
      row.dashboardFolderUrl || '',
      row.dashboardImageUrls || []
    ];
  });

  return { rows: dashboardRows, summary: summary };
}

function getDashboardSourceSortOrder_(sourceType) {
  if (sourceType === DASHBOARD_SOURCE_SHIPPING) return 1;
  if (sourceType === DASHBOARD_SOURCE_SUPPLY) return 2;
  if (sourceType === DASHBOARD_SOURCE_PRODUCT) return 3;
  return 99;
}

function writeDashboardRows_(sheet, rows) {
  removeDashboardFilter_(sheet);
  removeDashboardBandings_(sheet);

  const clearRows = getDashboardWritableRowCount_(sheet);
  const dataRange = sheet.getRange(DASHBOARD_DATA_START_ROW, 1, clearRows, DASHBOARD_COLUMNS);
  dataRange.clearContent();
  dataRange.clearFormat();
  const metaRange = sheet.getRange(DASHBOARD_DATA_START_ROW, DASHBOARD_META_START_COLUMN, clearRows, DASHBOARD_META_COLUMNS);
  metaRange.clearContent();
  sheet.getRange(DASHBOARD_DATA_START_ROW, DASHBOARD_DONE_COLUMN, clearRows, 1).removeCheckboxes();

  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, DASHBOARD_COLUMNS).setValues([getDashboardHeaders_()]);
  formatDashboardTableHeader_(sheet);

  const rowsToRender = rows.slice(0, DASHBOARD_RENDER_MAX_ROWS);
  if (rowsToRender.length === 0) {
    applyDashboardFilter_(sheet, 0);
    return;
  }

  const visibleRows = rowsToRender.map(function (row) {
    return row.slice(0, DASHBOARD_COLUMNS);
  });
  const metaRows = rowsToRender.map(function (row) {
    return row.slice(DASHBOARD_COLUMNS, DASHBOARD_COLUMNS + DASHBOARD_META_COLUMNS);
  });

  sheet.getRange(DASHBOARD_DATA_START_ROW, 1, visibleRows.length, DASHBOARD_COLUMNS).setValues(visibleRows);
  applyDashboardProductLinks_(sheet, rowsToRender);
  sheet.getRange(DASHBOARD_DATA_START_ROW, DASHBOARD_META_START_COLUMN, metaRows.length, DASHBOARD_META_COLUMNS).setValues(metaRows);
  sheet.getRange(DASHBOARD_DATA_START_ROW, 3, rowsToRender.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(DASHBOARD_DATA_START_ROW, DASHBOARD_DONE_COLUMN, rowsToRender.length, 1).insertCheckboxes();
  sheet.getRange(DASHBOARD_DATA_START_ROW, 1, rowsToRender.length, DASHBOARD_COLUMNS).setWrap(true);
  sheet.hideColumns(DASHBOARD_META_START_COLUMN, DASHBOARD_META_COLUMNS);
  sheet.hideColumns(DASHBOARD_DATE_STORE_COLUMN);
  formatDashboardDataRows_(sheet, rowsToRender.length);
  applyDashboardFilter_(sheet, rowsToRender.length);
}

function applyDashboardProductLinks_(sheet, rowsToRender) {
  rowsToRender.forEach(function (row, index) {
    const folderUrl = toOptionalString_(row[12], '');
    if (folderUrl) {
      sheet.getRange(DASHBOARD_DATA_START_ROW + index, 8).setRichTextValue(buildSingleLinkRichText_('フォルダを開く', folderUrl));
    }
  });
}

function writeDashboardSummary_(sheet, summary) {
  sheet.getRange('E3').setValue(summary.total);
  sheet.getRange('E4').setValue(summary.shipping);
  sheet.getRange('E5').setValue(summary.supply);
  sheet.getRange('E6').setValue(summary.product);
  sheet.getRange('E7').setValue(summary.pending);
}

function formatDashboardTableHeader_(sheet) {
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, DASHBOARD_COLUMNS)
    .setBackground(DASHBOARD_HEADER_COLOR)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(DASHBOARD_HEADER_ROW, 28);
}

function formatDashboardDataRows_(sheet, rowCount) {
  const tableRange = sheet.getRange(DASHBOARD_HEADER_ROW, 1, rowCount + 1, DASHBOARD_COLUMNS);
  tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  formatDashboardTableHeader_(sheet);

  sheet.getRange(DASHBOARD_DATA_START_ROW, 1, rowCount, DASHBOARD_COLUMNS)
    .setVerticalAlignment('top')
    .setBorder(true, true, true, true, true, true, '#D9E2EF', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(DASHBOARD_DATA_START_ROW, 1, rowCount, 1).setFontWeight('bold');
  sheet.getRange(DASHBOARD_DATA_START_ROW, 6, rowCount, 2).setHorizontalAlignment('center');
  sheet.getRange(DASHBOARD_DATA_START_ROW, 9, rowCount, 2).setHorizontalAlignment('center');
}

function setDashboardColumnWidths_(sheet) {
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 105);
  sheet.setColumnWidth(4, 320);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 70);
  sheet.setColumnWidth(7, 70);
  sheet.setColumnWidth(8, 260);
  sheet.setColumnWidth(9, 95);
  sheet.setColumnWidth(10, 55);
}

function getDashboardWritableRowCount_(sheet) {
  return Math.min(
    DASHBOARD_RENDER_MAX_ROWS,
    Math.max(sheet.getMaxRows() - DASHBOARD_DATA_START_ROW + 1, 1)
  );
}

function applyDashboardFilter_(sheet, rowCount) {
  const filterRowCount = Math.max(rowCount + 1, 1);
  sheet.getRange(DASHBOARD_HEADER_ROW, 1, filterRowCount, DASHBOARD_COLUMNS).createFilter();
}

function removeDashboardFilter_(sheet) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
}

function removeDashboardBandings_(sheet) {
  sheet.getBandings().forEach(function (banding) {
    banding.remove();
  });
}

function normalizeDateKey_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, getSpreadsheetTimeZone_(), 'yyyy-MM-dd');
  }
  const text = normalizeTextKey_(value);
  if (!text) {
    return '';
  }

  const parsed = new Date(text);
  if (Number.isFinite(parsed.getTime())) {
    return Utilities.formatDate(parsed, getSpreadsheetTimeZone_(), 'yyyy-MM-dd');
  }
  return text;
}

function writeDashboardStoredDates_(sheet, dateKeys) {
  const values = [];
  const seen = {};
  dateKeys.forEach(function (dateKey) {
    const normalized = normalizeDateKey_(dateKey);
    if (!parseDateKey_(normalized) || seen.hasOwnProperty(normalized)) {
      return;
    }
    seen[normalized] = true;
    values.push([normalized]);
  });

  const storeRange = sheet.getRange(
    DASHBOARD_DATE_STORE_START_ROW,
    DASHBOARD_DATE_STORE_COLUMN,
    DASHBOARD_DATE_STORE_MAX_ROWS,
    1
  );
  storeRange.clearContent();
  if (values.length > 0) {
    sheet.getRange(
      DASHBOARD_DATE_STORE_START_ROW,
      DASHBOARD_DATE_STORE_COLUMN,
      values.length,
      1
    ).setValues(values);
  }
}

function setupDashboardEditWarnings_(sheet) {
  removeDashboardProtectionByDescription_(sheet, DASHBOARD_BUTTON_PROTECTION_DESCRIPTION);
  removeDashboardProtectionByDescription_(sheet, DASHBOARD_DATE_STORE_PROTECTION_DESCRIPTION);

  sheet.getRange('A3:B3')
    .protect()
    .setDescription(DASHBOARD_BUTTON_PROTECTION_DESCRIPTION)
    .setWarningOnly(true);

  sheet.getRange(DASHBOARD_DATE_STORE_START_ROW, DASHBOARD_DATE_STORE_COLUMN, DASHBOARD_DATE_STORE_MAX_ROWS, 1)
    .protect()
    .setDescription(DASHBOARD_DATE_STORE_PROTECTION_DESCRIPTION)
    .setWarningOnly(true);
}

function removeDashboardProtectionByDescription_(sheet, description) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (protection) {
    if (protection.getDescription() === description) {
      protection.remove();
    }
  });
}

function getDashboardStoredDatesForDialog_() {
  const sheet = getRequiredSheetByName_(DASHBOARD_SHEET_NAME);
  ensureDashboardStaticLayout_(sheet);
  return getDashboardTargetDates_(sheet);
}

function saveDashboardSelectedDates(dateKeys) {
  if (!Array.isArray(dateKeys)) {
    throw new Error('dateKeys must be an array.');
  }
  if (dateKeys.length === 0) {
    throw new Error('日付を1件以上選択してください。');
  }

  const sheet = getRequiredSheetByName_(DASHBOARD_SHEET_NAME);
  writeDashboardStoredDates_(sheet, dateKeys);
  const selectedDates = getDashboardTargetDates_(sheet);
  sheet.getRange('B3').setValue(formatDashboardDateSummary_(selectedDates));
  refreshDashboard();
}

function formatDashboardDateSummary_(dateKeys) {
  if (!dateKeys || dateKeys.length === 0) {
    return '';
  }
  const sorted = dateKeys.slice().sort();
  if (sorted.length === 1) {
    return sorted[0] + ' (1日)';
  }
  return sorted[0] + ' ~ ' + sorted[sorted.length - 1] + ' (' + sorted.length + '日)';
}

function showDashboardCalendar() {
  const template = HtmlService.createTemplateFromFile('dashboard_calendar');
  template.initialDatesJson = JSON.stringify(getDashboardStoredDatesForDialog_());
  const html = template.evaluate()
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '希望着日を選択');
}

function getSpreadsheetTimeZone_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet ? spreadsheet.getSpreadsheetTimeZone() : 'Asia/Tokyo';
}

function toDoneBoolean_(value) {
  if (value === true) {
    return true;
  }
  return String(value).toUpperCase() === 'TRUE';
}

function toComparableTime_(value) {
  if (value instanceof Date) {
    return value.getTime();
  }
  const parsed = new Date(value);
  if (Number.isFinite(parsed.getTime())) {
    return parsed.getTime();
  }
  return Number.MAX_SAFE_INTEGER;
}

function normalizeTextKey_(value) {
  return String(value === undefined || value === null ? '' : value).trim();
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

function pushLineTextMessage_(config, messageText) {
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + config.lineToken
    },
    payload: JSON.stringify({
      to: config.targetUserId,
      messages: [
        {
          type: 'text',
          text: messageText
        }
      ]
    }),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(
      'LINE push request failed (' + statusCode + '): ' + response.getContentText()
    );
  }
}

function normalizeShippingPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const normalized = {
    userId: toOptionalString_(data.userId, ''),
    userName: toOptionalString_(data.userName, '未入力'),
    carrier: toOptionalString_(data.carrier, '未選択'),
    arrivalDate: toRequiredString_(data.arrivalDate, 'arrivalDate'),
    destination: toRequiredString_(data.destination, 'destination'),
    minCt: toNonNegativeInteger_(data.minCt, 'minCt'),
    maxCt: toNonNegativeInteger_(data.maxCt, 'maxCt')
  };

  if (normalized.minCt > normalized.maxCt) {
    throw new Error('minCt must be less than or equal to maxCt.');
  }

  return normalized;
}

function normalizeOrderPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const userId = toOptionalString_(data.userId, '');
  const userName = toOptionalString_(data.userName, '未入力');
  const arrivalDate = toRequiredString_(data.arrivalDate, 'arrivalDate');
  const freeNote = toOptionalString_(data.freeNote, '');
  if (!Array.isArray(data.items)) {
    throw new Error('items must be an array.');
  }

  const items = data.items.map(function (item, index) {
    const row = ensureObject_(item, 'items[' + index + ']');
    const itemNo = row.itemNo === undefined || row.itemNo === null || row.itemNo === ''
      ? parseLeadingItemNo_(row.name)
      : toPositiveInteger_(row.itemNo, 'items[' + index + '].itemNo');
    return {
      itemNo: itemNo,
      name: toRequiredString_(row.name, 'items[' + index + '].name'),
      qty: toNonNegativeInteger_(row.qty, 'items[' + index + '].qty')
    };
  }).filter(function (item) {
    return item.qty > 0;
  });

  if (items.length === 0 && !freeNote) {
    throw new Error('Either items with qty > 0 or freeNote is required.');
  }

  return {
    userId: userId,
    userName: userName,
    arrivalDate: arrivalDate,
    items: items,
    freeNote: freeNote
  };
}

function normalizeProductRequestPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const userId = toOptionalString_(data.userId, '');
  const userName = toOptionalString_(data.userName, '未入力');
  const arrivalDate = toRequiredString_(data.arrivalDate, 'arrivalDate');
  const requestText = toOptionalString_(data.requestText, '');
  const images = normalizeProductImages_(data.images);

  if (!requestText && images.length === 0) {
    throw new Error('requestText or images is required.');
  }

  return {
    userId: userId,
    userName: userName,
    arrivalDate: arrivalDate,
    requestText: requestText,
    images: images
  };
}

function safePushLineTextMessage_(config, messageText) {
  try {
    pushLineTextMessage_(config, messageText);
    return { ok: true };
  } catch (error) {
    Logger.log('LINE notification failed: ' + toErrorMessage_(error));
    return {
      ok: false,
      error: toErrorMessage_(error)
    };
  }
}

function normalizeProductImages_(value) {
  if (value === undefined || value === null || value === '') {
    return [];
  }
  if (!Array.isArray(value)) {
    throw new Error('images must be an array.');
  }
  return value.map(function (item, index) {
    const image = ensureObject_(item, 'images[' + index + ']');
    const contentType = toRequiredString_(image.contentType, 'images[' + index + '].contentType');
    if (contentType.indexOf('image/') !== 0) {
      throw new Error('images[' + index + '].contentType must start with image/.');
    }
    const base64Data = toRequiredString_(image.base64Data, 'images[' + index + '].base64Data');
    return {
      contentType: contentType,
      base64Data: base64Data,
      fileName: toOptionalString_(image.fileName, 'image_' + (index + 1) + '.jpg')
    };
  }).slice(0, 5);
}

function saveProductRequestImages_(requestId, requestData, config) {
  if (!requestData.images || requestData.images.length === 0) {
    return {
      imageCount: 0,
      imageUrls: [],
      folderUrl: ''
    };
  }

  const rootFolder = getProductRequestRootFolder_(config);
  const dateFolder = getOrCreateChildFolder_(rootFolder, normalizeDateKey_(requestData.arrivalDate));
  const userFolder = getOrCreateChildFolder_(dateFolder, sanitizeDriveName_(requestData.userName));
  const folder = userFolder.createFolder(requestId);
  const timestamp = Utilities.formatDate(new Date(), getSpreadsheetTimeZone_(), 'yyyyMMdd-HHmmss');
  const imageUrls = requestData.images.map(function (image, index) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(image.base64Data),
      image.contentType,
      buildProductImageFileName_(timestamp, index, image.contentType)
    );
    const file = folder.createFile(blob);
    return file.getUrl();
  });

  return {
    imageCount: imageUrls.length,
    imageUrls: imageUrls,
    folderUrl: folder.getUrl()
  };
}

function getProductRequestRootFolder_(config) {
  if (config.productRequestImageFolderId) {
    return DriveApp.getFolderById(config.productRequestImageFolderId);
  }

  const folderName = '商品希望画像';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function cleanupOldProductRequestImageFolders() {
  const config = getAppConfig_();
  const rootFolder = getProductRequestRootFolder_(config);
  const trashThreshold = getCleanupThresholdDate_(PRODUCT_IMAGE_TRASH_DAYS);
  const deleteThreshold = getCleanupThresholdDate_(PRODUCT_IMAGE_DELETE_DAYS);

  const folders = rootFolder.getFolders();
  let trashedCount = 0;
  while (folders.hasNext()) {
    const folder = folders.next();
    const folderDate = parseDateKey_(folder.getName());
    if (!folderDate) {
      continue;
    }
    folderDate.setHours(0, 0, 0, 0);
    if (folderDate.getTime() < trashThreshold.getTime()) {
      folder.setTrashed(true);
      trashedCount += 1;
    }
  }

  const deletedCount = deleteTrashedProductRequestDateFolders_(rootFolder.getId(), deleteThreshold);
  const result = {
    trashed: trashedCount,
    deleted: deletedCount
  };
  Logger.log('Product image cleanup result: ' + JSON.stringify(result));
  return result;
}

function getCleanupThresholdDate_(days) {
  const threshold = new Date();
  threshold.setHours(0, 0, 0, 0);
  threshold.setDate(threshold.getDate() - days);
  return threshold;
}

function deleteTrashedProductRequestDateFolders_(rootFolderId, deleteThreshold) {
  let deletedCount = 0;
  let pageToken = null;
  const query = [
    "'" + rootFolderId + "' in parents",
    "mimeType = 'application/vnd.google-apps.folder'",
    "trashed = true"
  ].join(' and ');

  do {
    const response = Drive.Files.list({
      q: query,
      maxResults: 100,
      pageToken: pageToken
    });
    const items = response.items || [];
    items.forEach(function (item) {
      const folderDate = parseDateKey_(item.title);
      if (!folderDate) {
        return;
      }
      folderDate.setHours(0, 0, 0, 0);
      if (folderDate.getTime() < deleteThreshold.getTime()) {
        Drive.Files.remove(item.id);
        deletedCount += 1;
      }
    });
    pageToken = response.nextPageToken;
  } while (pageToken);

  return deletedCount;
}

function installProductImageCleanupTrigger() {
  removeProductImageCleanupTriggers_();
  ScriptApp.newTrigger('cleanupOldProductRequestImageFolders')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

function removeProductImageCleanupTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'cleanupOldProductRequestImageFolders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function getOrCreateChildFolder_(parentFolder, folderName) {
  const safeName = sanitizeDriveName_(folderName);
  const folders = parentFolder.getFoldersByName(safeName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(safeName);
}

function buildProductImageFileName_(timestamp, index, contentType) {
  const ext = inferImageExtension_(contentType);
  return timestamp + '_' + String(index + 1).padStart(2, '0') + '.' + ext;
}

function inferImageExtension_(contentType) {
  if (contentType === 'image/png') return 'png';
  if (contentType === 'image/webp') return 'webp';
  return 'jpg';
}

function sanitizeDriveName_(text) {
  return String(text || '').replace(/[\\/:*?"<>|]/g, '_').trim() || 'unknown';
}

function toShippingRow_(shippingData) {
  return [
    createRequestId_('SHIP'),
    new Date(),
    shippingData.userId,
    shippingData.userName,
    shippingData.carrier,
    shippingData.arrivalDate,
    shippingData.destination,
    shippingData.minCt,
    shippingData.maxCt,
    false
  ];
}

function toSupplyOrderRow_(orderData) {
  return [
    createRequestId_('SUP'),
    new Date(),
    orderData.userId,
    orderData.userName,
    orderData.arrivalDate,
    formatOrderItemsForSheet_(orderData.items),
    orderData.freeNote,
    false
  ];
}

function toProductRequestRow_(requestData) {
  const imageUrls = requestData.imageUrls || [];
  return [
    requestData.requestId,
    new Date(),
    requestData.userId,
    requestData.userName,
    requestData.arrivalDate,
    requestData.requestText,
    imageUrls.length,
    requestData.imageFolderUrl || '',
    imageUrls.join('\n'),
    false
  ];
}

function createRequestId_(prefix) {
  const timestamp = Utilities.formatDate(new Date(), getSpreadsheetTimeZone_(), 'yyyyMMdd-HHmmss');
  const randomPart = String(Math.floor(Math.random() * 10000)).padStart(4, '0');
  return prefix + '-' + timestamp + '-' + randomPart;
}

function normalizeHistoryRequest_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const status = normalizeHistoryStatus_(data.status);
  return {
    userId: toRequiredString_(data.userId, 'userId'),
    status: status,
    limit: toHistoryLimit_(data.limit),
    offset: toHistoryOffset_(data.offset)
  };
}

function toHistoryLimit_(value) {
  if (value === undefined || value === null || value === '') {
    return HISTORY_DEFAULT_LIMIT;
  }
  const limit = Number(value);
  if (!Number.isFinite(limit) || Math.floor(limit) !== limit || limit <= 0) {
    throw new Error('limit must be a positive integer.');
  }
  return Math.min(limit, HISTORY_MAX_LIMIT);
}

function toHistoryOffset_(value) {
  if (value === undefined || value === null || value === '') {
    return 0;
  }
  const offset = Number(value);
  if (!Number.isFinite(offset) || Math.floor(offset) !== offset || offset < 0) {
    throw new Error('offset must be a non-negative integer.');
  }
  return offset;
}

function normalizeHistoryStatus_(value) {
  const status = normalizeTextKey_(value).toLowerCase();
  if (status === 'pending' || status === 'done') {
    return status;
  }
  return 'all';
}

function normalizeDeleteRequestPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  return {
    userId: toRequiredString_(data.userId, 'userId'),
    requestId: toRequiredString_(data.requestId, 'requestId')
  };
}

function deleteUserRequest_(deleteRequest, config) {
  const targets = [
    { type: DASHBOARD_SOURCE_SHIPPING, sheetName: config.masterSheetName, userColumn: 3, folderColumn: 0 },
    { type: DASHBOARD_SOURCE_SUPPLY, sheetName: config.supplySheetName, userColumn: 3, folderColumn: 0 },
    { type: DASHBOARD_SOURCE_PRODUCT, sheetName: config.productRequestSheetName, userColumn: 3, folderColumn: 8 }
  ];

  for (let index = 0; index < targets.length; index += 1) {
    const target = targets[index];
    const sheet = getExistingSheetByName_(target.sheetName);
    if (!sheet) {
      continue;
    }
    const found = findRequestRow_(sheet, deleteRequest.requestId, target.userColumn);
    if (!found) {
      continue;
    }
    if (found.userId !== deleteRequest.userId) {
      throw new Error('この依頼は削除できません。');
    }
    const folderUrl = target.folderColumn > 0 ? getCellFirstLinkOrText_(sheet.getRange(found.row, target.folderColumn)) : '';
    if (target.type === DASHBOARD_SOURCE_PRODUCT && folderUrl) {
      trashDriveFolderByUrl_(folderUrl);
    }
    sheet.deleteRow(found.row);
    return { deleted: true, type: target.type };
  }

  throw new Error('削除対象の依頼が見つかりません。');
}

function findRequestRow_(sheet, requestId, userColumn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  const values = sheet.getRange(2, 1, lastRow - 1, Math.max(userColumn, 1)).getValues();
  for (let index = 0; index < values.length; index += 1) {
    if (toOptionalString_(values[index][0], '') === requestId) {
      return {
        row: index + 2,
        userId: toOptionalString_(values[index][userColumn - 1], '')
      };
    }
  }
  return null;
}

function getCellFirstLinkOrText_(range) {
  const links = getRichTextLinks_(range.getRichTextValue());
  if (links.length > 0) {
    return links[0];
  }
  return toOptionalString_(range.getValue(), '');
}

function trashDriveFolderByUrl_(url) {
  const folderId = extractDriveFolderId_(url);
  if (!folderId) {
    return;
  }
  try {
    DriveApp.getFolderById(folderId).setTrashed(true);
  } catch (error) {
    Logger.log('Product image folder trash failed: ' + toErrorMessage_(error));
  }
}

function extractDriveFolderId_(url) {
  const text = toOptionalString_(url, '');
  const match = text.match(/\/folders\/([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : '';
}

function getHistoryItems_(requestData, config) {
  const shippingItems = readShippingRows_(config.masterSheetName).map(function (row) {
    return {
      requestId: toOptionalString_(row.requestId, ''),
      userId: toOptionalString_(row.userId, ''),
      type: DASHBOARD_SOURCE_SHIPPING,
      requestedAt: row.timestamp,
      arrivalDate: row.arrivalDate,
      summary: row.summary,
      detailText: [
        toOptionalString_(row.carrier, '未選択'),
        toOptionalString_(row.minCt, '') + '-' + toOptionalString_(row.maxCt, '') + 'CT'
      ].join(' / '),
      done: toDoneBoolean_(row.done)
    };
  });
  const supplyItems = readSupplyRows_(config.supplySheetName).map(function (row) {
    return {
      requestId: toOptionalString_(row.requestId, ''),
      userId: toOptionalString_(row.userId, ''),
      type: DASHBOARD_SOURCE_SUPPLY,
      requestedAt: row.timestamp,
      arrivalDate: row.arrivalDate,
      summary: row.summary,
      detailText: '自由記入: ' + toOptionalString_(row.supplement, '（なし）'),
      done: toDoneBoolean_(row.done)
    };
  });
  const productItems = readProductRequestRows_(config.productRequestSheetName).map(function (row) {
    return {
      requestId: toOptionalString_(row.requestId, ''),
      userId: toOptionalString_(row.userId, ''),
      type: DASHBOARD_SOURCE_PRODUCT,
      requestedAt: row.timestamp,
      arrivalDate: row.arrivalDate,
      summary: row.summary,
      detailText: toOptionalString_(row.supplement, '画像なし'),
      done: toDoneBoolean_(row.done)
    };
  });

  const matchedItems = shippingItems.concat(supplyItems).concat(productItems)
    .filter(function (item) {
      return item.userId && item.userId === requestData.userId;
    })
    .filter(function (item) {
      if (requestData.status === 'pending') {
        return !item.done;
      }
      if (requestData.status === 'done') {
        return item.done;
      }
      return true;
    })
    .sort(function (left, right) {
      return toComparableTime_(right.requestedAt) - toComparableTime_(left.requestedAt);
    });
  const start = requestData.offset;
  const pageItems = matchedItems.slice(start, start + requestData.limit);
  return {
    items: pageItems.map(function (item) {
      return {
        requestId: item.requestId,
        type: item.type,
        requestedAt: toIsoString_(item.requestedAt),
        arrivalDate: normalizeDateKey_(item.arrivalDate),
        summary: toOptionalString_(item.summary, ''),
        detailText: item.detailText,
        done: item.done
      };
    }),
    hasMore: matchedItems.length > start + pageItems.length,
    nextOffset: start + pageItems.length
  };
}

function toIsoString_(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return value.toISOString();
  }
  const parsed = new Date(value);
  if (Number.isFinite(parsed.getTime())) {
    return parsed.toISOString();
  }
  return '';
}

function formatOrderItemsForSheet_(items) {
  const formattedItems = items.map(function (item) {
    return formatOrderItemLabel_(item) + ' x ' + item.qty;
  });
  const lines = [];

  for (let index = 0; index < formattedItems.length; index += 5) {
    lines.push(formattedItems.slice(index, index + 5).join(' / '));
  }

  return lines.join('\n');
}

function formatOrderItemsForNotification_(items) {
  return items.map(function (item) {
    return formatOrderItemLabel_(item) + ' x ' + item.qty;
  }).join('\n');
}

function formatOrderItemLabel_(item) {
  const name = stripNumber_(item.name);
  if (item.itemNo && Number.isFinite(item.itemNo)) {
    return item.itemNo + '.' + name;
  }
  return name;
}

function stripNumber_(name) {
  return String(name).replace(/^\s*\d+\.\s*/, '');
}

function parseLeadingItemNo_(name) {
  const match = String(name).match(/^\s*(\d+)\./);
  if (!match) {
    return null;
  }
  return Number(match[1]);
}

function buildShippingNotificationText_(shippingData) {
  return [
    '🚚送り依頼が届きました！',
    '依頼者: ' + shippingData.userName,
    '運送会社: ' + shippingData.carrier,
    '希望着日: ' + shippingData.arrivalDate,
    '送り先: ' + shippingData.destination,
    '最低カートン: ' + shippingData.minCt,
    '最高カートン: ' + shippingData.maxCt
  ].join('\n');
}

function buildOrderNotificationText_(orderData) {
  const orderItemsText = orderData.items.length > 0
    ? formatOrderItemsForNotification_(orderData.items)
    : '（なし）';
  const freeNoteText = orderData.freeNote || '（なし）';

  return [
    '📦️備品注文が届きました！',
    '依頼者: ' + orderData.userName,
    '希望着日: ' + orderData.arrivalDate,
    '注文内容:',
    orderItemsText,
    '自由記入: ' + freeNoteText
  ].join('\n');
}

function buildProductRequestNotificationText_(requestData) {
  const body = requestData.requestText || '（本文なし）';
  return [
    '🧩商品希望が届きました',
    '依頼者: ' + requestData.userName,
    '希望着日: ' + requestData.arrivalDate,
    '画像枚数: ' + String((requestData.imageUrls || []).length),
    '希望内容:',
    body
  ].join('\n');
}

function createJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureObject_(value, fieldName) {
  if (!value || Object.prototype.toString.call(value) !== '[object Object]') {
    throw new Error(fieldName + ' must be an object.');
  }
  return value;
}

function toRequiredString_(value, fieldName) {
  const text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) {
    throw new Error(fieldName + ' is required.');
  }
  return text;
}

function toOptionalString_(value, defaultValue) {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || defaultValue;
}

function toNonNegativeInteger_(value, fieldName) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue) || numberValue < 0 || Math.floor(numberValue) !== numberValue) {
    throw new Error(fieldName + ' must be a non-negative integer.');
  }
  return numberValue;
}

function toPositiveInteger_(value, fieldName) {
  const numberValue = Number(value);
  if (!Number.isFinite(numberValue) || numberValue <= 0 || Math.floor(numberValue) !== numberValue) {
    throw new Error(fieldName + ' must be a positive integer.');
  }
  return numberValue;
}

function toErrorMessage_(error) {
  if (error && error.message) {
    return error.message;
  }
  return String(error);
}
