function getAppConfig_() {
  return {
    liffId: getRequiredScriptProperty_('LIFF_ID'),
    lineToken: getRequiredScriptProperty_('LINE_TOKEN'),
    masterSheetName: getRequiredScriptProperty_('MASTER_SHEET_NAME'),
    supplySheetName: getRequiredScriptProperty_('SUPPLY_SHEET_NAME'),
    targetUserId: getRequiredScriptProperty_('TARGET_USER_ID')
  };
}

function getRequiredScriptProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Missing script property: ' + key);
  }
  return value;
}

function appendShippingToSheet_(shippingData, config) {
  const sheet = getRequiredSheetByName_(config.masterSheetName);
  insertRequestRow_(sheet, toShippingRow_(shippingData));
}

function appendSupplyOrderToSheet_(orderData, config) {
  const sheet = getRequiredSheetByName_(config.supplySheetName);
  insertRequestRow_(sheet, toSupplyOrderRow_(orderData));
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
  } finally {
    lock.releaseLock();
  }
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

// Reply to a webhook event using replyToken
function pushLineReplyMessage_(config, replyToken, messageText) {
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + config.lineToken
    },
    payload: JSON.stringify({
      replyToken: replyToken,
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
    throw new Error('LINE reply request failed (' + statusCode + '): ' + response.getContentText());
  }
}

function normalizeShippingPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const normalized = {
    userName: toOptionalString_(data.userName, '未入力'),
    carrier: toOptionalString_(data.carrier, '未選択'),
    arrivalDate: toRequiredString_(data.arrivalDate, 'arrivalDate'),
    destination: toRequiredString_(data.destination, 'destination'),
    minCt: toNonNegativeInteger_(data.minCt, 'minCt'),
    maxCt: toNonNegativeInteger_(data.maxCt, 'maxCt'),
    hasSupplies: toOptionalString_(data.hasSupplies, '無'),
    hasRemaining: toOptionalString_(data.hasRemaining, '無')
  };

  return normalized;
}

function normalizeOrderPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const userName = toOptionalString_(data.userName, '未入力');
  const freeNote = toOptionalString_(data.freeNote, '');
  if (!Array.isArray(data.items)) {
    throw new Error('items must be an array.');
  }

  const items = data.items.map(function (item, index) {
    const row = ensureObject_(item, 'items[' + index + ']');
    return {
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
    userName: userName,
    items: items,
    freeNote: freeNote
  };
}

function toShippingRow_(shippingData) {
  return [
    new Date(),
    shippingData.userName,
    shippingData.carrier,
    shippingData.arrivalDate,
    shippingData.destination,
    shippingData.minCt,
    shippingData.maxCt,
    shippingData.hasSupplies,
    shippingData.hasRemaining,
    false
  ];
}

function toSupplyOrderRow_(orderData) {
  return [
    new Date(),
    orderData.userName,
    formatOrderItemsForSheet_(orderData.items),
    orderData.freeNote,
    false
  ];
}

function formatOrderItemsForSheet_(items) {
  const formattedItems = items.map(function (item) {
    const name = stripNumber_(item.name);
    return name + ' x ' + item.qty;
  });
  const lines = [];

  for (let index = 0; index < formattedItems.length; index += 5) {
    lines.push(formattedItems.slice(index, index + 5).join(' / '));
  }

  return lines.join('\n');
}

function formatOrderItemsForNotification_(items) {
  return items.map(function (item) {
    const name = stripNumber_(item.name);
    return name + ' x ' + item.qty;
  }).join('\n');
}

function stripNumber_(name) {
  return String(name).replace(/^\s*\d+\.\s*/, '');
}

function buildShippingNotificationText_(shippingData) {
  return [
    '🚚送り依頼が届きました！',
    '依頼者: ' + shippingData.userName,
    '運送会社: ' + shippingData.carrier,
    '希望着日: ' + shippingData.arrivalDate,
    '送り先: ' + shippingData.destination,
    '最低カートン: ' + shippingData.minCt,
    '最高カートン: ' + shippingData.maxCt,
    '備品注文: ' + shippingData.hasSupplies,
    '残から: ' + shippingData.hasRemaining
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
    '注文内容:',
    orderItemsText,
    '自由記入: ' + freeNoteText
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

function toErrorMessage_(error) {
  if (error && error.message) {
    return error.message;
  }
  return String(error);
}
