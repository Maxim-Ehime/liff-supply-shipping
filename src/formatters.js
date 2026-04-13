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

  if (normalized.minCt > normalized.maxCt) {
    throw new Error('minCt must be less than or equal to maxCt.');
  }

  return normalized;
}

function normalizeOrderPayload_(rawData) {
  const data = ensureObject_(rawData, 'data');
  const userName = toOptionalString_(data.userName, '未入力');
  if (!Array.isArray(data.items)) {
    throw new Error('items must be an array.');
  }

  const items = data.items.map(function(item, index) {
    const row = ensureObject_(item, 'items[' + index + ']');
    return {
      name: toRequiredString_(row.name, 'items[' + index + '].name'),
      qty: toNonNegativeInteger_(row.qty, 'items[' + index + '].qty')
    };
  }).filter(function(item) {
    return item.qty > 0;
  });

  if (items.length === 0) {
    throw new Error('At least one item with qty > 0 is required.');
  }

  return {
    userName: userName,
    items: items
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
    false
  ];
}

function formatOrderItemsForSheet_(items) {
  const formattedItems = items.map(function(item) {
    return item.name + ' x ' + item.qty;
  });
  const lines = [];

  for (let index = 0; index < formattedItems.length; index += 5) {
    lines.push(formattedItems.slice(index, index + 5).join(' / '));
  }

  return lines.join('\n');
}

function formatOrderItemsForNotification_(items) {
  return items.map(function(item) {
    return item.name + ' x ' + item.qty;
  }).join('\n');
}

function buildShippingNotificationText_(shippingData) {
  return [
    '【送り依頼】',
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
  return [
    '【備品発注】',
    '依頼者: ' + orderData.userName,
    '注文内容:',
    formatOrderItemsForNotification_(orderData.items)
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
