function doPost(e) {
  try {
    const request = parseRequestBody_(e);
    const action = toRequiredString_(request.action, 'action');
    const config = getAppConfig_();
    const handler = getDoPostActionHandler_(action);
    return createJsonResponse_(handler(request.data, config, action));
  } catch (error) {
    return createJsonResponse_({
      ok: false,
      error: toErrorMessage_(error)
    });
  }
}

function getDoPostActionHandler_(action) {
  const handlers = {
    liff_shipping: handleShippingRequest_,
    liff_order: handleSupplyOrderRequest_,
    liff_product_request: handleProductRequest_,
    liff_history: handleHistoryRequest_,
    liff_delete_request: handleDeleteRequest_
  };
  const handler = handlers[action];
  if (!handler) {
    throw new Error('Unsupported action: ' + action);
  }
  return handler;
}

function handleShippingRequest_(data, config, action) {
  const shippingData = normalizeShippingPayload_(data);
  appendShippingToSheet_(shippingData, config);
  const notification = safePushLineTextMessage_(config, buildShippingNotificationText_(shippingData));
  return { ok: true, action: action, notification: notification };
}

function handleSupplyOrderRequest_(data, config, action) {
  const orderData = normalizeOrderPayload_(data);
  appendSupplyOrderToSheet_(orderData, config);
  const notification = safePushLineTextMessage_(config, buildOrderNotificationText_(orderData));
  return { ok: true, action: action, notification: notification };
}

function handleProductRequest_(data, config, action) {
  const requestData = normalizeProductRequestPayload_(data);
  const requestId = createRequestId_('PRD');
  const saved = saveProductRequestImages_(requestId, requestData, config);
  const payload = Object.assign({}, requestData, {
    requestId: requestId,
    imageUrls: saved.imageUrls,
    imageFolderUrl: saved.folderUrl
  });
  appendProductRequestToSheet_(payload, config);
  const notification = safePushLineTextMessage_(config, buildProductRequestNotificationText_(payload));
  return {
    ok: true,
    action: action,
    requestId: requestId,
    imageCount: saved.imageCount,
    notification: notification
  };
}

function handleHistoryRequest_(data, config, action) {
  const historyRequest = normalizeHistoryRequest_(data);
  const history = getHistoryItems_(historyRequest, config);
  return {
    ok: true,
    action: action,
    items: history.items,
    hasMore: history.hasMore,
    nextOffset: history.nextOffset
  };
}

function handleDeleteRequest_(data, config, action) {
  const deleteRequest = normalizeDeleteRequestPayload_(data);
  const result = deleteUserRequest_(deleteRequest, config);
  return {
    ok: true,
    action: action,
    requestId: deleteRequest.requestId,
    deleted: result.deleted,
    type: result.type
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ダッシュボード')
    .addItem('日付を選択', 'showDashboardCalendar')
    .addItem('初期作成', 'setupDashboard')
    .addItem('更新', 'refreshDashboard')
    .addItem('依頼者順を補完', 'supplementRequesterOrder')
    .addSeparator()
    .addItem('権限を確認', 'authorizeRequiredScopes')
    .addItem('商品希望リンクを整形', 'formatProductRequestSheetLinks')
    .addItem('商品希望画像を整理', 'cleanupOldProductRequestImageFolders')
    .addItem('商品希望画像の自動整理を設定', 'installProductImageCleanupTrigger')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range || !e.source) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== DASHBOARD_SHEET_NAME) {
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 3 && col === 1) {
    showDashboardCalendar();
    return;
  }

  if ((row === 3 || row === 4) && col === 2) {
    refreshDashboard();
    return;
  }

  if (row >= DASHBOARD_DATA_START_ROW && col === DASHBOARD_DONE_COLUMN) {
    syncDashboardDoneToSource(e);
  }
}

function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Request body is empty.');
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('Request body must be valid JSON: ' + toErrorMessage_(error));
  }
}
