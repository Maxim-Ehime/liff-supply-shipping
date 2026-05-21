function doPost(e) {
  try {
    const request = parseRequestBody_(e);
    const action = toRequiredString_(request.action, 'action');
    const config = getAppConfig_();

    if (action === 'liff_shipping') {
      const shippingData = normalizeShippingPayload_(request.data);
      appendShippingToSheet_(shippingData, config);
      pushLineTextMessage_(config, buildShippingNotificationText_(shippingData));
      return createJsonResponse_({ ok: true, action: action });
    }

    if (action === 'liff_order') {
      const orderData = normalizeOrderPayload_(request.data);
      appendSupplyOrderToSheet_(orderData, config);
      pushLineTextMessage_(config, buildOrderNotificationText_(orderData));
      return createJsonResponse_({ ok: true, action: action });
    }

    if (action === 'liff_product_request') {
      const requestData = normalizeProductRequestPayload_(request.data);
      const requestId = createRequestId_('PRD');
      const saved = saveProductRequestImages_(requestId, requestData, config);
      const payload = Object.assign({}, requestData, {
        requestId: requestId,
        imageUrls: saved.imageUrls,
        reviewStatus: saved.imageCount === 0 && containsLineMention_(requestData.requestText)
          ? '要LINE確認'
          : '未確認'
      });
      appendProductRequestToSheet_(payload, config);
      pushLineTextMessage_(config, buildProductRequestNotificationText_(payload));
      return createJsonResponse_({
        ok: true,
        action: action,
        requestId: requestId,
        imageCount: saved.imageCount
      });
    }

    if (action === 'liff_history') {
      const historyRequest = normalizeHistoryRequest_(request.data);
      const items = getHistoryItems_(historyRequest, config);
      return createJsonResponse_({
        ok: true,
        action: action,
        items: items
      });
    }

    throw new Error('Unsupported action: ' + action);
  } catch (error) {
    return createJsonResponse_({
      ok: false,
      error: toErrorMessage_(error)
    });
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ダッシュボード')
    .addItem('日付を選択', 'showDashboardCalendar')
    .addItem('初期作成', 'setupDashboard')
    .addItem('更新', 'refreshDashboard')
    .addItem('依頼者順を補完', 'supplementRequesterOrder')
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
