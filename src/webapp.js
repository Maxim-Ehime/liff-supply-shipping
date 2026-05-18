function doPost(e) {
  try {
    const request = parseRequestBody_(e);
    const action = toRequiredString_(request.action, 'action');
    const config = getAppConfig_();

    if (action === 'liff_shipping') {
      const shippingData = normalizeShippingPayload_(request.data);
      appendShippingToSheet_(shippingData, config);
      pushLineTextMessage_(config, buildShippingNotificationText_(shippingData));
      if (shippingData.lineUserId) {
        try {
          pushLineTextMessageTo_(config, shippingData.lineUserId, buildShippingReceiptText_(shippingData));
        } catch (notifyError) {
          console.error('Failed to notify shipping requester:', toErrorMessage_(notifyError));
        }
      }
      return createJsonResponse_({ ok: true, action: action });
    }

    if (action === 'liff_order') {
      const orderData = normalizeOrderPayload_(request.data);
      appendSupplyOrderToSheet_(orderData, config);
      pushLineTextMessage_(config, buildOrderNotificationText_(orderData));
      if (orderData.lineUserId) {
        try {
          pushLineTextMessageTo_(config, orderData.lineUserId, buildOrderReceiptText_(orderData));
        } catch (notifyError) {
          console.error('Failed to notify supply requester:', toErrorMessage_(notifyError));
        }
      }
      return createJsonResponse_({ ok: true, action: action });
    }

    throw new Error('Unsupported action: ' + action);
  } catch (error) {
    return createJsonResponse_({
      ok: false,
      error: toErrorMessage_(error)
    });
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
