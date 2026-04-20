function doPost(e) {
  try {
    const request = parseRequestBody_(e);
    const config = getAppConfig_();

    // If this request looks like a LINE webhook (contains events array), process events
    if (request && Array.isArray(request.events)) {
      handleLineWebhook_(request, config);
      return createJsonResponse_({ ok: true });
    }

    const action = toRequiredString_(request.action, 'action');

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

    throw new Error('Unsupported action: ' + action);
  } catch (error) {
    return createJsonResponse_({
      ok: false,
      error: toErrorMessage_(error)
    });
  }
}

// Process incoming webhook events from LINE Messaging API
function handleLineWebhook_(request, config) {
  if (!request.events || !Array.isArray(request.events)) return;

  request.events.forEach(function (event) {
    try {
      if (event.type !== 'message' || !event.message || event.message.type !== 'text') return;

      const replyToken = event.replyToken;
      if (!replyToken) return; // cannot reply without token

      const idSent = String(event.message.text || '').trim();
      const userId = (event.source && event.source.userId) ? event.source.userId : 'unknown';

      const message = [
        'あなたのuserId: ' + userId
      ].join('\n');

      pushLineReplyMessage_(config, replyToken, message);
    } catch (err) {
      console.error('Error handling LINE event:', err && err.message ? err.message : err);
    }
  });
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
