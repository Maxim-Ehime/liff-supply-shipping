const ALLOWED_ACTIONS = new Set(['liff_order', 'liff_shipping']);

exports.handler = async function handler(event) {
  if (event.httpMethod !== 'POST') {
    return createJsonResponse(405, { ok: false, error: 'Method Not Allowed' });
  }

  const gasWebAppId = process.env.GAS_WEBAPP_ID;
  if (!gasWebAppId) {
    return createJsonResponse(500, { ok: false, error: 'Missing environment variable: GAS_WEBAPP_ID' });
  }

  let payload;
  try {
    payload = JSON.parse(event.body || '');
  } catch (error) {
    return createJsonResponse(400, { ok: false, error: 'Request body must be valid JSON.' });
  }

  if (!payload || typeof payload !== 'object') {
    return createJsonResponse(400, { ok: false, error: 'Request payload must be an object.' });
  }

  if (!ALLOWED_ACTIONS.has(payload.action)) {
    return createJsonResponse(400, { ok: false, error: 'Unsupported action.' });
  }

  const gasEndpoint = 'https://script.google.com/macros/s/' + gasWebAppId + '/exec';
  const response = await fetch(gasEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8'
    },
    body: JSON.stringify(payload)
  });

  const responseBody = await response.text();
  const responseType = response.headers.get('content-type') || 'application/json; charset=utf-8';

  return {
    statusCode: response.status,
    headers: {
      'Content-Type': responseType
    },
    body: responseBody
  };
};

function createJsonResponse(statusCode, payload) {
  return {
    statusCode: statusCode,
    headers: {
      'Content-Type': 'application/json; charset=utf-8'
    },
    body: JSON.stringify(payload)
  };
}
