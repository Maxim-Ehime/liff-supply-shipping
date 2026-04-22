exports.handler = async function handler() {
  const liffId = process.env.LIFF_ID || '';
  const appConfig = {
    LIFF_ID: liffId,
    GAS_PROXY_URL: '/api/gas'
  };

  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/javascript; charset=utf-8',
      'Cache-Control': 'no-store'
    },
    body: 'window.APP_CONFIG = Object.freeze(' + JSON.stringify(appConfig) + ');'
  };
};
