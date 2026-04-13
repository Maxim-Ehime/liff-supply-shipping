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
