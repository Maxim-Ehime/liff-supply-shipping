(function (global) {
  const environments = {
    production: {
      gasUrl: 'https://script.google.com/macros/s/AKfycbz6M7Gi704Vc3qlMQ7GD4Hh8mq0rf4I8QH-kqT7oOs1T8jkYJB6JQFQmMS3IhGD31VE/exec',
      liffId: '2009669107-iTCKuJvo'
    },
    test: {
      gasUrl: 'https://script.google.com/macros/s/AKfycbz1GRizrjZXu0mQ-N58HgTwWJ55BuG4VteRkP4aPmhuWOaeRJNTejgVZuHRmwpsr4E/exec',
      liffId: '2009819015-Pj5rk2Ru'
    }
  };

  const activeEnvironment = 'test';
  const activeConfig = environments[activeEnvironment];

  global.LiffConfig = {
    ENVIRONMENT: activeEnvironment,
    GAS_URL: activeConfig.gasUrl,
    LIFF_ID: activeConfig.liffId
  };
})(window);
