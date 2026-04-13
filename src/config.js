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
