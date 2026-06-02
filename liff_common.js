(function (global) {
  const PENDING_FORM_KEY = 'pendingLiffForm';
  const PENDING_FORM_TIME_KEY = 'pendingLiffFormTime';
  const PENDING_FORM_TTL_MS = 120000;
  const VALID_PENDING_FORMS = ['order', 'history', 'product'];

  function normalizeForm(form) {
    return String(form || '').toLowerCase();
  }

  function clearPendingLiffForm() {
    try {
      sessionStorage.removeItem(PENDING_FORM_KEY);
      sessionStorage.removeItem(PENDING_FORM_TIME_KEY);
    } catch (error) {
      // Ignore storage errors in restricted browser contexts.
    }
  }

  function rememberPendingLiffForm(form) {
    const normalized = normalizeForm(form);
    if (VALID_PENDING_FORMS.indexOf(normalized) === -1) return;

    try {
      sessionStorage.setItem(PENDING_FORM_KEY, normalized);
      sessionStorage.setItem(PENDING_FORM_TIME_KEY, String(Date.now()));
    } catch (error) {
      // Ignore storage errors in restricted browser contexts.
    }
  }

  function getPendingLiffForm() {
    try {
      const pendingForm = normalizeForm(sessionStorage.getItem(PENDING_FORM_KEY));
      const pendingTime = Number(sessionStorage.getItem(PENDING_FORM_TIME_KEY));
      if (
        VALID_PENDING_FORMS.indexOf(pendingForm) !== -1 &&
        Number.isFinite(pendingTime) &&
        Date.now() - pendingTime < PENDING_FORM_TTL_MS
      ) {
        return pendingForm;
      }
    } catch (error) {
      // Ignore storage errors in restricted browser contexts.
    }
    return '';
  }

  function getLoginRedirectUri(form) {
    const url = new URL('./index.html', window.location.href);
    url.searchParams.set('form', normalizeForm(form));
    return url.href;
  }

  global.LiffRouting = {
    clearPendingLiffForm,
    getLoginRedirectUri,
    getPendingLiffForm,
    rememberPendingLiffForm
  };
})(window);
