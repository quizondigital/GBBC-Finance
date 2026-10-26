(function () {
  var seconds = Number(document.body.getAttribute('data-idle-timeout') || '600');
  if (!seconds || seconds < 30) return;

  var timeoutMs = seconds * 1000;
  var pingUrl = document.body.getAttribute('data-session-ping') || '/session/ping';
  var loginUrl = document.body.getAttribute('data-login-url') || '/login';
  var timer = null;
  var lastPing = 0;
  var pingGapMs = Math.min(60000, Math.max(15000, Math.floor(timeoutMs / 4)));

  function signOutIdle() {
    window.location.href = loginUrl;
  }

  function pingSession() {
    var now = Date.now();
    if (now - lastPing < pingGapMs) return;
    lastPing = now;
    fetch(pingUrl, {
      method: 'GET',
      credentials: 'same-origin',
      headers: { Accept: 'application/json' },
      cache: 'no-store',
    }).then(function (response) {
      if (response.status === 401) signOutIdle();
    }).catch(function () { /* ignore transient network errors */ });
  }

  function resetIdleTimer() {
    window.clearTimeout(timer);
    timer = window.setTimeout(signOutIdle, timeoutMs);
    pingSession();
  }

  ['mousemove', 'mousedown', 'keydown', 'scroll', 'touchstart', 'click', 'visibilitychange'].forEach(function (eventName) {
    document.addEventListener(eventName, function () {
      if (eventName === 'visibilitychange' && document.hidden) return;
      resetIdleTimer();
    }, { passive: true });
  });

  resetIdleTimer();
})();
