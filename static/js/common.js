/*!
 * common.js — 全站共用前端邏輯（由 fix_project.py 產生；請改這一份，不要再複製到各模板）
 *  1. 全域載入遮罩
 *  2. fetch 回 401（登入逾時）時自動導向登入頁
 *  3. 閒置自動登出 + 使用中定期 ping（前端計時與後端 session 同步）
 *  4. 裝置指紋（每次登入只送一次）
 */
(function () {
  'use strict';

  var script = document.currentScript;
  var cfg = (script && script.dataset) || {};
  var isAuth = cfg.auth === '1';
  var loginUrl = cfg.loginUrl || '/billing/login';
  var logoutUrl = cfg.logoutUrl || '/billing/logout';
  var pingUrl = cfg.pingUrl || '/api/ping';
  var logUrl = cfg.logUrl || '/api/save_log';
  var idleMs = (parseInt(cfg.idleMinutes, 10) || 10) * 60 * 1000;
  var pingEveryMs = Math.min(3 * 60 * 1000, Math.floor(idleMs / 2));
  var warnBeforeMs = Math.min(60 * 1000, Math.floor(idleMs / 5));
  var nativeFetch = window.fetch ? window.fetch.bind(window) : null;

  /* ---------- 1. 載入遮罩 ---------- */
  function ensureOverlay() {
    var el = document.getElementById('global-loading');
    if (el) return el;
    el = document.createElement('div');
    el.id = 'global-loading';
    el.style.cssText = 'display:none;position:fixed;top:0;left:0;width:100%;height:100%;' +
      'background:rgba(0,0,0,0.4);z-index:9999;justify-content:center;align-items:center;';
    var box = document.createElement('div');
    box.style.cssText = 'background:white;padding:30px 50px;border-radius:12px;font-size:22px;' +
      'font-weight:bold;box-shadow:0 4px 20px rgba(0,0,0,0.3);';
    box.textContent = '\u23F3 加載中...請稍後';
    el.appendChild(box);
    document.body.appendChild(el);
    return el;
  }
  function showGlobalLoading() { ensureOverlay().style.display = 'flex'; }
  function hideGlobalLoading() {
    var el = document.getElementById('global-loading');
    if (el) el.style.display = 'none';
  }
  window.showGlobalLoading = showGlobalLoading;
  window.hideGlobalLoading = hideGlobalLoading;

  document.addEventListener('click', function (e) {
    var link = e.target && e.target.closest ? e.target.closest('a') : null;
    if (!link) return;
    var rawHref = link.getAttribute('href');
    if (!rawHref || rawHref.indexOf('javascript') === 0 || rawHref.charAt(0) === '#') return;
    if (link.target === '_blank' || link.hasAttribute('download')) return;
    if (e.ctrlKey || e.metaKey || e.shiftKey || e.button === 1) return;
    showGlobalLoading();
  });
  document.addEventListener('submit', function () { showGlobalLoading(); });
  window.addEventListener('pageshow', hideGlobalLoading);

  /* ---------- 2. fetch 遇到 401 → 導向登入頁 ---------- */
  var redirecting = false;
  function goLogin() {
    if (redirecting) return;
    redirecting = true;
    var next = location.pathname + location.search;
    location.replace(loginUrl + (loginUrl.indexOf('?') < 0 ? '?' : '&') + 'next=' + encodeURIComponent(next));
  }
  window.goLogin = goLogin;

  if (nativeFetch) {
    window.fetch = function (input, init) {
      return nativeFetch(input, init).then(function (res) {
        if (res.status === 401) {
          var url = typeof input === 'string' ? input : ((input && input.url) || '');
          if (url.charAt(0) === '/' || url.indexOf(location.origin) === 0) goLogin();
        }
        return res;
      });
    };
  }

  /* ---------- 3. 閒置登出 + 定期 ping ---------- */
  var sessKey = 'fp_sent:' + (cfg.sess || '');

  if (isAuth) {
    var lastActivity = Date.now();
    var lastPing = Date.now();
    var banner = null;

    var showBanner = function (sec) {
      if (!banner) {
        banner = document.createElement('div');
        banner.style.cssText = 'position:fixed;left:50%;bottom:20px;transform:translateX(-50%);z-index:10000;' +
          'background:#b45309;color:#fff;padding:12px 22px;border-radius:8px;font-size:16px;' +
          'box-shadow:0 4px 16px rgba(0,0,0,0.3);';
        document.body.appendChild(banner);
      }
      banner.textContent = '\u26A0\uFE0F 長時間未操作，將於 ' + sec + ' 秒後自動登出。移動滑鼠或按任意鍵即可繼續使用。';
    };
    var hideBanner = function () {
      if (banner && banner.parentNode) banner.parentNode.removeChild(banner);
      banner = null;
    };
    var onActivity = function () { lastActivity = Date.now(); if (banner) hideBanner(); };
    ['mousemove', 'keydown', 'click', 'scroll', 'touchstart', 'input'].forEach(function (ev) {
      window.addEventListener(ev, onActivity, { passive: true, capture: true });
    });

    var doLogout = function () {
      try { sessionStorage.removeItem(sessKey); } catch (e) {}
      var done = function () { location.replace(loginUrl); };
      if (!nativeFetch) { done(); return; }
      nativeFetch(logoutUrl, { method: 'POST', credentials: 'same-origin', keepalive: true }).then(done, done);
    };
    var ping = function () {
      lastPing = Date.now();
      return fetch(pingUrl, {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'X-Requested-With': 'fetch' }
      }).catch(function () {});
    };
    var tick = function () {
      var now = Date.now();
      var idle = now - lastActivity;
      if (idle >= idleMs) { doLogout(); return; }
      if (idle >= idleMs - warnBeforeMs) { showBanner(Math.ceil((idleMs - idle) / 1000)); return; }
      // 使用中：每隔 pingEveryMs 通知後端延長 session，避免輸入到一半被登出
      if (lastActivity > lastPing && now - lastPing >= pingEveryMs) ping();
    };
    setInterval(tick, 5000);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) tick(); });

    // 手動按登出時，清掉指紋旗標讓下次登入重新記錄
    document.addEventListener('submit', function (e) {
      var action = e.target && e.target.getAttribute ? (e.target.getAttribute('action') || '') : '';
      if (action && action.indexOf(logoutUrl) !== -1) {
        try { sessionStorage.removeItem(sessKey); } catch (err) {}
      }
    });
  }

  /* ---------- 4. 裝置指紋（每次登入只送一次） ---------- */
  function b64(str) {
    var bytes = new TextEncoder().encode(str);
    var bin = '';
    for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }

  function getReadableDeviceFingerprint() {
    var ua = navigator.userAgent;
    var deviceName = 'Unknown Device';
    var browserName = 'Unknown Browser';

    if (/iphone/i.test(ua)) {
      deviceName = 'Apple iPhone';
    } else if (/ipad/i.test(ua)) {
      deviceName = 'Apple iPad';
    } else if (/macintosh|mac os x/i.test(ua)) {
      deviceName = 'Mac';
    } else if (/android/i.test(ua)) {
      deviceName = 'Android Device';
      var match = ua.match(/\b([A-Z0-9]+-[A-Z0-9]+)\b/);
      if (match) deviceName = 'Android (' + match[1] + ')';
    } else if (/win/i.test(ua)) {
      deviceName = 'Windows PC';
    } else if (/linux/i.test(ua)) {
      deviceName = 'Linux PC';
    }

    if (/edg/i.test(ua)) {
      browserName = 'Edge';
    } else if (/chrome/i.test(ua) && !/chromium/i.test(ua)) {
      browserName = 'Chrome';
    } else if (/safari/i.test(ua) && !/chrome/i.test(ua)) {
      browserName = 'Safari';
    } else if (/firefox/i.test(ua)) {
      browserName = 'Firefox';
    } else if (/line/i.test(ua)) {
      browserName = 'LINE App';
    }

    var res = screen.width + 'x' + screen.height;
    var shortHash = 'fingerprint_error';
    try {
      var canvas = document.createElement('canvas');
      var gl = canvas.getContext('webgl') || canvas.getContext('experimental-webgl');
      var renderer = gl ? gl.getParameter(gl.RENDERER) : 'unknown';
      var data = [
        ua,
        navigator.language || '',
        res,
        (Intl.DateTimeFormat().resolvedOptions().timeZone || ''),
        renderer
      ].join('|');
      shortHash = b64(data).substring(0, 32);
    } catch (e) {
      try { shortHash = b64(ua).substring(0, 32); } catch (e2) {}
    }
    return deviceName + ' / ' + browserName + ' (' + res + ')-' + shortHash;
  }

  function sendFingerprintOnce() {
    try { if (sessionStorage.getItem(sessKey) === '1') return; } catch (e) {}
    fetch(logUrl, {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fingerprint: getReadableDeviceFingerprint() })
    }).then(function (res) {
      if (res.ok) { try { sessionStorage.setItem(sessKey, '1'); } catch (e) {} }
    }).catch(function () {});
  }

  if (isAuth) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', sendFingerprintOnce);
    else sendFingerprintOnce();
  }
})();
