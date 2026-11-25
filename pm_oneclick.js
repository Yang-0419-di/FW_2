/* pm_oneclick.js — Git 發佈頁的「⚡ 一鍵執行」區塊
   由 pm_app.py 送出管理介面時自動掛上（不修改 index.html）。
   位置：放在 Git 發佈頁「檢查發佈內容」按鈕所在區塊的上方。
   流程：下載當月 EIP 報表 → 資料更新（當月）→ Git 發佈（不再確認）。任何一步失敗就停止。 */
(function () {
  'use strict';
  var ANCHOR_TEXT = '檢查發佈內容';
  var BOX_ID = 'pmOneClick';
  var pollTimer = null;
  var lastState = null;

  function el(tag, attrs, text) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === 'style') e.style.cssText = attrs[k]; else e.setAttribute(k, attrs[k]);
    }
    if (text != null) e.textContent = text;
    return e;
  }

  function api(method, url, body) {
    return fetch(url, {
      method: method, cache: 'no-store',
      headers: body ? { 'Content-Type': 'application/json' } : {},
      body: body ? JSON.stringify(body) : undefined
    }).then(function (r) {
      return r.json().catch(function () { return {}; }).then(function (j) {
        if (!r.ok) throw new Error(j.error || ('HTTP ' + r.status));
        return j;
      });
    });
  }

  function injectStyle() {
    if (document.getElementById('pmOneClickStyle')) return;
    var css = [
      '#pmOneClick{border:2px solid #f59e0b;background:#fffbeb;border-radius:10px;padding:14px 16px;margin:0 0 16px 0;font-size:14px;color:#1f2937;}',
      '#pmOneClick h3{margin:0 0 6px 0;font-size:17px;}',
      '#pmOneClick .oc-desc{margin:0 0 10px 0;line-height:1.6;color:#4b5563;}',
      '#pmOneClick .oc-row{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:6px 0;}',
      '#pmOneClick input{padding:6px 8px;border:1px solid #cbd5e1;border-radius:6px;font-size:14px;}',
      '#pmOneClick .oc-go{background:#d97706;color:#fff;border:none;border-radius:8px;padding:10px 22px;font-size:16px;font-weight:bold;cursor:pointer;}',
      '#pmOneClick .oc-go:disabled{background:#9ca3af;cursor:default;}',
      '#pmOneClick .oc-status{font-weight:bold;margin:8px 0 4px 0;}',
      '#pmOneClick .oc-status.ok{color:#15803d;} #pmOneClick .oc-status.err{color:#b91c1c;} #pmOneClick .oc-status.run{color:#b45309;}',
      '#pmOneClick pre{background:#111827;color:#e5e7eb;border-radius:8px;padding:10px;max-height:320px;overflow:auto;font-size:12px;white-space:pre-wrap;word-break:break-all;margin:6px 0 0 0;}',
      '#pmOneClick .oc-msg{color:#b91c1c;margin-top:6px;white-space:pre-wrap;}',
      '#pmOneClick .oc-link{background:none;border:none;color:#2563eb;text-decoration:underline;cursor:pointer;padding:0;font-size:14px;}'
    ].join('\n');
    var s = el('style', { id: 'pmOneClickStyle' }); s.textContent = css; document.head.appendChild(s);
  }

  function build() {
    var box = el('div', { id: BOX_ID });
    box.appendChild(el('h3', null, '⚡ 一鍵執行'));
    box.appendChild(el('p', { 'class': 'oc-desc' },
      '依序自動執行：① 下載當月 EIP 報表 → ② 資料更新（月份＝當月）→ ③ Git 發佈（不再確認，直接推送）。' +
      '適合確定資料沒問題時使用；任何一步失敗都會停止，不會發佈。'));

    var cred = el('div', { 'class': 'oc-row oc-cred', style: 'display:none' });
    cred.appendChild(el('span', null, 'EIP 帳號'));
    cred.appendChild(el('input', { type: 'text', id: 'ocUser', autocomplete: 'username', style: 'width:140px' }));
    cred.appendChild(el('span', null, '密碼'));
    cred.appendChild(el('input', { type: 'password', id: 'ocPass', autocomplete: 'current-password', style: 'width:140px' }));
    cred.appendChild(el('span', { style: 'color:#6b7280;font-size:12px' }, '（只用於這次下載，不會存檔）'));
    box.appendChild(cred);

    var row = el('div', { 'class': 'oc-row' });
    var go = el('button', { type: 'button', 'class': 'oc-go' }, '⚡ 一鍵執行');
    row.appendChild(go);
    var month = el('span', { 'class': 'oc-month', style: 'color:#6b7280' }, '');
    row.appendChild(month);
    box.appendChild(row);

    box.appendChild(el('div', { 'class': 'oc-msg' }, ''));
    box.appendChild(el('div', { 'class': 'oc-status' }, ''));
    var logToggle = el('button', { type: 'button', 'class': 'oc-link', style: 'display:none' }, '顯示執行紀錄');
    box.appendChild(logToggle);
    var pre = el('pre', { style: 'display:none' });
    box.appendChild(pre);
    var reload = el('button', { type: 'button', 'class': 'oc-link', style: 'display:none;margin-left:12px' }, '重新整理畫面（載入更新後的資料）');
    box.appendChild(reload);

    logToggle.addEventListener('click', function () {
      var show = pre.style.display === 'none';
      pre.style.display = show ? '' : 'none';
      logToggle.textContent = show ? '隱藏執行紀錄' : '顯示執行紀錄';
      if (show) pre.scrollTop = pre.scrollHeight;
    });
    reload.addEventListener('click', function () { location.reload(); });
    go.addEventListener('click', function () { start(box); });
    return box;
  }

  function q(box, sel) { return box.querySelector(sel); }

  function render(box, st) {
    if (!box || !st) return;
    lastState = st;
    var go = q(box, '.oc-go'), status = q(box, '.oc-status'), pre = q(box, 'pre'), toggle = q(box, '.oc-link'), msg = q(box, '.oc-msg');
    var reload = box.querySelectorAll('.oc-link')[1];
    q(box, '.oc-cred').style.display = (!st.env_user || !st.env_pass) && !st.running ? 'flex' : 'none';
    q(box, '.oc-month').textContent = '資料月份：' + st.month;
    go.disabled = !!st.running || !st.git_enabled || !st.ready;
    go.textContent = st.running ? '執行中…' : '⚡ 一鍵執行';
    if (!st.git_enabled) msg.textContent = 'Git 功能目前是關閉的（pm_config.json 的 git_enabled 為 false），無法使用一鍵執行。';

    var mins = Math.floor((st.elapsed || 0) / 60), secs = (st.elapsed || 0) % 60;
    var t = (mins ? mins + ' 分 ' : '') + secs + ' 秒';
    if (st.running) { status.className = 'oc-status run'; status.textContent = '⏳ ' + st.stage + '（已執行 ' + t + '）'; }
    else if (st.done && st.ok) { status.className = 'oc-status ok'; status.textContent = '✅ ' + st.stage + '（共 ' + t + '）'; }
    else if (st.done) { status.className = 'oc-status err'; status.textContent = '❌ 已停止：' + (st.error || '').split('\n')[0]; }
    else { status.className = 'oc-status'; status.textContent = ''; }

    var hasLog = st.log && st.log.length > 1;
    toggle.style.display = hasLog ? '' : 'none';
    if (hasLog) {
      var atBottom = pre.scrollTop + pre.clientHeight >= pre.scrollHeight - 20;
      pre.textContent = st.log.join('\n');
      if (atBottom) pre.scrollTop = pre.scrollHeight;
      if (st.running || (st.done && !st.ok)) { pre.style.display = ''; toggle.textContent = '隱藏執行紀錄'; }
    }
    reload.style.display = st.done && !st.running ? '' : 'none';
  }

  function poll(box) {
    clearTimeout(pollTimer);
    api('GET', '/api/oneclick/state').then(function (st) {
      render(document.getElementById(BOX_ID) || box, st);
      if (st.running) pollTimer = setTimeout(function () { poll(box); }, 1500);
    }).catch(function () { pollTimer = setTimeout(function () { poll(box); }, 4000); });
  }

  function start(box) {
    var msg = q(box, '.oc-msg'); msg.textContent = '';
    var body = {};
    var u = q(box, '#ocUser'), p = q(box, '#ocPass');
    if (q(box, '.oc-cred').style.display !== 'none') {
      body.user = (u.value || '').trim(); body.password = p.value || '';
      if (!body.user || !body.password) { msg.textContent = '請輸入 EIP 帳號與密碼'; return; }
    }
    q(box, '.oc-go').disabled = true;
    api('POST', '/api/oneclick/start', body).then(function () {
      if (p) p.value = '';
      poll(box);
    }).catch(function (e) {
      msg.textContent = e.message; q(box, '.oc-go').disabled = false;
    });
  }

  // 找到 Git 發佈頁的「檢查發佈內容」按鈕，把區塊放在它所在區塊的上方
  function findAnchor() {
    var btns = document.querySelectorAll('button, input[type=button], input[type=submit], a');
    for (var i = 0; i < btns.length; i++) {
      var b = btns[i];
      var t = (b.textContent || b.value || '').trim();
      if (t.indexOf(ANCHOR_TEXT) >= 0 && !b.closest('#' + BOX_ID)) return b;
    }
    return null;
  }

  function mount() {
    if (document.getElementById(BOX_ID)) return;
    var a = findAnchor();
    if (!a) return;
    injectStyle();
    var box = build();
    var host = a.parentElement;
    host.parentNode.insertBefore(box, host);
    if (lastState) render(box, lastState);
    poll(box);
  }

  function init() {
    mount();
    // 管理介面切換分頁時會重新產生畫面，持續確認區塊存在
    var pending = false;
    new MutationObserver(function () {
      if (pending) return; pending = true;
      setTimeout(function () { pending = false; mount(); }, 150);
    }).observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
