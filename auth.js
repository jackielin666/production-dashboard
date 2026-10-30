/*
 * 生產鍋數分析儀表板 — 簡易登入（方案 A）
 *
 * 用途：防止外人隨手點開。這是「門簾」不是「鎖」：
 *   靜態網站沒有後端，data/latest.xlsx、reports/*.json 知道網址仍可直接下載。
 *
 * 規則：
 *   - 首頁、月報、月報索引都要登入，登入一次全站通用（同一個分頁內）。
 *   - 關閉分頁或瀏覽器 → 需重新登入（sessionStorage）。
 *   - 閒置 15 分鐘自動登出；第 14 分鐘跳出提醒。
 *   - 滑鼠、鍵盤、觸控、捲動都算「有在使用」。
 *
 * 帳號維護（新增／刪除／重設密碼）：改下方 ACCOUNTS，
 *   雜湊值用 `python3 tools/auth_hash.py 帳號 密碼` 產生。密碼明碼不要寫進任何檔案。
 */
(function () {
  'use strict';

  var ACCOUNTS = {
    // 帳號: { name: 顯示名稱, admin: 是否為系統管理員（僅代表負責帳號維護，畫面無額外功能）, hash: SHA-256 }
    'adm0': { name: 'Jackie', admin: true,  hash: '242c8e63530b463d5aa6b6b1bba2ed8f08b72f97fb733d166a83879fc3b761f6' },
    'adm1': { name: 'Tommy',  admin: false, hash: 'cdee661c30f93384d4357c1c8e3aafd9e05d3933fe0102c175b9be49a3f3dcaa' },
    '140':  { name: '廠長',   admin: false, hash: '6ee2d9f5ea372b63112bb133522708e5354b2b969667b0bbd461f178b4bfb9c2' }
  };
  var SALT = 'pd-v1';
  var IDLE_MS = 15 * 60 * 1000;   // 閒置自動登出
  var WARN_MS = 14 * 60 * 1000;   // 跳出「即將登出」提醒
  var KEY = 'pd-auth';

  var root = document.documentElement;

  // 1) 立刻遮住頁面內容（在 <head> 同步執行，避免內容先閃現）
  var css = document.createElement('style');
  css.textContent = [
    'html.pd-locked body{visibility:hidden!important;overflow:hidden!important;}',
    'html.pd-locked #pd-auth{visibility:visible!important;}',
    '#pd-auth{position:fixed;inset:0;z-index:100000;background:#eef1f8;display:flex;align-items:center;justify-content:center;padding:16px;',
    "  font-family:'Noto Sans TC',-apple-system,'Segoe UI','Microsoft JhengHei',sans-serif;color:#0f172a;}",
    '#pd-auth form{text-align:left;width:100%;max-width:360px;background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:28px 24px;box-sizing:border-box;}',
    '#pd-auth h1{font-size:20px;margin:0 0 4px;font-weight:700;}',
    '#pd-auth .pd-sub{font-size:13.5px;color:#64748b;margin:0 0 20px;}',
    '#pd-auth label{display:block;font-size:13.5px;font-weight:600;margin:0 0 6px;color:#334155;}',
    '#pd-auth input{width:100%;box-sizing:border-box;font-size:16px;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;margin:0 0 14px;background:#fff;color:#0f172a;}',
    '#pd-auth input:focus{outline:2px solid #1d4ed8;outline-offset:0;border-color:#1d4ed8;}',
    '#pd-auth button{width:100%;font-size:16px;font-weight:600;padding:11px;border:0;border-radius:8px;background:#1d4ed8;color:#fff;cursor:pointer;}',
    '#pd-auth button:disabled{opacity:.6;cursor:default;}',
    '#pd-auth .pd-err{min-height:20px;font-size:13.5px;color:#b45309;margin:-4px 0 10px;}',
    '#pd-auth .pd-note{font-size:12px;color:#94a3b8;margin:14px 0 0;line-height:1.6;}',
    '#pd-badge{position:fixed;right:12px;bottom:12px;z-index:99990;display:flex;align-items:center;gap:8px;font-size:12.5px;',
    "  font-family:'Noto Sans TC',-apple-system,'Segoe UI','Microsoft JhengHei',sans-serif;",
    '  background:#fff;color:#334155;border:1px solid #e2e8f0;border-radius:999px;padding:5px 6px 5px 12px;}',
    '#pd-badge button{font:inherit;font-weight:600;border:1px solid #cbd5e1;background:#f8fafc;color:#334155;border-radius:999px;padding:3px 10px;cursor:pointer;}',
    '#pd-warn{position:fixed;left:50%;top:16px;transform:translateX(-50%);z-index:99995;max-width:calc(100% - 32px);box-sizing:border-box;',
    "  font-family:'Noto Sans TC',-apple-system,'Segoe UI','Microsoft JhengHei',sans-serif;",
    '  background:#0f172a;color:#fff;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:12px;font-size:14px;}',
    '#pd-warn button{font:inherit;font-weight:600;border:0;background:#fff;color:#0f172a;border-radius:8px;padding:6px 12px;cursor:pointer;white-space:nowrap;}',
    'body.pdf #pd-badge,body.pdf #pd-warn{display:none!important;}',
    '@media print{#pd-badge,#pd-warn,#pd-auth{display:none!important;}}'
  ].join('\n');
  (document.head || root).appendChild(css);

  function now() { return Date.now(); }
  function load() {
    try {
      var s = JSON.parse(sessionStorage.getItem(KEY) || 'null');
      if (s && ACCOUNTS[s.u] && typeof s.t === 'number' && now() - s.t < IDLE_MS) return s;
    } catch (e) { /* 讀不到就視為未登入 */ }
    return null;
  }
  function save(s) { try { sessionStorage.setItem(KEY, JSON.stringify(s)); } catch (e) { /* ignore */ } }
  function clear() { try { sessionStorage.removeItem(KEY); } catch (e) { /* ignore */ } }

  var session = load();
  if (!session) { clear(); root.classList.add('pd-locked'); }

  function sha256Hex(text) {
    return crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) {
        return ('0' + b.toString(16)).slice(-2);
      }).join('');
    });
  }

  function onReady(fn) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', fn);
    else fn();
  }

  // 2) 登入畫面
  function showLogin() {
    var box = document.createElement('div');
    box.id = 'pd-auth';
    box.innerHTML =
      '<form autocomplete="on" novalidate>' +
      '<h1>生產鍋數分析儀表板</h1>' +
      '<p class="pd-sub">五惠食品內部使用，請先登入</p>' +
      '<label for="pd-user">帳號</label>' +
      '<input id="pd-user" name="username" autocomplete="username" autocapitalize="off" spellcheck="false" required>' +
      '<label for="pd-pass">密碼</label>' +
      '<input id="pd-pass" name="password" type="password" autocomplete="current-password" required>' +
      '<div class="pd-err" role="alert"></div>' +
      '<button type="submit">登入</button>' +
      '<p class="pd-note">閒置 15 分鐘或關閉分頁後需重新登入。<br>忘記密碼請洽管理部 Jackie。</p>' +
      '</form>';
    document.body.appendChild(box);
    var form = box.querySelector('form');
    var user = box.querySelector('#pd-user');
    var pass = box.querySelector('#pd-pass');
    var err = box.querySelector('.pd-err');
    var btn = box.querySelector('button');
    user.focus();

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var u = user.value.trim();
      var p = pass.value;
      if (!u || !p) { err.textContent = '請輸入帳號與密碼'; return; }
      btn.disabled = true;
      var acct = Object.prototype.hasOwnProperty.call(ACCOUNTS, u) ? ACCOUNTS[u] : null;
      sha256Hex(SALT + '|' + u + '|' + p).then(function (h) {
        btn.disabled = false;
        if (!acct || h !== acct.hash) {
          err.textContent = '帳號或密碼錯誤';
          pass.value = '';
          pass.focus();
          return;
        }
        session = { u: u, t: now() };
        save(session);
        box.parentNode.removeChild(box);
        root.classList.remove('pd-locked');
        startSession();
        // 讓圖表依實際尺寸重新排版
        try { window.dispatchEvent(new Event('resize')); } catch (e2) { /* ignore */ }
      }, function () {
        btn.disabled = false;
        err.textContent = '此瀏覽器無法登入，請改用 Chrome、Edge 或 Safari 最新版';
      });
    });
  }

  function logout() {
    session = null;   // 先清掉，避免換頁前的 pagehide 又把登入狀態存回去
    clear();
    location.reload();
  }

  // 3) 登入後：右下角身分標籤、閒置計時
  var lastWrite = 0;
  function touch() {
    if (!session) return;
    var t = now();
    session.t = t;
    hideWarn();
    if (t - lastWrite > 5000) { lastWrite = t; save(session); }
  }

  var warnEl = null;
  function showWarn() {
    if (warnEl) return;
    warnEl = document.createElement('div');
    warnEl.id = 'pd-warn';
    warnEl.setAttribute('role', 'alert');
    warnEl.innerHTML = '<span>閒置過久，1 分鐘後將自動登出</span><button type="button">繼續使用</button>';
    warnEl.querySelector('button').addEventListener('click', touch);
    document.body.appendChild(warnEl);
  }
  function hideWarn() {
    if (warnEl) { warnEl.parentNode.removeChild(warnEl); warnEl = null; }
  }

  function check() {
    if (!session) return;
    var idle = now() - session.t;
    if (idle >= IDLE_MS) { logout(); return; }
    if (idle >= WARN_MS) showWarn();
  }

  function startSession() {
    var acct = ACCOUNTS[session.u];
    var badge = document.createElement('div');
    badge.id = 'pd-badge';
    badge.innerHTML = '<span></span><button type="button">登出</button>';
    badge.querySelector('span').textContent = acct.name + ' 已登入';
    badge.querySelector('button').addEventListener('click', logout);
    document.body.appendChild(badge);

    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'wheel', 'scroll'].forEach(function (ev) {
      window.addEventListener(ev, touch, { passive: true, capture: true });
    });
    save(session);
    setInterval(check, 5000);
    // 手機鎖屏／切換 App 回來時，計時器可能被暫停，回到頁面立即檢查
    document.addEventListener('visibilitychange', function () { if (!document.hidden) check(); });
    // 換頁前存下最後操作時間；按「上一頁」回來時，若已在別頁登出就重新上鎖
    window.addEventListener('pagehide', function () { if (session) save(session); });
    window.addEventListener('pageshow', function (e) {
      if (e.persisted && !load()) { session = null; location.reload(); return; }
      check();
    });
  }

  onReady(function () {
    if (session) startSession();
    else showLogin();
  });
})();
