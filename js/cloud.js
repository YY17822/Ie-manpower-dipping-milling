/*
 * cloud.js — 登入與共用設定（Supabase）。
 *
 * 部門的人共用同一份「記住的設定」：誰最後存，大家打開看到的就是誰的。
 *   資料庫  Supabase 專案（js/config.js 的 SUPABASE_PUBLIC_URL / SUPABASE_PUBLISHABLE_KEY）
 *   位置    schema ie_manpower 的 settings 表，一筆 = (app, key) → value(jsonb)
 *           app 是這個 PWA 的專案代號（PROJECT_ID）
 *           key 是「步驟.輸入項目」，例如 step7.workHours（第 7 步總表的工作時數）
 *           一個輸入項目一筆；6 個月這種一起看的值放在同一筆（value 是陣列）
 *           這樣兩個人同時改同一步的不同項目，也不會互相蓋掉
 *   登入    Supabase Auth 的 email + 密碼（帳號由管理者在 Supabase 後台建立，不開放註冊）
 *
 * 各步驟用 IEApp.settings.read(步驟) / write(步驟, 物件) 讀寫一整步的設定（例如 'step7'），
 * 這裡負責拆成一個個輸入項目上傳、下載時再組回去：
 *   讀    一律讀這台電腦的本機備份（localStorage），所以是同步的、離線也能用
 *   寫    先存本機，再在背景上傳 Supabase（連續打字時等停下來 1 秒才傳）
 *   開啟  登入後先把 Supabase 上最新的設定抓下來蓋掉本機備份，才載入各步驟
 *   離線  照常運作，改過的設定記為「還沒上傳」，下次連上網路時再傳
 *
 * 連線用 supabase-js 套件（vendor/supabase.js，放在專案裡，離線也能載入）。
 * 雲端 Supabase、本機自架的 Supabase、公司主機上用 Docker 架的 Supabase 都用同一套寫法，
 * 換資料庫時只要改 js/config.js 的 SUPABASE_PUBLIC_URL、SUPABASE_PUBLISHABLE_KEY。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var h = IEApp.h;
  var C = global.APP_CONFIG || {};
  // 設定名稱照自架 Supabase 的 .env（SUPABASE_PUBLIC_URL、SUPABASE_PUBLISHABLE_KEY）
  var URL_BASE = String(C.SUPABASE_PUBLIC_URL || '').replace(/\/+$/, '');
  var KEY = C.SUPABASE_PUBLISHABLE_KEY || '';
  var SCHEMA = C.SUPABASE_SCHEMA || 'ie_manpower';
  var APP = IEApp.projectId;
  var LIB = global.supabase && typeof global.supabase.createClient === 'function' ? global.supabase : null;
  var ENABLED = !!(URL_BASE && KEY);

  var AUTH_KEY = IEApp.storeKey('_auth');         // supabase-js 存登入資訊的位置（只存這台電腦）
  var OLD_SESSION_KEY = IEApp.storeKey('_session'); // 版本 31 以前自己存的登入資訊，第一次打開時搬給 supabase-js
  var SYNC_KEY = IEApp.storeKey('_sync');         // { dirty: { 'step9.otHours': true }, meta: { 'step9.otHours': { by, at } } }
  var PUSH_DELAY = 1000;
  var TIMEOUT = 10000;

  // 舊的設定組名稱 → 現在的（組名統一用步驟編號）。本機備份和雲端的舊資料都照這張表搬
  // Dipping_Milling：第 6 步（需求工時匯總）、第 7 步（RMA）、舊總表、RMA人力需求已經刪除，
  // 那些步驟留下的舊資料（step5、step6、report、rma-manpower）對到 null＝丟掉，
  // 不然會和新的 step5（計算需求工時）、step6（匯入人力報表）撞名
  var LEGACY_GROUPS = {
    totals: 'step2', columns: 'step2-5', step3: 'step4', step4: 'step5', manpower: 'step6', 'report-dm': 'step7',
    step5: null, step6: null, report: null, 'rma-manpower': null
  };
  // 新格式裡改過名的組（版本 28 用過 shared，版本 29 起改成寫明是哪兩步共用）
  var RENAMED_GROUPS = { shared: 'step2-5' };
  var LAYOUT_KEY = IEApp.storeKey('_layout');      // 本機備份的格式版本
  var LAYOUT = 3;                                  // 2 = 組名用步驟編號、一個輸入項目一筆；3 = shared 改名 step2-5

  var client = null;           // supabase-js 的連線
  var session = null;          // supabase-js 的登入資訊（session.user.email 是登入者）
  var loggingOut = false;
  var started = false;         // 開啟時的登入、同步做完了沒
  var SLOW_START = 4000;       // 登入過的人，開啟時最多等這麼久；連不上就先用這台電腦的設定
  var pushTimer = null;
  var pushing = null;
  var status = { state: 'idle', text: '' };
  var els = {};

  /* ---------- 本機儲存 ---------- */
  function readLocal(k) {
    try { return localStorage.getItem(k); } catch (e) { return null; }
  }
  function writeLocal(k, v) {
    try { if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) { /* 存不了就算了 */ }
  }
  function readJson(k, fallback) {
    try { var v = JSON.parse(readLocal(k) || 'null'); return v === null ? fallback : v; } catch (e) { return fallback; }
  }
  function syncState() {
    var s = readJson(SYNC_KEY, {}) || {};
    s.dirty = s.dirty || {};
    s.meta = s.meta || {};
    return s;
  }
  function saveSyncState(s) { writeLocal(SYNC_KEY, JSON.stringify(s)); }

  // 這台電腦上這個專案的設定組（step1、step9…，不含 _ 開頭的內部資料）
  function localGroups() {
    var prefix = IEApp.storeKey('');
    var out = [];
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf(prefix) === 0 && k.charAt(prefix.length) !== '_') out.push(k.slice(prefix.length));
      }
    } catch (e) { /* 不能用 localStorage */ }
    return out;
  }
  function readGroup(group) {
    var v = readJson(IEApp.storeKey(group), null);
    return v && typeof v === 'object' && !Array.isArray(v) ? v : {};
  }
  function itemKey(group, field) { return group + '.' + field; }
  function splitKey(key) {
    var i = String(key).indexOf('.');
    return i < 0 ? null : { group: key.slice(0, i), field: key.slice(i + 1) };
  }
  function same(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
  function inGroups(key, groups) {
    var k = splitKey(key);
    return !!k && groups.indexOf(k.group) >= 0;
  }

  /* ---------- 各步驟用的設定 API ---------- */
  var settings = {
    // 回傳這一步之前存的物件；沒有就回傳 null
    read: function (group) { return readJson(IEApp.storeKey(group), null); },
    // 存下這一步的設定：本機馬上存好；有改到的輸入項目稍後在背景上傳
    write: function (group, value) {
      var before = readGroup(group);
      if (same(before, value)) return;                 // 內容沒變（例如打開畫面時重存一次）就不算修改
      writeLocal(IEApp.storeKey(group), JSON.stringify(value));
      if (!ENABLED || !value || typeof value !== 'object') return;
      var s = syncState();
      var changed = false;
      Object.keys(value).forEach(function (field) {
        if (!same(before[field], value[field])) { s.dirty[itemKey(group, field)] = true; changed = true; }
      });
      if (!changed) return;
      saveSyncState(s);
      updateNotes();
      schedulePush();
    },
    // 這幾步的設定最後是誰、什麼時候改的（取最新的一個輸入項目）；沒有回傳 null
    lastChange: function (groups) {
      var meta = syncState().meta;
      var best = null;
      Object.keys(meta).forEach(function (key) {
        var m = meta[key];
        if (inGroups(key, groups || []) && m && m.at && (!best || m.at > best.at)) best = m;
      });
      return best;
    }
  };

  /* ---------- 連線 ---------- */
  // 每個請求最多等 TIMEOUT 毫秒，連不上時才不會一直卡在「同步中」
  function timedFetch(input, init) {
    init = init || {};
    if (init.signal || typeof AbortController === 'undefined') return fetch(input, init);
    var ctrl = new AbortController();
    var timer = setTimeout(function () { ctrl.abort(); }, TIMEOUT);
    return fetch(input, Object.assign({}, init, { signal: ctrl.signal })).then(function (res) {
      clearTimeout(timer); return res;
    }, function (e) { clearTimeout(timer); throw e; });
  }

  function createClient() {
    return LIB.createClient(URL_BASE, KEY, {
      db: { schema: SCHEMA },
      auth: { storageKey: AUTH_KEY, persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
      global: { fetch: timedFetch }
    });
  }

  // 把 supabase-js 回傳的錯誤整理成：network（連不上）、auth（要重新登入）或一般錯誤
  function toError(error, status) {
    var msg = (error && (error.message || error.msg)) || String(error || '');
    var err = new Error(msg);
    err.status = status || (error && error.status) || 0;
    err.code = error && error.code;
    if (err.status === 0 || (error && error.name === 'AuthRetryableFetchError') ||
        /Failed to fetch|NetworkError|Load failed|AbortError|FetchError|fetch failed/i.test(msg)) {
      err.network = true;
      err.message = '連不上 Supabase（網路中斷或被擋）';
    } else if (err.status === 401 || err.code === 'PGRST301' || err.code === 'PGRST303' || /JWT/i.test(msg)) {
      err.auth = true;
    }
    return err;
  }

  // 呼叫資料表（query 是回傳 supabase-js 查詢的函式）；登入憑證失效時換一次再試
  function run(query, retried) {
    if (!session) return Promise.reject(Object.assign(new Error('沒有登入'), { auth: true }));
    return query().then(function (r) {
      if (!r.error) return r.data;
      var e = toError(r.error, r.status);
      if (e.auth && !retried) {
        return client.auth.refreshSession().then(function (x) {
          if (x.error) throw Object.assign(toError(x.error), x.error.name === 'AuthRetryableFetchError' ? {} : { auth: true });
          session = x.data.session;
          return run(query, true);
        });
      }
      throw e;
    }, function (e) { throw toError(e); });
  }

  /* ---------- 登入 ---------- */
  function userEmail() { return session && session.user ? session.user.email || '' : ''; }

  // 離線時 supabase-js 拿不到登入資訊，從它存在本機的資料讀出上次的登入者
  function storedEmail() {
    var v = readJson(AUTH_KEY, null);
    return v && v.user ? v.user.email || '' : '';
  }

  function signIn(email, password) {
    return client.auth.signInWithPassword({ email: email, password: password }).then(function (r) {
      if (r.error) throw toError(r.error);
      session = r.data.session;
      renderUser();
    });
  }

  // 版本 31 以前的登入資訊交給 supabase-js，已經登入的人不用重新登入
  function adoptOldSession() {
    var old = readJson(OLD_SESSION_KEY, null);
    if (!old || !old.refresh_token) return Promise.resolve(null);
    return client.auth.setSession({ access_token: old.access_token, refresh_token: old.refresh_token }).then(function (r) {
      if (r.error && r.error.name === 'AuthRetryableFetchError') return null;   // 離線：下次有網路再搬
      writeLocal(OLD_SESSION_KEY, null);
      return r.error ? null : r.data.session;
    }, function () { return null; });
  }

  // 開啟時拿登入資訊：'ok'、'offline'（有登入過但連不上）或 'login'（要登入）
  function restoreSession() {
    return client.auth.getSession().then(function (r) {
      if (r.data && r.data.session) { session = r.data.session; return 'ok'; }
      if (r.error && toError(r.error).network) return 'offline';
      return adoptOldSession().then(function (s) {
        if (s) { session = s; return 'ok'; }
        return readJson(OLD_SESSION_KEY, null) ? 'offline' : 'login';
      });
    }, function (e) { return toError(e).network && storedEmail() ? 'offline' : 'login'; });
  }

  /* ---------- 同步 ---------- */
  // 把 Supabase 上這個 PWA 的設定抓下來：每個輸入項目，本機沒有「還沒上傳」的修改就用雲端的；
  // 本機有、雲端沒有的輸入項目，標成要上傳（第一次使用時把這台電腦原本的設定傳上去）。
  function pull() {
    setStatus('busy', '同步中…');
    return run(function () {
      var q = client.from('settings').select('key,value,updated_at,updated_by').eq('app', APP);
      // 連不上時不要讓 supabase-js 自己重試好幾次（離線打開會多等好幾秒）；連上網路時這裡會自己再同步
      return typeof q.retry === 'function' ? q.retry(false) : q;
    })
      .then(function (rows) {
        var s = syncState();
        var groups = {};                               // 要寫回本機的設定組
        var remote = {};
        function group(g) { if (!groups[g]) groups[g] = readGroup(g); return groups[g]; }
        // 舊格式（一步一筆，key 沒有「.」）：轉成一個項目一筆，並標成要用新格式上傳；雲端已有新格式的項目以新格式為準
        var legacy = (rows || []).filter(function (r) { return !splitKey(r.key); });
        var renamed = (rows || []).filter(function (r) { var k = splitKey(r.key); return k && RENAMED_GROUPS[k.group]; });
        var current = (rows || []).filter(function (r) { var k = splitKey(r.key); return k && !RENAMED_GROUPS[k.group]; });
        current.forEach(function (r) { remote[r.key] = true; });
        // 改過名的組（例如 shared.wipCol → step2-5.wipCol）：雲端還沒有新名稱的項目才用舊的，並用新名稱上傳
        renamed.forEach(function (r) {
          var k = splitKey(r.key);
          var g = RENAMED_GROUPS[k.group];
          var key = itemKey(g, k.field);
          if (remote[key] || s.dirty[key]) return;
          group(g)[k.field] = r.value;
          s.dirty[key] = true;
        });
        legacy.forEach(function (r) {
          if (LEGACY_GROUPS[r.key] === null) return;   // 已刪除步驟的舊資料，丟掉
          var g = LEGACY_GROUPS[r.key] || r.key;
          var v = r.value && typeof r.value === 'object' && !Array.isArray(r.value) ? r.value : {};
          Object.keys(v).forEach(function (field) {
            var key = itemKey(g, field);
            if (remote[key] || s.dirty[key]) return;
            group(g)[field] = v[field];
            s.dirty[key] = true;
          });
        });
        current.forEach(function (r) {
          var k = splitKey(r.key);
          if (!s.dirty[r.key]) group(k.group)[k.field] = r.value;
          s.meta[r.key] = { by: r.updated_by || '', at: r.updated_at || '' };
        });
        Object.keys(groups).forEach(function (g) { writeLocal(IEApp.storeKey(g), JSON.stringify(groups[g])); });
        localGroups().forEach(function (g) {
          var v = readGroup(g);
          Object.keys(v).forEach(function (field) {
            var key = itemKey(g, field);
            if (!remote[key] && v[field] !== null && v[field] !== undefined) s.dirty[key] = true;
          });
        });
        saveSyncState(s);
        updateNotes();
        return push();
      })
      .catch(handleError);
  }

  function schedulePush() {
    clearTimeout(pushTimer);
    pushTimer = setTimeout(function () { push().catch(function () {}); }, PUSH_DELAY);
  }

  // 把「還沒上傳」的設定傳到 Supabase
  function push() {
    if (pushing) return pushing.then(push);
    var s = syncState();
    var names = Object.keys(s.dirty);
    if (!names.length) { setStatus('ok', '已同步'); return Promise.resolve(); }
    if (!session) return Promise.resolve();
    var rows = [];
    var skipped = [];
    names.forEach(function (name) {
      var k = splitKey(name);
      var value = k ? readGroup(k.group)[k.field] : undefined;
      if (value === null || value === undefined) skipped.push(name);      // 已經沒有這個項目了，不用傳
      else rows.push({ app: APP, key: name, value: value });
    });
    if (!rows.length) {
      var st0 = syncState();
      skipped.forEach(function (n) { delete st0.dirty[n]; });
      saveSyncState(st0);
      setStatus('ok', '已同步');
      return Promise.resolve();
    }
    setStatus('busy', '同步中…');
    pushing = run(function () {
      return client.from('settings').upsert(rows, { onConflict: 'app,key' }).select('key,updated_at,updated_by');
    }).then(function (saved) {
      var st = syncState();
      skipped.forEach(function (n) { delete st.dirty[n]; });
      rows.forEach(function (sent) {
        // 上傳的這段時間又改了的，留著下一輪再傳
        var k = splitKey(sent.key);
        if (same(sent.value, readGroup(k.group)[k.field])) delete st.dirty[sent.key];
      });
      (saved || []).forEach(function (row) { st.meta[row.key] = { by: row.updated_by || '', at: row.updated_at || '' }; });
      saveSyncState(st);
      updateNotes();
      pushing = null;
      if (Object.keys(st.dirty).length) return push();
      setStatus('ok', '已同步');
    }, function (e) {
      pushing = null;
      return handleError(e);
    });
    return pushing;
  }

  function handleError(e) {
    if (e && e.network) setStatus('offline', '離線：設定先存在這台電腦，連上網路後會自動上傳');
    else if (e && e.auth) { setStatus('error', '登入已過期，請重新登入'); showLogin(true); }
    else setStatus('error', '同步失敗：' + (e && e.message ? e.message : e));
  }

  /* ---------- 畫面：上方的登入者與同步狀態 ---------- */
  function setStatus(state, text) {
    status = { state: state, text: text };
    if (els.status) {
      els.status.textContent = text;
      els.status.className = 'cloud-status is-' + state;
    }
  }

  function renderUser() {
    if (!els.user) return;
    var email = userEmail() || storedEmail();
    els.user.textContent = email ? shortName(email) : '';
    els.user.title = email;
    els.logout.hidden = !email;
  }

  function buildTopbar() {
    var bar = document.querySelector('.topbar');
    if (!bar || els.box) return;
    els.status = h('span', { class: 'cloud-status', 'aria-live': 'polite' });
    els.user = h('span', { class: 'cloud-user' });
    els.logout = h('button', { type: 'button', class: 'cloud-logout', text: '登出', hidden: true, onclick: logout });
    els.box = h('span', { class: 'cloud-box' }, els.status, els.user, els.logout);
    var ver = document.getElementById('app-version');
    bar.insertBefore(els.box, ver);
    setStatus(status.state, status.text);
    renderUser();
  }

  function shortName(email) { return String(email || '').split('@')[0]; }

  function formatTime(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  // 各步驟上方「上次由誰修改」那一行
  function noteFor(names) {
    var el = h('p', { class: 'sync-note' });
    el.setAttribute('data-keys', names.join(','));
    fillNote(el);
    return el;
  }
  function fillNote(el) {
    var names = (el.getAttribute('data-keys') || '').split(',').filter(Boolean);
    var m = settings.lastChange(names);
    var dirty = syncState().dirty;
    var waiting = Object.keys(dirty).some(function (key) { return inGroups(key, names); });
    var text = m ? '共用設定：上次由 ' + shortName(m.by) + ' 於 ' + formatTime(m.at) + ' 修改'
      : '共用設定：還沒有人存過';
    if (waiting) text += '（這台電腦有修改還沒上傳）';
    el.textContent = text;
  }
  function updateNotes() {
    Array.prototype.forEach.call(document.querySelectorAll('.sync-note'), fillNote);
  }

  /* ---------- 畫面：登入 ---------- */
  var loginWaiters = [];       // 等登入完成的地方（開啟時的流程、用到一半過期）

  function showLogin(expired) {
    setStatus('busy', '請登入');
    if (els.login) return new Promise(function (res) { loginWaiters.push(res); });
    var email = h('input', { id: 'login-email', type: 'email', class: 'in', autocomplete: 'username', required: true,
      value: userEmail() || storedEmail() });
    var pass = h('input', { id: 'login-pass', type: 'password', class: 'in', autocomplete: 'current-password', required: true });
    var msg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });
    var btn = h('button', { type: 'submit', class: 'btn-primary', text: '登入' });
    var form = h('form', { class: 'login-card', 'aria-labelledby': 'login-title' },
      h('h1', { id: 'login-title', text: (document.querySelector('.app-name') || {}).textContent || '登入' }),
      h('p', { class: 'login-sub', text: expired ? '登入已過期，請重新登入。輸入的設定都還在這台電腦。' : '請登入。部門共用同一份設定，誰最後存，大家看到的就是誰的。' }),
      h('label', { for: 'login-email', text: '帳號（email）' }), email,
      h('label', { for: 'login-pass', text: '密碼' }), pass,
      msg, btn);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      msg.hidden = true;
      btn.disabled = true;
      btn.textContent = '登入中…';
      signIn(email.value.trim(), pass.value).then(function () {
        hideLogin();
        var waiting = loginWaiters; loginWaiters = [];
        waiting.forEach(function (res) { res(); });
        if (started) pull();                // 開啟後才要求登入（過期、換了資料庫）：登入後補同步
      }).catch(function (err) {
        msg.textContent = err.network ? '連不上 Supabase，請確認網路後再試。'
          : (err.code === 'invalid_credentials' || err.status === 400) ? '帳號或密碼不對。'
          : '登入失敗：' + err.message;
        msg.hidden = false;
        btn.disabled = false;
        btn.textContent = '登入';
      });
    });
    els.login = h('div', { class: 'login-overlay', role: 'dialog', 'aria-modal': 'true' }, form);
    document.body.appendChild(els.login);
    setTimeout(function () { (email.value ? pass : email).focus(); }, 0);
    return new Promise(function (res) { loginWaiters.push(res); });
  }

  function hideLogin() {
    if (els.login) { els.login.remove(); els.login = null; }
  }

  function logout() {
    var pending = Object.keys(syncState().dirty).length;
    if (pending && !global.confirm('這台電腦還有 ' + pending + ' 組設定沒上傳，登出後要等下次登入才會上傳。確定要登出嗎？')) return;
    loggingOut = true;
    var done = function () { global.location.reload(); };
    // scope local：只登出這台電腦；連不上 Supabase 也會清掉本機的登入資訊
    client.auth.signOut({ scope: 'local' }).then(done, done);
  }

  /* ---------- 開啟時 ---------- */
  // 回傳 Promise：登入好、抓完雲端設定（或確定離線）之後才 resolve，接著才載入各步驟
  function start() {
    buildTopbar();
    if (!ENABLED) { setStatus('off', ''); return Promise.resolve(); }
    if (!LIB) {
      setStatus('error', '找不到 supabase-js（vendor/supabase.js），設定只存在這台電腦');
      return Promise.resolve();
    }
    client = createClient();
    client.auth.onAuthStateChange(function (event, s) {
      var had = !!session;
      if (s) session = s;
      // 登入被 Supabase 取消（例如帳號停用、在後台被登出）時請使用者重新登入
      // （開啟時的登入由 start 自己處理，這裡只管用到一半被登出）
      if (event === 'SIGNED_OUT' && !loggingOut) {
        session = null;
        if (started) setTimeout(function () { handleError({ auth: true }); }, 0);
        return;
      }
      // 離線打開、之後 supabase-js 自己換到新的登入憑證：補做同步
      if (!had && s && started) setTimeout(function () { renderUser(); pull(); }, 0);
    });
    renderUser();
    var restoring = restoreSession();
    // 登入過的人：連不上 Supabase 時（supabase-js 會一直重試換憑證）不要卡著，先用這台電腦的設定
    var first = storedEmail() || readJson(OLD_SESSION_KEY, null)
      ? Promise.race([restoring, new Promise(function (res) { setTimeout(function () { res('slow'); }, SLOW_START); })])
      : restoring;
    return first.then(function (r) {
      renderUser();
      if (r === 'offline' || r === 'slow') {
        setStatus('offline', '離線：設定先存在這台電腦，連上網路後會自動上傳');
        if (r === 'slow') restoring.then(function (r2) { afterRestore(r2); });
        return 'offline';
      }
      if (r === 'login') return showLogin(false);
    }).then(function (r) {
      if (r === 'offline') return;
      return pull();
    }).then(function () {
      started = true;
      global.addEventListener('online', function () {
        if (session) { pull(); return; }
        restoreSession().then(afterRestore);
      });
    });
  }

  // 開啟後才拿到結果的登入狀態（原本離線）
  function afterRestore(r) {
    if (r === 'ok' && session) { renderUser(); pull(); }
    else if (r === 'login') showLogin(true);
  }

  // 本機備份從舊組名搬到新組名（只做一次；要先全部讀出來再寫，因為 step3→step4、step4→step5 會接在一起）
  (function migrateLayout() {
    var layout = readJson(LAYOUT_KEY, 0);
    if (layout >= LAYOUT) return;
    if (layout === 2) {                            // 版本 28 → 29：只有 shared 改名
      Object.keys(RENAMED_GROUPS).forEach(function (name) {
        var raw = readLocal(IEApp.storeKey(name));
        var to = IEApp.storeKey(RENAMED_GROUPS[name]);
        if (raw !== null && readLocal(to) === null) writeLocal(to, raw);
        writeLocal(IEApp.storeKey(name), null);
      });
      writeLocal(SYNC_KEY, null);
      writeLocal(LAYOUT_KEY, String(LAYOUT));
      return;
    }
    var old = {};
    Object.keys(LEGACY_GROUPS).forEach(function (name) {
      var raw = readLocal(IEApp.storeKey(name));
      if (raw !== null) old[name] = raw;
    });
    Object.keys(old).forEach(function (name) { writeLocal(IEApp.storeKey(name), null); });
    Object.keys(old).forEach(function (name) {
      if (!LEGACY_GROUPS[name]) return;            // 對到 null：已刪除步驟的舊資料，只刪不搬
      var to = IEApp.storeKey(LEGACY_GROUPS[name]);
      if (readLocal(to) === null) writeLocal(to, old[name]);
    });
    writeLocal(SYNC_KEY, null);                    // 舊格式的同步紀錄不再適用，下次同步重新整理
    writeLocal(LAYOUT_KEY, String(LAYOUT));
  })();

  IEApp.settings = settings;
  IEApp.cloud = {
    enabled: ENABLED,
    start: start,
    noteFor: noteFor,
    pull: pull,
    push: push,
    status: function () { return status; }
  };
})(window);
