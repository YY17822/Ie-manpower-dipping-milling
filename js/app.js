/*
 * app.js — 外框：左側步驟清單、切換步驟、各步驟之間共用的資料、載入 ExcelJS。
 * 每個步驟是 js/steps/ 底下的一個檔案，用 IEApp.registerStep() 掛進來。
 */
(function (global) {
  'use strict';

  var EXCELJS_LOCAL = 'vendor/exceljs.min.js';
  var EXCELJS_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';

  // 整個流程的步驟。還沒開發的步驟先列出來，做到時再補上對應的檔案。
  var STEPS = [
    { id: 'mps-filter', title: '篩選 MPS 報表', note: '挑出要算的料號與月份' },
    { id: 'mps-totals', title: '匯總', note: 'Cutting、Dipping 各月 Total' },
    { id: 'p003-import', title: '匯入 P003 標準工時', note: '加到同一個檔案' },
    { id: 'ct-lookup', title: '帶入標準工時 CT.', note: '機台 CT.、目檢 CT.、檢驗率' },
    { id: 'demand-hours', title: '計算需求工時', note: '機台、目檢需求工時' },
    { id: 'manpower', title: '匯入人力報表', note: '未來六個月的人力預計' },
    { id: 'report', title: '總表', note: '各月人力需求總覽' }
  ];

  /* ---------- 這個專案的身分（js/config.js） ---------- */
  var CONFIG = global.APP_CONFIG || {};
  var PROJECT_ID = CONFIG.PROJECT_ID || 'ie-manpower';

  // 記住的設定（localStorage）的名稱一律加上專案代號，例如 ie-manpower-dipping-milling.step1，
  // 同一個網址底下的其他專案才不會讀到或蓋掉這裡的設定。
  function storeKey(name) { return PROJECT_ID + '.' + name; }

  // 舊版的設定名稱（ie-manpower.xxx）搬到新名稱，搬完把舊的刪掉。
  // 一定要在各步驟讀設定之前做，所以放在 app.js 一載入就執行。
  (function migrateOldSettings() {
    var old = CONFIG.LEGACY_STORE_PREFIX;
    if (!old) return;
    try {
      var keys = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf(old) === 0) keys.push(k);
      }
      keys.forEach(function (k) {
        var nk = storeKey(k.slice(old.length));
        if (localStorage.getItem(nk) === null) localStorage.setItem(nk, localStorage.getItem(k));
        localStorage.removeItem(k);
      });
    } catch (e) { /* 瀏覽器不讓用 localStorage 就算了 */ }
  })();

  // 各步驟的程式。登入、抓完共用設定之後才載入（各步驟一載入就會讀設定）
  var STEP_FILES = [
    'js/steps/step1-mps-filter.js',
    'js/steps/step2-mps-totals.js',
    'js/steps/step3-p003-import.js',
    'js/steps/step4-ct-lookup.js',
    'js/steps/step5-demand-hours.js',
    'js/steps/step6-manpower.js',
    'js/steps/step7-report.js'
  ];
  // 各步驟用到的共用設定（畫面上「上次由誰修改」看這幾個）
  var STEP_SETTINGS = {
    'mps-filter': ['step1'],
    'mps-totals': ['step2', 'step2-5'],
    'ct-lookup': ['step4'],
    'demand-hours': ['step5', 'step2-5'],
    'manpower': ['step6'],
    'report': ['step7']
  };

  var registry = {};
  var stepsLoaded = false;       // 各步驟的程式載入了沒（登入前還沒載入）
  var currentId = null;

  /* ---------- 小工具：建立 DOM ---------- */
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach(function (k) {
        var v = attrs[k];
        if (v === null || v === undefined || v === false) return;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (k.slice(0, 2) === 'on') el.addEventListener(k.slice(2), v);
        else el.setAttribute(k, v === true ? '' : v);
      });
    }
    for (var i = 2; i < arguments.length; i++) {
      var child = arguments[i];
      if (child === null || child === undefined || child === false) continue;
      if (Array.isArray(child)) child.forEach(function (c) { if (c) el.appendChild(c); });
      else el.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
    }
    return el;
  }

  /* ---------- 載入 ExcelJS：先找資料夾裡的副本，沒有才連網路 ---------- */
  function loadScript(src, crossOrigin) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = src;
      if (crossOrigin) s.crossOrigin = 'anonymous';
      s.onload = resolve;
      s.onerror = function () { s.remove(); reject(new Error(src)); };
      document.head.appendChild(s);
    });
  }

  var excelPromise = null;
  function loadExcelJS() {
    if (global.ExcelJS) return Promise.resolve(global.ExcelJS);
    if (!excelPromise) {
      excelPromise = loadScript(EXCELJS_LOCAL)
        .catch(function () { return loadScript(EXCELJS_CDN, true); })
        .then(function () {
          if (!global.ExcelJS) throw new Error('ExcelJS');
          return global.ExcelJS;
        })
        .catch(function () {
          excelPromise = null;
          throw new Error('載入不了 Excel 元件（ExcelJS）。第一次使用需要連上網路；' +
            '如果公司網路擋外部網站，請把 exceljs.min.js 放進 vendor 資料夾（做法見 README）。');
        });
    }
    return excelPromise;
  }

  /* ---------- 步驟清單 ---------- */
  function renderStepList() {
    var list = document.getElementById('step-list');
    list.textContent = '';
    STEPS.forEach(function (step, i) {
      var ready = !!registry[step.id];
      var active = step.id === currentId;
      var body = [
        h('span', { class: 'step-num', text: String(i + 1) }),
        h('span', { class: 'step-text' },
          h('span', { class: 'step-title', text: step.title }),
          h('span', { class: 'step-note', text: ready ? step.note : stepsLoaded ? '尚未開發' : '' }))
      ];
      var item = ready
        ? h('button', {
            type: 'button', class: 'step' + (active ? ' is-active' : ''),
            'aria-current': active ? 'step' : null,
            onclick: function () { go(step.id); }
          }, body)
        : h('div', { class: 'step is-todo' }, body);
      list.appendChild(h('li', null, item));
    });
  }

  function go(id) {
    if (!registry[id]) return;
    currentId = id;
    renderStepList();
    var panel = document.getElementById('step-panel');
    panel.textContent = '';
    registry[id].render(panel, IEApp);
    // 共用設定：上次由誰修改（放在步驟標題下面）
    var names = STEP_SETTINGS[id];
    if (names && IEApp.cloud && IEApp.cloud.enabled) {
      var head = panel.querySelector('.step-head');
      var note = IEApp.cloud.noteFor(names);
      if (head) head.appendChild(note); else panel.insertBefore(note, panel.firstChild);
    }
  }

  // 依序載入各步驟的程式（?v= 和版本號一樣，瀏覽器才不會用舊檔）
  function loadSteps() {
    var v = (global.ExcelIO && global.ExcelIO.VERSION) || '';
    return STEP_FILES.reduce(function (p, src) {
      return p.then(function () {
        return loadScript(src + '?v=' + v).catch(function () {
          throw new Error('載入不了 ' + src);
        });
      });
    }, Promise.resolve());
  }

  function registerStep(step) { registry[step.id] = step; }

  function start() {
    var ver = document.getElementById('app-version');
    if (ver) ver.textContent = '版本 ' + ((global.ExcelIO && global.ExcelIO.VERSION) || '?');
    renderStepList();
    loadExcelJS().catch(function () { /* 上傳檔案時會再試一次並顯示訊息 */ });
    var cloud = IEApp.cloud ? IEApp.cloud.start() : Promise.resolve();
    cloud.then(loadSteps).then(function () {
      stepsLoaded = true;
      var first = STEPS.filter(function (s) { return registry[s.id]; })[0];
      renderStepList();
      if (first) go(first.id);
    }).catch(function (err) {
      var panel = document.getElementById('step-panel');
      panel.textContent = '';
      panel.appendChild(h('p', { class: 'msg is-error', text: '程式載入失敗：' + (err && err.message ? err.message : err) + '。請重新整理再試一次。' }));
    });
    // 離線快取只在用 http(s) 開啟時才有（直接點兩下 index.html 開啟也能用，只是不能安裝）
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      navigator.serviceWorker.register('sw.js').catch(function () {});
    }
  }

  var IEApp = {
    state: {},          // 各步驟的成果放這裡，下一步直接拿來用
    projectId: PROJECT_ID,
    storeKey: storeKey,
    h: h,
    registerStep: registerStep,
    go: go,
    loadExcelJS: loadExcelJS
  };
  global.IEApp = IEApp;

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})(window);
