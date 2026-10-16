/*
 * 第 9 步：總表
 *   在活頁簿裡新增「總表」工作表。目前先建立表格的樣式、格式、外框格線和位置，
 *   各格的內容（公式）之後再一項一項補上。
 *
 *   版面（B 到 J 行、第 2 到 24 列）：
 *     第 2 列      法定加班上限46hr/人（紅底白字）
 *     第 3 列      表頭：B、C 合併畫斜線；D 到 I 是六個月；J 是備註（深藍底白字）
 *     第 4 列      MPS (Production)
 *     第 5–14 列   需求工時、工作天數、工作時間、加班時間、綜合效率、出勤率
 *     第 15–18 列  需求人力（On line、Off line_直接、Off line_間接、小計）
 *     第 19–22 列  現有人力（早+夜）（同上，小計(on+off)）
 *     第 23 列     人力需求差異(現有-需求)
 *     第 24 列     Loading Rate%(含加班)
 *
 *   月份表頭用公式參照 MPS 工作表「需求工時」的月份（第 5 步輸入的），例如 ='MPS'!N3。
 *
 *   已經填好的列（都是公式，參照 MPS 工作表；用標籤文字找出是哪一列）：
 *     MPS (Production)      ← 第 2 步匯總的「Total」那一列（當月已含 WIP），例如 ='MPS'!E76
 *     需求工時(Top 3)        ← 「Top 3 總需求工時」那一列，例如 ='MPS'!N77
 *     需求工時(Non-Top 3)    ← 「Others 總需求工時」那一列
 *     需求工時（重工）        ← 「RMA總需求工時」那一列
 *
 *   使用者在畫面上輸入的列：
 *     平日工作天數     各月自己輸入（每個月不同，不記住）
 *     工作時間         各月自己輸入，預設 480 分
 *     加班時間         輸入每日加班幾小時、加班出勤人數 %，填公式 =2*60*80%，備註寫輸入的數值
 *   由輸入算出來的列：
 *     假日工作天數     =(46-D8*2)/8（法定加班上限 46 小時 − 平日天數 × 每天加班 2 小時）÷ 每天 8 小時
 *     綜合效率%（Top-3 機種）、綜合效率%（Non-Top 3 機種）、出勤率%   各月自己輸入，預設 80%、95%、95%
 *   工作時間、效率、出勤率、加班的兩個數字會記在這台電腦的瀏覽器裡，下次打開還是上次輸入的值。
 *     需求人力 Off line_直接   輸入文書、物料、材製損、包裝、領班的人數（每個月一樣），填公式 =2+3+1+4+1，備註寫各種人數
 *     需求人力 Off line_間接   輸入物料、領班的人數，同上
 *   算出來的列：
 *     需求人力 On line   =D5/(D8+D9)/(D10+D11)/D12/D14+D6/(D8+D9)/(D10+D11)/D13/D14+D7/D8/D10/D13/D14
 *     需求人力 小計      =D15+D16（On line + Off line_直接，不含 Off line_間接，備註有寫）
 *   從「人力」工作表抓的列（第 8 步確認的列和當月那一行，例如 ='人力'!D4）：
 *     現有人力 On line、Off line_直接、Off line_間接
 *     現有人力 小計(on+off)  =D19+D20（不含 Off line_間接，備註有寫）
 *   上傳第 8 步的檔案時，用第 8 步上次確認、記在瀏覽器裡的位置。
 *   最後兩列：
 *     人力需求差異(現有-需求)   =D22-D18
 *     Loading Rate%(含加班)     =D22/D18（顯示成 %）
 *
 *   第二張表（第 24 列下面空三列，第 28 到 50 列）：和第一張一樣，紅色格子寫「YO加班上限78hr/人」，
 *   假日工作天數改成 =(78-D34*2)/8。78 可以在畫面上改（會記住），紅色格子和公式一起變。
 *   其他輸入和抓來的數值都參照第一張表（例如 D34 =D8），需求人力、差異、Loading Rate 用第二張表自己的數字算。
 *
 *   來源有兩種：
 *     a. 剛才第 8 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 8 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withReport。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = XIO.REPORT_SHEET;
  var STEP1_KEY = IEApp.storeKey('step1');
  var TITLE = '需求工時';
  var SUFFIX = '_總表';
  var EXTRA = ['P003', '人力', SHEET_NAME];
  var STORE_KEY = IEApp.storeKey('report');
  var MONTHS = XIO.REPORT.months;
  var DEFAULT_WORK_MINUTES = '480';
  // 需求人力 Off line 的兩組人數（人員種類在 excel-io.js 的 REPORT_OFFLINE）
  var OFFLINE = [
    { key: 'offDirect', label: '需求人力 Off line_直接' },
    { key: 'offIndirect', label: '需求人力 Off line_間接' }
  ];

  var S = {
    upload: null,         // 上傳的第 8 步檔案 { name, buffer, sheets }
    preferUpload: false,
    work: null,           // 拿來預覽的活頁簿 { src, loading, error, wb, main }
    inputs: null          // 輸入框裡的文字 { weekdays: [6], workMinutes: [6], effTop3: [6], effNonTop3: [6], attendance: [6], otHours, otRate }
  };
  var els = {};
  var timer = null;

  function blank6(v) { var a = []; for (var i = 0; i < MONTHS; i++) a.push(v); return a; }

  // 各月自己輸入、會記住上次的值的列（percent: 輸入的是 %，例如 80 代表 80%）
  var REMEMBERED = [
    { key: 'workMinutes', label: '工作時間（分/天）', what: '工作時間', def: DEFAULT_WORK_MINUTES },
    { key: 'effTop3', label: '綜合效率%（Top-3 機種）', what: '綜合效率%（Top-3 機種）', def: '80', percent: true },
    { key: 'effNonTop3', label: '綜合效率%（Non-Top 3 機種）', what: '綜合效率%（Non-Top 3 機種）', def: '95', percent: true },
    { key: 'attendance', label: '出勤率%', what: '出勤率%', def: '95', percent: true }
  ];

  function loadInputs() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var out = {
      weekdays: blank6(''),
      otHours: s.otHours === undefined ? '' : String(s.otHours),
      otRate: s.otRate === undefined ? '' : String(s.otRate),
      yoLimit: s.yoLimit === undefined || s.yoLimit === '' ? String(XIO.YO_LIMIT_DEFAULT) : String(s.yoLimit)
    };
    OFFLINE.forEach(function (g) {
      var saved = Array.isArray(s[g.key]) ? s[g.key] : [];
      out[g.key] = XIO.REPORT_OFFLINE[g.key].map(function (c, i) { return saved[i] === undefined || saved[i] === null ? '' : String(saved[i]); });
    });
    REMEMBERED.forEach(function (row) {
      var saved = Array.isArray(s[row.key]) ? s[row.key].slice(0, MONTHS) : [];
      while (saved.length < MONTHS) saved.push(row.def);
      out[row.key] = saved.map(function (v) { return v === undefined || v === null ? row.def : String(v); });
    });
    return out;
  }
  function saveInputs() {
    var i = S.inputs;
    var keep = { otHours: i.otHours, otRate: i.otRate, offDirect: i.offDirect, offIndirect: i.offIndirect, yoLimit: i.yoLimit };
    REMEMBERED.forEach(function (row) { keep[row.key] = i[row.key]; });
    try { localStorage.setItem(STORE_KEY, JSON.stringify(keep)); } catch (e) { /* 存不了就算了 */ }
  }

  // 輸入框的文字 → 數字。回傳 { values, error, blanks: [說明…] }
  function readInputs() {
    var i = S.inputs;
    var out = { weekdays: [], otHours: null, otRate: null };
    var blanks = [];
    function num(text, what, percent) {
      var t = String(text);
      if (percent) t = t.replace(/[%％]\s*$/, '');      // 允許打成 80%
      var n = XIO.toNumber(t);
      if (n !== null && isNaN(n)) throw new Error(what + '不是數字。');
      return n;
    }
    try {
      for (var m = 0; m < MONTHS; m++) {
        out.weekdays.push(num(i.weekdays[m], monthName(m) + '的平日工作天數'));
      }
      REMEMBERED.forEach(function (row) {
        out[row.key] = i[row.key].map(function (v, m) { return num(v, monthName(m) + '的' + row.what, row.percent); });
      });
      out.otHours = num(i.otHours, '每日加班小時');
      out.otRate = num(i.otRate, '加班出勤人數 %', true);
      out.yoLimit = num(i.yoLimit, 'YO 加班上限');
      if (out.yoLimit === null) throw new Error('請輸入 YO 加班上限（小時）。');
      OFFLINE.forEach(function (g) {
        out[g.key] = XIO.REPORT_OFFLINE[g.key].map(function (c, k) { return num(i[g.key][k], g.label + '的' + c + '人數'); });
      });
    } catch (e) {
      return { error: e.message };
    }
    var missWd = out.weekdays.filter(function (v) { return v === null; }).length;
    if (missWd) blanks.push('有 ' + missWd + ' 個月還沒輸入平日工作天數（假日工作天數會照 0 天算）。');
    REMEMBERED.forEach(function (row) {
      var miss = out[row.key].filter(function (v) { return v === null; }).length;
      if (miss) blanks.push('有 ' + miss + ' 個月沒有' + row.what + '。');
    });
    if (out.otHours === null || out.otRate === null) blanks.push('每日加班小時和加班出勤人數 % 都輸入後，才會填加班時間。');
    OFFLINE.forEach(function (g) {
      if (!out[g.key].some(function (v) { return v !== null; })) blanks.push(g.label + '還沒輸入人數，這一列先留空白。');
    });
    if (missWd) blanks.push('平日工作天數沒填的月份，需求人力 On line 會是 #DIV/0!。');
    return { values: out, blanks: blanks };
  }

  function monthName(m) {
    var r = S.work && S.work.refs && S.work.refs[m];
    return r && r.text ? r.text : '第 ' + (m + 1) + ' 個月';
  }

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  function isExtra(name) { return EXTRA.some(function (n) { return sameName(n, name); }); }

  function activeSource() {
    var prev = IEApp.state.withManpower;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? Promise.resolve(src.prev.build()) : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isExtra(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：沿用前面步驟的；上傳的檔案用第 1 步記住的設定（預設第 5 列）
  function dataStartRow(src) {
    if (src.kind === 'prev' && src.prev.dataStartRow) return src.prev.dataStartRow;
    var n = 0;
    try {
      var s1 = JSON.parse(localStorage.getItem(STEP1_KEY) || '{}') || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  // 月份表頭要參照的儲存格（MPS 工作表「需求工時」那一塊的月份）
  function monthRefs(main, src) {
    var start = dataStartRow(src);
    var keyCol = src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;
    var plan = XIO.planSummary(main, { dataStartRow: start, keyCol: keyCol, title: TITLE });
    if (!plan.found) return null;
    var headRow = start - 1 >= 3 ? 3 : start - 1;
    if (headRow < 1) return null;
    var refs = [];
    for (var i = 0; i < Math.min(plan.width, XIO.REPORT.months); i++) {
      var c = plan.startCol + i;
      var cell = main.findCell(headRow, c);
      refs.push({ sheet: main.name, col: c, row: headRow, text: cell ? XIO.displayText(cell) : '' });
    }
    return refs;
  }

  // 現有人力：從「人力」工作表抓第 8 步確認過的那幾列
  var MP_SHEET = '人力';
  var MP_STORE_KEY = IEApp.storeKey('manpower');
  var MP_LINKS = [
    { key: 'haveOnline', item: 'online', label: 'Online' },
    { key: 'haveOffDirect', item: 'offDirect', label: 'Off line（直接）' },
    { key: 'haveOffIndirect', item: 'offIndirect', label: 'Off line（間接）' }
  ];

  // 第 8 步確認的位置：剛做完第 8 步就用它的；上傳的檔案用第 8 步記在瀏覽器裡的位置
  function manpowerPositions(src) {
    if (src.kind === 'prev' && src.prev.rows && src.prev.firstMonthCol) {
      return { rows: src.prev.rows, firstCol: src.prev.firstMonthCol, from: 'chain' };
    }
    var s = {};
    try { s = JSON.parse(localStorage.getItem(MP_STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var firstCol = parseInt(s.lastFirstCol, 10) || 0;
    if (!s.rows || !firstCol) return null;
    return { rows: s.rows, firstCol: firstCol, from: 'saved' };
  }

  // 回傳 { haveOnline: { found, label, sheet, row, refs, reason }, … }
  function manpowerLinks(wb, src) {
    var ws = wb.worksheets.filter(function (w) { return sameName(w.name, MP_SHEET); })[0];
    var pos = ws ? manpowerPositions(src) : null;
    var out = {};
    MP_LINKS.forEach(function (link) {
      var info = { found: false, label: link.label, sheet: MP_SHEET, refs: [], from: pos && pos.from };
      var r = pos ? parseInt(pos.rows[link.item], 10) : 0;
      if (!ws) info.reason = '檔案裡沒有「' + MP_SHEET + '」工作表（第 8 步）';
      else if (!pos || !r) info.reason = '不知道「' + link.label + '」在人力工作表的哪一列，請重做第 8 步';
      else {
        info.found = true;
        info.row = r;
        var row = ws.findRow(r);
        for (var i = 0; i < MONTHS; i++) {
          var c = pos.firstCol + i;
          var n = XIO.toNumber(row ? row.findCell(c) && XIO.plainValue(row.findCell(c).value) : null);
          info.refs.push({ sheet: ws.name, col: c, row: r, value: n === null || isNaN(n) ? null : n, numFmt: 'General' });
        }
      }
      out[link.key] = info;
    });
    return out;
  }

  // findReportLinks 的結果 → addReportSheet 要的 { key: refs }
  function linkRefs(links) {
    var out = {};
    Object.keys(links).forEach(function (k) { if (links[k].found) out[k] = links[k].refs; });
    return out;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f9-base',
      title: '上傳第 8 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 8 步匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    if (!S.inputs) S.inputs = loadInputs();
    els.entry = h('div', { class: 'entry-area' });
    els.monthHeads = [];
    buildEntry();

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.preview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：RMA人力需求',
      onclick: function () { IEApp.go('rma-manpower'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '總表' }),
      h('p', { text: '在同一個檔案裡新增「總表」工作表，用公式從 MPS、人力工作表抓產量、需求工時和現有人力，加上畫面上輸入的天數、工時、效率、出勤率和 Off line 人數，算出需求人力、人力差異和 Loading Rate。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 8 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '輸入各月的數值' }),
        h('p', { class: 'hint', text: '平日工作天數每個月都要輸入；工作時間預設 480 分，綜合效率預設 80%（Top-3）、95%（Non-Top 3），出勤率預設 95%，會記住上次輸入的值。假日工作天數會用公式自動算：（法定加班上限 46 小時 − 平日工作天數 × 每天加班 2 小時）÷ 8。加班時間 = 每日加班小時 × 60 × 加班出勤人數 %。Off line 的人數每個月一樣，會加總填入，各種人數寫在備註。' }),
        els.entry),
      h('li', { class: 'task' },
        h('h2', { text: '總表' }),
        els.summary, els.notes, els.result, els.preview,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    refresh();
  }

  /* ---------- 輸入表 ---------- */
  function input(value, label, onInput, cls) {
    var el = h('input', {
      type: 'text', class: 'in ' + (cls || 'entry-num'), inputmode: 'decimal', autocomplete: 'off',
      value: value, 'aria-label': label
    });
    el.addEventListener('input', function () { onInput(el.value); onType(); });
    return el;
  }

  function buildEntry() {
    var inp = S.inputs;
    var head = h('tr', null, h('th', { scope: 'col', text: '項目' }));
    for (var m = 0; m < MONTHS; m++) {
      var th = h('th', { scope: 'col', text: monthName(m) });
      els.monthHeads.push(th);
      head.appendChild(th);
    }
    var body = h('tbody');
    [{ key: 'weekdays', label: '平日工作天數（天/月）' }].concat(REMEMBERED).forEach(function (row) {
      var tr = h('tr', null, h('th', { scope: 'row', text: row.label }));
      inp[row.key].forEach(function (v, i) {
        tr.appendChild(h('td', null, input(v, row.label + '，' + monthName(i), function (val) { inp[row.key][i] = val; })));
      });
      body.appendChild(tr);
    });
    els.entry.appendChild(h('div', { class: 'entry-scroll' },
      h('table', { class: 'entry' }, h('thead', null, head), body)));
    els.entry.appendChild(h('div', { class: 'field-row entry-ot' },
      h('label', { for: 'f9-ot-hours', text: '加班時間：每日加班' }),
      (function () { var el = input(inp.otHours, '每日加班幾小時', function (v) { inp.otHours = v; }, 'in-tiny'); el.id = 'f9-ot-hours'; return el; })(),
      h('span', { text: '小時，加班出勤人數' }),
      (function () { var el = input(inp.otRate, '加班出勤人數 %', function (v) { inp.otRate = v; }, 'in-tiny'); el.id = 'f9-ot-rate'; return el; })(),
      h('span', { text: '%' })));

    // 第二張表：YO 加班上限
    els.entry.appendChild(h('div', { class: 'field-row entry-ot' },
      h('label', { for: 'f9-yo-limit', text: 'YO 加班上限（第二張表）：' }),
      (function () { var el = input(inp.yoLimit, 'YO 加班上限幾小時', function (v) { inp.yoLimit = v; }, 'in-tiny'); el.id = 'f9-yo-limit'; return el; })(),
      h('span', { text: '小時/人（預設 ' + XIO.YO_LIMIT_DEFAULT + '）' })));

    // 需求人力 Off line：各種人員的人數（每個月一樣）
    OFFLINE.forEach(function (g) {
      var row = h('div', { class: 'field-row entry-ot' }, h('span', { class: 'field-label', text: g.label + '：' }));
      XIO.REPORT_OFFLINE[g.key].forEach(function (cat, k) {
        var id = 'f9-' + g.key + '-' + k;
        var el = input(inp[g.key][k], g.label + ' ' + cat + ' 人數', function (v) { inp[g.key][k] = v; }, 'in-tiny');
        el.id = id;
        row.appendChild(h('label', { for: id, class: 'entry-cat', text: cat }));
        row.appendChild(el);
      });
      els.entry.appendChild(row);
    });
  }

  function onType() {
    saveInputs();
    els.exportNote.textContent = '';
    clearTimeout(timer);
    timer = setTimeout(refresh, 250);
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(function (n) { return !isExtra(n); })) {
          throw new Error('這個檔案沒有 MPS 工作表。請上傳第 8 步匯出的檔案。');
        }
        S.upload = { name: file.name, buffer: res.buffer, sheets: names };
        S.preferUpload = true;
        els.exportNote.textContent = '';
        refresh();
      })
      .catch(function (err) {
        els.srcMsg.textContent = err && err.message ? err.message : String(err);
        els.srcMsg.hidden = false;
        refresh();
      });
  }

  function switchSource() {
    S.preferUpload = !S.preferUpload;
    els.exportNote.textContent = '';
    refresh();
  }

  // 先產生一份活頁簿，加上總表，給畫面預覽
  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      var refs = monthRefs(main, src);
      work.refs = refs;
      work.links = XIO.findReportLinks(main, { dataStartRow: dataStartRow(src), months: XIO.REPORT.months });
      var mp = manpowerLinks(wb, src);
      Object.keys(mp).forEach(function (k) { work.links[k] = mp[k]; });
      work.wb = wb;
      work.main = main;
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = IEApp.state.withManpower;
    IEApp.state.withReport = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';
    els.preview.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 8 步的結果。回到第 8 步，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 8 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.upload.sheets.join('、') + '。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 8 步剛做好的結果' : '改用上傳的檔案';

    if (!src) { setSummary('先準備好第 8 步的 Excel。'); return; }
    ensureWork(src);
    if (S.work.loading) { setSummary('準備資料中…'); return; }
    if (S.work.error) { setSummary(S.work.error, true); return; }

    var work = S.work;
    // 輸入表的月份表頭
    els.monthHeads.forEach(function (th, m) { th.textContent = monthName(m); });
    var read = readInputs();
    if (read.error) { setSummary(read.error, true); return; }
    // 每次都照目前的輸入重新產生總表（給預覽用）
    work.report = XIO.addReportSheet(work.wb, { monthRefs: work.refs || [], links: linkRefs(work.links), inputs: read.values, yoLimit: read.values.yoLimit });
    var ws = work.report.worksheet;
    var names = work.wb.worksheets.map(function (w) { return w.name; });
    var replacing = src.kind === 'file' && S.upload.sheets.some(function (n) { return sameName(n, SHEET_NAME); });
    var text = '匯出的檔案有 ' + names.length + ' 個工作表：' + names.join('、') + '。「' + SHEET_NAME +
      '」有兩張表：上面是法定加班上限 46hr，下面（第 28 列起）是 YO 加班上限 ' + read.values.yoLimit + 'hr。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會重新建立。';
    setSummary(text, false, replacing);

    if (work.refs && work.refs.length) {
      var r0 = work.refs[0];
      els.notes.appendChild(h('li', null, '月份表頭參照 ' + work.main.name + ' 工作表「需求工時」的月份，例如 D3 ＝ ' +
        "='" + r0.sheet + "'!" + XIO.numToCol(r0.col) + r0.row + '（' + work.refs.map(function (r) { return r.text || '空白'; }).join('、') + '）。'));
    } else {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        work.main.name + ' 工作表裡找不到「需求工時」的月份，總表的月份表頭先留空白。'));
    }

    var warn = renderLinks(work);
    read.blanks.forEach(function (t) { warn = true; els.notes.appendChild(h('li', { class: 'is-warn' }, t)); });
    if (warn) els.summary.className = 'summary is-warn';

    ui.renderSheet(els.preview, ws, {
      headerEnd: XIO.REPORT.headRow,
      lastRow: XIO.REPORT.lastRow + XIO.REPORT_SECOND_OFFSET,
      colCount: XIO.REPORT.noteCol,
      rows: XIO.REPORT.lastRow + XIO.REPORT_SECOND_OFFSET,
      caption: '總表的版面（顏色、框線在 Excel 裡才看得到）。'
    });

    els.exportBtn.disabled = false;
    var refs = work.refs || [];
    var links = linkRefs(work.links);
    var inputs = read.values;
    IEApp.state.withReport = {
      outName: outputName(src),
      mainSheet: work.main.name,
      reportSheet: SHEET_NAME,
      reportRows: work.report.rows,
      monthCol: work.report.monthCol,
      monthCount: work.report.months,
      inputs: inputs,
      build: function () { return buildWorkbook(src, refs, links, inputs); }
    };
    els.nextBtn.disabled = false;
  }

  // 各列從哪裡抓、抓到的數值；有找不到的回傳 true
  function renderLinks(work) {
    var months = XIO.REPORT.months;
    var rowLabel = {};
    XIO.REPORT_ROWS.forEach(function (d) { rowLabel[d.key] = d.label || d.unit; });
    var warn = false;

    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '總表' }),
      h('th', { class: 'colhead', scope: 'col', text: '來源' }));
    for (var i = 0; i < months; i++) {
      var r = work.refs && work.refs[i];
      head.appendChild(h('th', { class: 'colhead', scope: 'col',
        text: XIO.numToCol(XIO.REPORT.monthCol + i) + (r && r.text ? '　' + r.text : '') }));
    }
    var body = h('tbody');
    XIO.REPORT_LINKS.concat(MP_LINKS).forEach(function (link) {
      var info = work.links[link.key];
      var src;
      var isMp = !!info.sheet;
      if (!info.found && isMp) {
        warn = true;
        src = '找不到';
        els.notes.appendChild(h('li', { class: 'is-warn' }, '現有人力 ' + rowLabel[link.key] + '：' + info.reason + '，這一列先留空白。'));
      } else if (isMp) {
        src = info.sheet + ' 工作表「' + link.label + '」第 ' + info.row + ' 列，' +
          XIO.numToCol(info.refs[0].col) + ' 到 ' + XIO.numToCol(info.refs[info.refs.length - 1].col) + ' 行' +
          (info.from === 'saved' ? '（第 8 步上次確認的位置）' : '');
        var blank = info.refs.filter(function (x) { return x.value === null; }).length;
        if (blank) {
          warn = true;
          els.notes.appendChild(h('li', { class: 'is-warn' }, '現有人力 ' + rowLabel[link.key] + '：有 ' + blank + ' 個月是空白或不是數字。'));
        }
      } else if (!info.found) {
        warn = true;
        src = '找不到「' + link.label + '」';
        els.notes.appendChild(h('li', { class: 'is-warn' }, rowLabel[link.key] + '：' + work.main.name +
          ' 工作表裡找不到「' + link.label + '」這一列（第 ' + link.step + ' 步做的），這一列先留空白。'));
      } else {
        var first = info.refs[0], last = info.refs[info.refs.length - 1];
        src = '「' + link.label + '」第 ' + info.row + ' 列' +
          (first ? '，' + XIO.numToCol(first.col) + ' 到 ' + XIO.numToCol(last.col) + ' 行' : '');
        if (info.refs.length < months) {
          warn = true;
          els.notes.appendChild(h('li', { class: 'is-warn' }, rowLabel[link.key] + '：「' + link.label + '」那一列只找到 ' +
            info.refs.length + ' 個月的數值，其餘留空白。'));
        }
      }
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: (isMp ? '現有人力 ' : '') + rowLabel[link.key] }),
        h('td', { text: src }));
      for (var m = 0; m < months; m++) {
        var ref = info.refs[m];
        var shown = !ref ? '' : ref.value === null ? (isMp ? '（空白）' : '（重新計算後才有）')
          : isMp ? String(Math.round(ref.value * 10) / 10) : Math.round(ref.value).toLocaleString('en-US');
        tr.appendChild(h('td', {
          class: ref && !(isMp && ref.value === null) ? 'is-num' : 'is-missing',
          text: shown
        }));
      }
      body.appendChild(tr);
    });

    var ex = work.links.section.found && work.links.section.refs[0];
    els.result.textContent = '';
    els.result.appendChild(h('p', { class: 'sheet-caption', text: '從其他工作表抓的列（需求工時只顯示整數）：' }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '總表已填的列' },
      h('table', { class: 'sheet is-result is-check' }, h('thead', null, head), body)));
    if (ex) {
      els.result.appendChild(h('p', { class: 'hint', text: '每格都是公式，例如總表 D' + XIO.REPORT_ROWS[0].row + ' ＝ =\'' +
        ex.sheet + "'!" + XIO.numToCol(ex.col) + ex.row + '。' }));
    }
    return warn;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, refs, links, inputs) {
    return buildFrom(src).then(function (wb) {
      XIO.addReportSheet(wb, { monthRefs: refs, links: links, inputs: inputs, yoLimit: inputs.yoLimit });
      return wb;
    });
  }

  function outputName(src) {
    var stem = src.name.replace(/\.[^.]+$/, '');
    if (stem.slice(-SUFFIX.length) === SUFFIX) stem = stem.slice(0, -SUFFIX.length);
    return stem + SUFFIX + '.xlsx';
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.withReport;
    if (!ready) return;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    ready.build()
      .then(function (wb) { return XIO.toBlob(wb); })
      .then(function (blob) {
        ui.download(blob, ready.outName);
        els.exportNote.textContent = '已匯出 ' + ready.outName + '，在瀏覽器的下載資料夾。';
      })
      .catch(function (err) {
        els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
      })
      .then(function () {
        els.exportBtn.disabled = false;
      });
  }

  IEApp.registerStep({ id: 'report', render: render });
})(window);
