/*
 * 第 7 步：總表（Dipping_Milling）
 *   在活頁簿裡新增「總表」工作表，建立表格的格式、文字、框線，並填入使用者在畫面上輸入的數值：
 *     工作天數           各月輸入（不會記住）
 *     工作時間           輸入工作時數、休息時數（小時，六個月共用），填 =(工作時數-休息時數)*60（分鐘）
 *     機台每日可動工時   各月輸入小時，填 =小時*60（分鐘）
 *     機台每日可動工時、工作班別數、出勤率、作業效率   各月輸入，有「套用到全部月份」按鈕（第一個月的值填到後面每個月）
 *     MPS Milling、Dipping   公式：參照 MPS 工作表該段（Milling = Cutting）第 2 步的 Total 列，例如 ='MPS'!E10；
 *                            沒有 Total 列就直接加總該段（當月也不加 WIP）
 *     Milling站
 *       機台需求工時   公式：參照 MPS 工作表 Milling 段「機台需求工時」第 5 步的 Total 列，例如 ='MPS'!P10
 *       人機比         輸入機台數、人員數（六個月共用），填 =機台數/人員數；Notes 寫「1 人顧 4 機」
 *       開機數         =MIN(ROUNDUP(機台需求工時/((工作天數+加班天數 - Milling)*機台每日可動工時),1),8)
 *                      最多 8 台（Dipping站最多 4 台），算出來超過就顯示上限
 *       顧機人數       =ROUNDUP(開機數/人機比,1)
 *       需求顧機人數   =ROUNDUP(顧機人數*工作班別數/出勤率/作業效率,0)
 *     Dipping站     同上（參照 Dipping 段的 Total 列、自己的機台數和人員數、加班天數 - Dipping）
 *     目檢站 需求工時 - Milling、- Dipping   公式：參照 MPS 工作表該段「目檢需求工時」第 5 步的 Total 列
 *     目檢站 需求人力 - Milling、- Dipping
 *       =ROUNDUP(需求工時/工作時間/工作班別數/(工作天數+加班天數 - 目檢)/出勤率/作業效率,1)
 *     需求人數 - 目檢   =ROUNDUP(需求人力 - Milling*工作班別數+需求人力 - Dipping*工作班別數,1)
 *     總需求人力       =需求顧機人數 - Milling+需求顧機人數 - dipping+需求人數 - 目檢
 *     總現有人力       =現有人力 Milling+Dipping（現有人力兩列之後再補上）
 *     人力需求差異     =總現有人力-總需求人力
 *   其他格的內容（公式）之後再一項一項補上。
 *
 *   版面（B 到 K 行、第 2 到 34 列，設定在 js/excel-io.js 的 DM_REPORT、DM_ROWS）：
 *     第 2 列      表頭：B:C 合併畫斜線；D 是 Unit；E 到 J 是六個月；K 是 Notes（深藍底白字）
 *     第 3–8 列    工作天數、工作時間、機台每日可動工時、工作班別數、出勤率、作業效率
 *     第 9–10 列   MPS：Milling、Dipping（淺藍底）
 *     第 11–16 列  Milling站：機台需求工時、人機比、開機數、顧機人數；加班天數 - Milling；需求顧機人數 - Milling（淺綠底）
 *     第 17–22 列  Dipping站：同上；加班天數 - Dipping；需求顧機人數 - dipping（淺紫底）
 *     第 23–28 列  目檢站：需求工時 - Milling、需求工時 - Dipping（min/月）、需求人力 - Milling、需求人力 - Dipping（人/班）；
 *                  加班天數 - 目檢；需求人數 - 目檢（淺褐底）
 *     第 29 列     總需求人力（淺橘底）
 *     第 30–31 列  現有人力：Milling、Dipping
 *     第 32 列     總現有人力（淺橘底，下面粗線）
 *     第 33 列     人力需求差異
 *     第 34 列     Loading Rate（淺橘底）
 *
 *   月份表頭用公式參照 MPS 工作表「機台需求工時」那一塊的月份（第 5 步輸入的），例如 ='MPS'!P3。
 *
 *   來源有兩種：
 *     a. 剛才第 6 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 6 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withReport（這是最後一步，之後要接新的步驟時可以直接拿來用）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STEP1_KEY = 'step1';
  var R = XIO.DM_REPORT;
  var SHEET_NAME = R.sheet;
  var MONTH_TITLE = XIO.DEMAND_TITLES.machine;      // 月份從「機台需求工時」那一塊抓
  var SUFFIX = '_總表';
  var EXTRA = ['P003', '人力', SHEET_NAME, 'RMA人力需求'];

  var S = {
    upload: null,         // 上傳的第 6 步檔案 { name, buffer, sheets }
    preferUpload: false,
    work: null            // { src, loading, error, wb, main, refs }
  };
  var els = {};

  /* ---------- 使用者輸入的數值 ---------- */
  // 每個月一格的列。percent：輸入的是 %（95 代表 95%）；copy：後面有「套用到全部月份」按鈕；remember：記在瀏覽器裡
  var GRID = [
    { key: 'workDays', label: '工作天數', unit: '天', remember: false },
    { key: 'machineHours', label: '機台每日可動工時', unit: '小時', remember: true, copy: true },
    { key: 'shifts', label: '工作班別數', unit: '班', remember: true, copy: true },
    { key: 'attendance', label: '出勤率', unit: '%', remember: true, copy: true, percent: true },
    { key: 'efficiency', label: '作業效率', unit: '%', remember: true, copy: true, percent: true }
  ];
  // 人機比：每一站輸入一組機台數、人員數（六個月共用）
  var STATIONS = [
    { key: 'cutting', label: 'Milling站', machines: 'millMachines', people: 'millPeople', id: 'mill' },
    { key: 'dipping', label: 'Dipping站', machines: 'dipMachines', people: 'dipPeople', id: 'dip' }
  ];
  var INPUT_KEY = 'step7';             // 共用設定的組名（IEApp.settings）
  var INP = loadInputs();       // { workDays: ['', …], machineHours: [...], …, workHours: '', restHours: '' }（輸入框的文字）

  function blankMonths() { var a = []; for (var i = 0; i < R.months; i++) a.push(''); return a; }
  function loadInputs() {
    var s = {};
    try { s = IEApp.settings.read(INPUT_KEY) || {}; } catch (e) { s = {}; }
    var out = { workHours: s.workHours || '', restHours: s.restHours || '' };
    STATIONS.forEach(function (st) { out[st.machines] = s[st.machines] || ''; out[st.people] = s[st.people] || ''; });
    GRID.forEach(function (g) {
      var saved = g.remember && Array.isArray(s[g.key]) ? s[g.key] : [];
      out[g.key] = blankMonths().map(function (v, i) { return saved[i] === undefined ? '' : String(saved[i]); });
    });
    return out;
  }
  function saveInputs() {
    var s = { workHours: INP.workHours, restHours: INP.restHours };
    STATIONS.forEach(function (st) { s[st.machines] = INP[st.machines]; s[st.people] = INP[st.people]; });
    GRID.forEach(function (g) { if (g.remember) s[g.key] = INP[g.key]; });
    try { IEApp.settings.write(INPUT_KEY, s); } catch (e) { /* 存不了就算了 */ }
  }

  // 輸入框的文字 → 數字。空白回傳 null；看不懂回傳 NaN。percent 時 95 或 95% 都回傳 0.95
  function toValue(text, percent) {
    var t = XIO.normalizeText(text || '').replace(/,/g, '');
    if (percent) t = t.replace(/[%％]\s*$/, '');
    t = t.trim();
    if (t === '') return null;
    var n = Number(t);
    if (isNaN(n) || n < 0) return NaN;
    return percent ? n / 100 : n;
  }

  // 讀出所有輸入：{ values: 給 addDmReportSheet 的 inputs, bad: [看不懂的], blank: [沒填的] }
  function readInputs() {
    var values = {}, bad = [], blank = [];
    var names = monthNames();
    GRID.forEach(function (g) {
      values[g.key] = INP[g.key].map(function (text, i) {
        var v = toValue(text, g.percent);
        if (v === null) blank.push(g.label + '（' + names[i] + '）');
        else if (isNaN(v)) bad.push(g.label + '（' + names[i] + '）「' + text + '」');
        return v === null || isNaN(v) ? null : v;
      });
    });
    [['workHours', '工作時數'], ['restHours', '休息時數']].forEach(function (p) {
      var v = toValue(INP[p[0]], false);
      if (v === null) blank.push(p[1]);
      else if (isNaN(v)) bad.push(p[1] + '「' + INP[p[0]] + '」');
      values[p[0]] = v === null || isNaN(v) ? null : v;
    });
    STATIONS.forEach(function (st) {
      [[st.machines, st.label + '的機台數'], [st.people, st.label + '的人員數']].forEach(function (p) {
        var v = toValue(INP[p[0]], false);
        if (v === null) blank.push(p[1]);
        else if (isNaN(v)) bad.push(p[1] + '「' + INP[p[0]] + '」');
        values[p[0]] = v === null || isNaN(v) ? null : v;
      });
      if (values[st.people] === 0) bad.push(st.label + '的人員數不能是 0');
    });
    if (values.workHours !== null && values.restHours !== null && values.restHours > values.workHours) {
      bad.push('休息時數比工作時數還多');
    }
    return { values: values, bad: bad, blank: blank };
  }

  function monthNames() {
    var refs = S.work && S.work.refs;
    var out = [];
    for (var i = 0; i < R.months; i++) out.push(refs && refs[i] && refs[i].text ? refs[i].text : '第 ' + (i + 1) + ' 個月');
    return out;
  }

  // 輸入的表格：一列一個項目、一行一個月份；建立一次，之後只更新月份名稱
  function buildInputGrid() {
    els.monthHeads = [];
    els.gridInputs = {};
    var head = h('tr', null, h('th', { scope: 'col', text: '項目' }));
    for (var i = 0; i < R.months; i++) {
      var th = h('th', { scope: 'col', class: 'is-month', text: '第 ' + (i + 1) + ' 個月' });
      els.monthHeads.push(th);
      head.appendChild(th);
    }
    head.appendChild(h('th', { scope: 'col' }));
    var body = h('tbody');
    GRID.forEach(function (g) {
      els.gridInputs[g.key] = [];
      var tr = h('tr', null, h('th', { scope: 'row' },
        h('span', { class: 'grid-label', text: g.label }), h('span', { class: 'grid-unit', text: g.unit })));
      for (var m = 0; m < R.months; m++) {
        var input = h('input', {
          type: 'text', class: 'in in-grid', inputmode: 'decimal', autocomplete: 'off',
          value: INP[g.key][m], 'aria-label': g.label + ' 第 ' + (m + 1) + ' 個月'
        });
        (function (key, idx, el) {
          el.addEventListener('input', function () { INP[key][idx] = el.value; onType(); });
        })(g.key, m, input);
        els.gridInputs[g.key].push(input);
        tr.appendChild(h('td', null, input));
      }
      tr.appendChild(h('td', null, g.copy ? h('button', {
        type: 'button', class: 'btn-secondary btn-small', text: '套用到全部月份',
        title: '把第一個月的值填到後面每個月',
        onclick: function () { copyFirst(g.key); }
      }) : (g.note ? h('span', { class: 'grid-note', text: g.note }) : null)));
      body.appendChild(tr);
    });
    return h('div', { class: 'sheet-scroll grid-scroll' }, h('table', { class: 'grid-inputs' }, h('thead', null, head), body));
  }

  function copyFirst(key) {
    var first = INP[key][0];
    for (var m = 1; m < R.months; m++) {
      INP[key][m] = first;
      els.gridInputs[key][m].value = first;
    }
    onType();
  }

  var typeTimer = null;
  function onType() {
    saveInputs();
    els.exportNote.textContent = '';
    clearTimeout(typeTimer);
    typeTimer = setTimeout(refresh, 250);
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
      var s1 = IEApp.settings.read(STEP1_KEY) || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  // 各月產量在哪幾行：和第 2、5 步共用的設定（沒有設定就用預設的第 5 到 10 行）
  function mpsColumns() {
    var shared = ui.columns.get();
    var months = XIO.parseColumns(shared.monthCols || '').cols;
    if (!months.length) months = [5, 6, 7, 8, 9, 10];
    return { monthCols: months };
  }

  // 總表 MPS 那兩列的公式（Milling = Cutting 那一段、Dipping）
  function mpsFormulas(main, src) {
    var cols = mpsColumns();
    return XIO.dmMpsFormulas(main, { headerEnd: dataStartRow(src) - 1, monthCols: cols.monthCols });
  }

  // Milling站、Dipping站 機台需求工時的公式（參照 MPS 工作表機台需求工時那一塊各段的 Total 列）
  function machineTotals(main, src) {
    return XIO.dmBlockTotalFormulas(main, { headerEnd: dataStartRow(src) - 1, title: MONTH_TITLE, months: R.months });
  }

  // 目檢站 需求工時的公式（參照 MPS 工作表目檢需求工時那一塊各段的 Total 列）
  var LABOR_TITLE = XIO.DEMAND_TITLES.labor;
  function laborTotals(main, src) {
    return XIO.dmBlockTotalFormulas(main, { headerEnd: dataStartRow(src) - 1, title: LABOR_TITLE, months: R.months });
  }

  // 月份表頭要參照的儲存格：MPS 工作表「機台需求工時」那一塊的月份（大表頭下一列）
  function monthRefs(main, src) {
    var headerEnd = dataStartRow(src) - 1;
    var block = XIO.findHeaderColumn(main, headerEnd, MONTH_TITLE);
    if (!block) return null;
    var titleRow = 0;
    for (var r = 1; r <= headerEnd && !titleRow; r++) {
      var cell = main.findCell(r, block);
      if (cell && XIO.lookupKey(cell.value) === XIO.lookupKey(MONTH_TITLE)) titleRow = r;
    }
    var headRow = titleRow ? titleRow + 1 : (headerEnd >= 3 ? 3 : headerEnd);
    if (headRow < 1 || headRow > headerEnd) return null;
    var refs = [];
    for (var i = 0; i < R.months; i++) {
      var c = block + i;
      var hc = main.findCell(headRow, c);
      refs.push({ sheet: main.name, col: c, row: headRow, text: hc ? XIO.displayText(hc) : '' });
    }
    return refs;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f9-base',
      title: '上傳第 6 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 6 步匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.layout = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    // 工作時間：工作時數、休息時數（六個月共用）
    els.workHours = h('input', { id: 'f9-work-h', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP.workHours, placeholder: '例如 12' });
    els.restHours = h('input', { id: 'f9-rest-h', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP.restHours, placeholder: '例如 1' });
    els.workMinutes = h('span', { class: 'hint-inline' });
    els.workHours.addEventListener('input', function () { INP.workHours = els.workHours.value; onType(); });
    // 人機比：每一站一列（機台數、人員數）
    els.ratioText = {};
    var ratioRows = STATIONS.map(function (st) {
      var mIn = h('input', { id: 'f9-' + st.id + '-m', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP[st.machines], placeholder: '例如 4' });
      var pIn = h('input', { id: 'f9-' + st.id + '-p', type: 'text', class: 'in in-tiny', inputmode: 'decimal', autocomplete: 'off', value: INP[st.people], placeholder: '例如 1' });
      mIn.addEventListener('input', function () { INP[st.machines] = mIn.value; onType(); });
      pIn.addEventListener('input', function () { INP[st.people] = pIn.value; onType(); });
      els.ratioText[st.key] = h('span', { class: 'hint-inline' });
      return h('div', { class: 'field-row' },
        h('span', { class: 'field-label label-fixed', text: st.label }),
        h('label', { for: 'f9-' + st.id + '-m', text: '機台數' }), mIn, h('span', { text: '台' }),
        h('label', { for: 'f9-' + st.id + '-p', text: '人員數' }), pIn, h('span', { text: '人' }),
        els.ratioText[st.key]);
    });
    els.restHours.addEventListener('input', function () { INP.restHours = els.restHours.value; onType(); });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '總表' }),
      h('p', { text: '在同一個檔案裡新增「總表」工作表：MPS、Milling站、Dipping站、目檢站、總需求人力、現有人力、人力需求差異、Loading Rate。月份表頭會參照 MPS 工作表「機台需求工時」的月份。這一步先填工作天數、工作時間、機台每日可動工時、工作班別數、出勤率、作業效率、MPS 的兩列，以及 Milling站、Dipping站的機台需求工時、人機比、開機數、顧機人數；其他格之後再補上。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 6 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '工作時間' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f9-work-h', text: '工作時數' }), els.workHours, h('span', { text: '小時' }),
          h('label', { for: 'f9-rest-h', text: '休息時數' }), els.restHours, h('span', { text: '小時' }),
          els.workMinutes),
        h('p', { class: 'hint', text: '六個月共用。總表的工作時間會填公式 =(工作時數-休息時數)*60，單位是分鐘。會記住上次輸入的值。' })),
      h('li', { class: 'task' },
        h('h2', { text: '各月的數值' }),
        buildInputGrid(),
        h('p', { class: 'hint', text: '出勤率、作業效率輸入 95 或 95% 都代表 95%。機台每日可動工時、工作班別數、出勤率、作業效率只要填第一個月，再按「套用到全部月份」。機台每日可動工時輸入小時，總表會換算成分鐘（例如 =22*60）。工作天數每個月不同，不會記住；其他會記住上次輸入的值。沒填的格子在總表留空白。' })),
      h('li', { class: 'task' },
        h('h2', { text: '人機比' }),
        ratioRows,
        h('p', { class: 'hint', text: '每一站一組，六個月共用。總表的人機比填公式 =機台數/人員數，Notes 寫「幾人顧幾機」。開機數 =ROUNDUP(機台需求工時/((工作天數+該站的加班天數)*機台每日可動工時),1)，超過上限時顯示上限（Milling站 8、Dipping站 4），顧機人數 =ROUNDUP(開機數/人機比,1)，都會自動填公式。會記住上次輸入的值。' })),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.layout,
        h('div', { class: 'actions' }, els.exportBtn, els.exportNote))));

    refresh();
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(function (n) { return !isExtra(n); })) {
          throw new Error('這個檔案沒有 MPS 工作表。請上傳第 6 步匯出的檔案。');
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

  function ensureWork(src) {
    if (S.work && S.work.src === src.key) return;
    var work = { src: src.key, loading: true };
    S.work = work;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;
      var main = pickMain(wb, src);
      if (!main) throw new Error('找不到 MPS 工作表。');
      work.wb = wb;
      work.main = main;
      work.refs = monthRefs(main, src);
      work.mps = mpsFormulas(main, src);
      work.machineTotals = machineTotals(main, src);
      work.laborTotals = laborTotals(main, src);
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
    els.notes.textContent = '';
    els.layout.textContent = '';
    refreshInputLabels();

    if (!src) {
      els.srcInfo.textContent = '還沒有第 6 步的結果。回到第 6 步，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 6 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 6 步剛做好的結果' : '改用上傳的檔案';

    if (!src) { setSummary('先準備好第 6 步的 Excel。'); return; }
    ensureWork(src);
    var work = S.work;
    if (work.loading) { setSummary('準備資料中…'); return; }
    if (work.error) { setSummary(work.error, true); return; }

    var refs = work.refs;
    var first = XIO.numToCol(R.monthCol), last = XIO.numToCol(R.monthCol + R.months - 1);
    setSummary('會新增「' + SHEET_NAME + '」工作表（B 到 K 行、第 ' + R.headRow + ' 到 ' + R.lastRow + ' 列）。' +
      (refs ? '月份表頭 ' + first + R.headRow + ' 到 ' + last + R.headRow + ' 參照 ' + work.main.name + ' 工作表「' + MONTH_TITLE + '」的月份：' +
        refs.map(function (x) { return x.text || '（空白）'; }).join('、') + '。' : ''), false, !refs);
    if (!refs) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        work.main.name + ' 工作表找不到「' + MONTH_TITLE + '」（第 5 步），月份表頭先留空白。'));
    }
    var mps = work.mps;
    var mpsText = [['cutting', 'Milling'], ['dipping', 'Dipping']].map(function (p) {
      var list = mps && mps[p[0]];
      if (!list) return p[1] + '：找不到';
      return p[1] + '：' + (list[0].fromTotal ? '參照第 2 步的 Total 列' : '直接加總該段') + '，例如 =' + list[0].formula;
    });
    els.notes.appendChild(h('li', { class: mps && mps.cutting && mps.dipping ? null : 'is-warn' },
      'MPS 的兩列（公式）：' + mpsText.join('；') + '。' +
      (mps && mps.cutting && mps.dipping ? '' : '找不到的那一列先留空白（MPS 工作表裡要有第 1 步的「Cutting（…）」「Dipping（…）」標題和料號）。')));

    var mt = work.machineTotals;
    STATIONS.forEach(function (st) {
      var list = mt && mt[st.key];
      if (list) {
        els.notes.appendChild(h('li', null, st.label + ' 機台需求工時（公式）：' +
          (list[0].fromTotal ? '參照第 5 步的 Total 列' : '直接加總該段') + '，例如 =' + list[0].formula + '。'));
      } else {
        els.notes.appendChild(h('li', { class: 'is-warn' },
          'MPS 工作表找不到' + st.label.replace('站', '') + '段的「' + MONTH_TITLE + '」（第 5 步），' + st.label + '的機台需求工時先留空白。'));
      }
    });

    var lt = work.laborTotals;
    STATIONS.forEach(function (st) {
      var list = lt && lt[st.key];
      var name = '目檢站 需求工時 - ' + st.label.replace('站', '');
      if (list) {
        els.notes.appendChild(h('li', null, name + '（公式）：' +
          (list[0].fromTotal ? '參照第 5 步的 Total 列' : '直接加總該段') + '，例如 =' + list[0].formula + '。'));
      } else {
        els.notes.appendChild(h('li', { class: 'is-warn' },
          'MPS 工作表找不到' + st.label.replace('站', '') + '段的「' + LABOR_TITLE + '」（第 5 步），' + name + '先留空白。'));
      }
    });

    var read = readInputs();
    if (read.bad.length) {
      setSummary('看不懂：' + read.bad.join('、') + '。請輸入數字。', true);
      return;
    }
    if (read.blank.length) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '還沒填：' + read.blank.slice(0, 12).join('、') + (read.blank.length > 12 ? ' …等 ' + read.blank.length + ' 格' : '') +
        '。沒填的格子在總表留空白，一樣可以匯出。'));
    }
    els.notes.appendChild(h('li', null, '目前填入工作天數、工作時間、機台每日可動工時、工作班別數、出勤率、作業效率、MPS 的兩列、Milling站和 Dipping站（機台需求工時、人機比、開機數、顧機人數、需求顧機人數）、目檢站（需求工時、需求人力）、需求人數 - 目檢、總需求人力、總現有人力、人力需求差異；現有人力、加班天數、Loading Rate 之後再補上。'));

    var inputs = read.values;
    els.exportBtn.disabled = false;
    IEApp.state.withReport = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: dataStartRow(src),
      reportSheet: SHEET_NAME,
      reportRows: rowKeys(),
      monthCol: R.monthCol,
      monthCount: R.months,
      inputs: inputs,
      build: function () { return buildWorkbook(src, refs, inputs); }
    };
  }

  // 月份名稱（從 MPS 抓到的）、工作時間換算後的分鐘，顯示在輸入區
  function refreshInputLabels() {
    var names = monthNames();
    (els.monthHeads || []).forEach(function (th, i) { th.textContent = names[i]; });
    var w = toValue(INP.workHours, false), r = toValue(INP.restHours, false);
    els.workMinutes.textContent = w !== null && r !== null && !isNaN(w) && !isNaN(r)
      ? '＝ 每班 ' + Math.round((w - r) * 60 * 100) / 100 + ' 分鐘' : '';
    STATIONS.forEach(function (st) {
      var mm = toValue(INP[st.machines], false), pp = toValue(INP[st.people], false);
      els.ratioText[st.key].textContent = mm !== null && pp !== null && !isNaN(mm) && !isNaN(pp) && pp > 0
        ? '＝ 人機比 ' + Math.round(mm / pp * 100) / 100 + '（' + pp + ' 人顧 ' + mm + ' 機）' : '';
    });
  }

  function rowKeys() {
    var out = {};
    XIO.DM_ROWS.forEach(function (x) { out[x.key] = x.r; });
    return out;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, refs, inputs) {
    return buildFrom(src).then(function (wb) {
      var main = pickMain(wb, src);
      var mt = main ? machineTotals(main, src) : null;
      var lt = main ? laborTotals(main, src) : null;
      XIO.addDmReportSheet(wb, {
        monthRefs: refs, inputs: inputs, mps: main ? mpsFormulas(main, src) : null,
        millingHours: mt ? mt.cutting : null,
        millingRatio: { machines: inputs.millMachines, people: inputs.millPeople },
        dippingHours: mt ? mt.dipping : null,
        dippingRatio: { machines: inputs.dipMachines, people: inputs.dipPeople },
        laborHours: lt || null
      });
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
