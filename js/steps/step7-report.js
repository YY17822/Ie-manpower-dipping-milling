/*
 * 第 7 步：總表（Dipping_Milling）
 *   在活頁簿裡新增「總表」工作表。目前先建立表格的格式、文字、框線；各格的內容（公式）之後再一項一項補上。
 *
 *   版面（B 到 K 行、第 2 到 34 列，設定在 js/excel-io.js 的 DM_REPORT、DM_ROWS）：
 *     第 2 列      表頭：B:C 合併畫斜線；D 是 Unit；E 到 J 是六個月；K 是 Notes（深藍底白字）
 *     第 3–8 列    工作天數、工作時間、機台每日可動工時[C]、工作班別數、出勤率、作業效率
 *     第 9–10 列   MPS：Milling、Dipping（淺藍底）
 *     第 11–16 列  Milling站：機台需求工時、人機比、開機數、顧機人數；加班天數-銑切；需求顧機人數-銑切（淺綠底）
 *     第 17–22 列  Dipping站：同上；加班天數-Dipping；需求顧機人數-dipping（淺紫底）
 *     第 23–28 列  需求時數-目檢：Milling、Dipping（min/月）、Milling、Dipping（人/班）；加班天數-目檢；需求顧機人數-目檢（淺褐底）
 *     第 29 列     總需求人力（淺橘底）
 *     第 30–31 列  現有人力：Milling、Dipping
 *     第 32 列     總現有人力（淺橘底，下面粗線）
 *     第 33 列     人力需求差異
 *     第 34 列     Loading Rate (9)=(X)/(1)（淺橘底）
 *
 *   月份表頭用公式參照 MPS 工作表「機台需求工時」那一塊的月份（第 5 步輸入的），例如 ='MPS'!P3。
 *
 *   來源有兩種：
 *     a. 剛才第 6 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 6 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withReport，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STEP1_KEY = IEApp.storeKey('step1');
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
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：RMA人力需求',
      onclick: function () { IEApp.go('rma-manpower'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '總表' }),
      h('p', { text: '在同一個檔案裡新增「總表」工作表，建立表格的格式、文字和框線：MPS、Milling站、Dipping站、需求時數-目檢、總需求人力、現有人力、人力需求差異、Loading Rate。月份表頭會參照 MPS 工作表「機台需求工時」的月份。各格的內容之後再補上。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 6 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.layout,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

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
    els.layout.textContent = '';

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
    els.notes.appendChild(h('li', null, '目前只建立格式、文字和框線，各格的數值之後再補上。'));
    renderLayout(refs);

    els.exportBtn.disabled = false;
    IEApp.state.withReport = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: dataStartRow(src),
      reportSheet: SHEET_NAME,
      reportRows: rowKeys(),
      monthCol: R.monthCol,
      monthCount: R.months,
      build: function () { return buildWorkbook(src, refs); }
    };
    els.nextBtn.disabled = false;
  }

  function rowKeys() {
    var out = {};
    XIO.DM_ROWS.forEach(function (x) { out[x.key] = x.r; });
    return out;
  }

  // 畫面上的版面預覽：各列的項目和單位
  function renderLayout(refs) {
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: '項目' }),
      h('th', { class: 'colhead', scope: 'col', text: 'Unit' }));
    var months = [];
    for (var i = 0; i < R.months; i++) months.push(refs && refs[i] ? refs[i].text : XIO.numToCol(R.monthCol + i));
    head.appendChild(h('th', { class: 'colhead', scope: 'col', text: months.join(' / ') }));
    var groups = { 9: 'MPS', 10: 'MPS', 11: 'Milling站', 12: 'Milling站', 13: 'Milling站', 14: 'Milling站',
      17: 'Dipping站', 18: 'Dipping站', 19: 'Dipping站', 20: 'Dipping站',
      23: '需求時數-目檢', 24: '需求時數-目檢', 25: '需求時數-目檢', 26: '需求時數-目檢', 30: '現有人力', 31: '現有人力' };
    var units = { 9: 'PCS', 10: 'PCS', 23: 'min/月', 24: 'min/月', 25: '人/班', 26: '人/班', 30: '人/天', 31: '人/天' };
    var body = h('tbody');
    XIO.DM_ROWS.forEach(function (x) {
      var label = x.b || ((groups[x.r] ? groups[x.r] + '｜' : '') + x.c);
      body.appendChild(h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(x.r) }),
        h('td', { text: label }),
        h('td', { text: x.u || units[x.r] || '' }),
        h('td', { text: '' })));
    });
    els.layout.appendChild(h('p', { class: 'sheet-caption', text: '總表的版面。' }));
    els.layout.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '總表的版面' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, refs) {
    return buildFrom(src).then(function (wb) {
      XIO.addDmReportSheet(wb, { monthRefs: refs });
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
