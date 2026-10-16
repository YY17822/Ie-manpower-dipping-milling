/*
 * 第 5 步：計算需求工時
 *   把每個料號各月的計畫產量乘上該列的 CT.，算出六個月各自的需求工時。
 *   當月比較特別：要先加上當月 WIP 再乘 CT.。
 *
 *     當月   =(WIP + 當月數量) * CT.      例如 =(D5+E5)*L5
 *     其餘月 = 數量 * CT.                 例如 =F5*L5
 *
 *   結果從 MPS 工作表「原本最後一行之後空一行」開始填，每格都是公式。
 *   表頭：第 2 列是跨所有月份置中的「需求工時」，第 3 列到表頭最後一列是各月的月份。
 *        月份由使用者自己輸入，程式不會自己去 MPS 的表頭找。
 *   數值只顯示整數（儲存格裡仍保留小數，加總時用的是完整的數值）。
 *
 *   預設的行（照第 1 步保留的順序）：第 4 行是 WIP，第 5 到第 10 行是當月和往後五個月。
 *
 *   來源有兩種：
 *     a. 剛才第 4 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 4 步匯出的檔案（MPS 工作表裡要有 CT. 這一行）
 *
 * 做完之後把結果放在 IEApp.state.withDemand，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = IEApp.storeKey('step4');       // 鍵名是舊的步驟編號，改掉會讓記住的設定不見，所以沿用
  var STEP1_KEY = IEApp.storeKey('step1');
  var REF_SHEET = 'P003';
  var CT_HEADER = 'CT.';
  var TITLE = '需求工時';
  var SUFFIX = '_需求工時';
  var WIP_DEFAULT = 4;          // 第 4 行是當月 WIP
  var MONTH_FIRST = 5;          // 第 5 行起是月份
  var MONTH_COUNT = 6;          // 當月 + 往後五個月

  var DEFAULTS = { wipCol: '', monthCols: '', startRow: '' };

  // 這一步的狀態（切到別的步驟再回來時還在）
  var S = {
    upload: null,         // 上傳的第 4 步檔案 { name, buffer }
    preferUpload: false,
    work: null,           // 拿來預覽、統計的活頁簿 { src, loading, error, wb, main, mainExt }
    headerEdits: {}       // 使用者輸入的月份表頭 { 行號: 文字 }
  };
  var settings = loadSettings();
  var els = {};
  var timer = null;

  // WIP、月份的行和第 2 步（匯總）共用同一份設定；資料起始列是這一步自己的。
  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var shared = ui.columns.get();
    return {
      wipCol: shared.wipCol,
      monthCols: shared.monthCols,
      startRow: s.startRow === undefined ? DEFAULTS.startRow : s.startRow
    };
  }
  function saveSettings() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify({ startRow: settings.startRow })); } catch (e) { /* 存不了就算了 */ }
    ui.columns.set({ wipCol: settings.wipCol, monthCols: settings.monthCols });
  }

  function isRefName(name) { return String(name).toLowerCase() === REF_SHEET.toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.withCT;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  function buildFrom(src) {
    return src.kind === 'prev' ? src.prev.build() : XIO.load(src.key.buffer.slice(0));
  }

  function pickMain(wb, src) {
    var others = wb.worksheets.filter(function (w) { return !isRefName(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    return others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
  }

  // MPS 的資料從第幾列開始：接著前面的步驟做下來的話沿用第 1 步的設定
  function chainStartRow() {
    var src = activeSource();
    return src && src.kind === 'prev' && src.prev.dataStartRow ? src.prev.dataStartRow : 0;
  }
  function startRow() {
    var chain = chainStartRow();
    if (chain) return chain;
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    if (n >= 1) return n;
    try {
      var s1 = JSON.parse(localStorage.getItem(STEP1_KEY) || '{}') || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }

  function colOf(letter, max) {
    var p = XIO.parseColumns(letter || '');
    return p.cols.length === 1 && !p.bad.length && p.cols[0] <= max ? p.cols[0] : 0;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};
    settings = loadSettings();        // 每次進來都重讀：第 2 步可能改過 WIP、月份的行

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f4-base',
      title: '上傳第 4 步匯出的 Excel',
      sub: '或把檔案拖進這裡。MPS 工作表裡要有 CT. 這一行',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 4 步帶入 CT.',
      onclick: function () { IEApp.go('ct-lookup'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.wipCol = h('select', { id: 'f4-wip', class: 'in-select' });
    els.monthCols = h('input', {
      id: 'f4-months', type: 'text', class: 'in in-select', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 E:J'
    });
    els.ctInfo = h('span', { class: 'field-static' });
    els.mainPreview = h('div', { class: 'sheet-area' });
    els.startRow = h('input', { id: 'f4-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric' });
    els.startNote = h('span');

    els.headers = h('div', { class: 'month-heads' });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：需求工時匯總',
      onclick: function () { IEApp.go('summary'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '計算需求工時' }),
      h('p', { text: '把每個料號各月的計畫產量乘上 CT.，算出各月的需求工時；當月會先加上 WIP 再乘。結果填在最後一行之後空一行的位置，每格都是公式，在 Excel 點儲存格就能看到是哪幾格算出來的。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 4 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: 'WIP 和各月產量在哪幾行' }),
        h('div', { class: 'field-grid' },
          h('label', { for: 'f4-wip', text: '當月 WIP 在' }), els.wipCol,
          h('label', { for: 'f4-months', text: '各月產量在' }), els.monthCols,
          h('span', { class: 'field-label', text: 'CT. 在' }), els.ctInfo),
        h('p', { class: 'hint', text: '各月產量的第一行是當月，會先加上 WIP 再乘 CT.；連續的行可以寫成 E:J，也可以點下面表格最上方的行號來選。這裡的設定和第 2 步（匯總）共用。' }),
        els.mainPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f4-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: '各月的表頭' }),
        h('p', { class: 'hint', text: '輸入每個月要顯示的表頭，例如 Oct、Nov。' }),
        els.headers),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.result,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.wipCol.addEventListener('change', function () {
      settings.wipCol = els.wipCol.value;
      saveSettings();
      els.exportNote.textContent = '';
      refresh();
    });
    els.monthCols.addEventListener('input', function () {
      settings.monthCols = els.monthCols.value;
      saveSettings();
      els.exportNote.textContent = '';
      clearTimeout(timer);
      timer = setTimeout(function () { refresh(true); }, 250);
    });
    els.startRow.addEventListener('change', function () {
      settings.startRow = els.startRow.value;
      saveSettings();
      refresh();
    });

    refresh();
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var others = res.workbook.worksheets.filter(function (w) { return !isRefName(w.name); });
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表。請上傳第 4 步匯出的檔案。');
        S.upload = { name: file.name, buffer: res.buffer };
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

  // 準備拿來預覽、統計的活頁簿。來源沒變就沿用，變了才重新產生。
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
      work.mainExt = XIO.extent(main);
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  function fillSelect(select, labels, colCount, chosen) {
    select.textContent = '';
    select.appendChild(h('option', { value: '', text: '請選擇' }));
    for (var c = 1; c <= colCount; c++) {
      var letter = XIO.numToCol(c);
      var label = labels[c] ? letter + '　' + labels[c] : letter;
      if (label.length > 42) label = label.slice(0, 41) + '…';
      select.appendChild(h('option', { value: letter, text: label }));
    }
    select.value = chosen ? XIO.numToCol(chosen) : '';
  }

  /* ---------- 更新畫面 ---------- */
  // keepMonthText：使用者正在打字時不要改寫輸入框的內容
  function refresh(keepMonthText) {
    var src = activeSource();
    var prev = IEApp.state.withCT;
    IEApp.state.withDemand = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 4 步的結果。回到第 4 步帶入 CT.，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 4 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 4 步剛做好的結果' : '改用上傳的檔案';

    var chain = chainStartRow();
    els.startRow.value = String(startRow());
    els.startRow.disabled = !!chain;
    els.startNote.textContent = chain ? '列開始（沿用第 1 步的設定）' : '列開始';

    if (!src) { clearChoices(); setSummary('先準備好第 4 步的 Excel。'); return; }

    ensureWork(src);
    var work = S.work;
    if (work.loading) { clearChoices(); setSummary('準備資料中…'); return; }
    if (work.error) { clearChoices(); setSummary(work.error, true); return; }

    var sr = startRow();
    var headerEnd = sr - 1;
    var cols = Math.max(work.mainExt.colCount, 1);
    var labels = XIO.columnLabels(work.main, headerEnd >= 2 ? 2 : 1, Math.max(headerEnd, 1), cols);

    // CT. 在哪一行：找表頭
    var ctCol = XIO.findHeaderColumn(work.main, headerEnd, CT_HEADER);
    els.ctInfo.textContent = ctCol ? XIO.numToCol(ctCol) + ' 行（表頭是「CT.」的那一行）' : '找不到';

    // WIP、月份：先用記住的，沒有就用預設的位置（第 4 行、第 5 到 10 行）
    var limit = ctCol ? ctCol - 1 : cols;        // 數量都在 CT. 左邊
    var wipCol = colOf(settings.wipCol, cols) || (WIP_DEFAULT <= limit ? WIP_DEFAULT : 0);
    var parsed = XIO.parseColumns(settings.monthCols);
    var monthCols = parsed.cols;
    if (!XIO.normalizeText(settings.monthCols)) {
      monthCols = [];
      for (var c = MONTH_FIRST; c < MONTH_FIRST + MONTH_COUNT && c <= limit; c++) monthCols.push(c);
    }

    fillSelect(els.wipCol, labels, cols, wipCol);
    if (!keepMonthText) els.monthCols.value = XIO.normalizeText(settings.monthCols) ? settings.monthCols : compact(monthCols);

    var keep = {};
    monthCols.forEach(function (mc) { keep[mc] = true; });
    ui.renderSheet(els.mainPreview, work.main, {
      headerEnd: headerEnd, lastRow: work.mainExt.lastRow, colCount: cols, rows: 5, keep: keep,
      onToggle: toggleMonth,
      caption: work.main.name + ' 工作表（前 5 列資料）。標起來的是各月產量。'
    });

    renderHeaderInputs(monthCols);

    if (!ctCol) {
      setSummary('MPS 工作表裡找不到表頭是「CT.」的那一行。請先做第 4 步帶入 CT.。', true);
      return;
    }
    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 E:J。', true);
      return;
    }
    if (!wipCol) { setSummary('選擇當月 WIP 在哪一行。'); return; }
    if (!monthCols.length) { setSummary('輸入各月產量在哪幾行。'); return; }
    var clash = monthCols.filter(function (mc) { return mc === wipCol || mc === ctCol || mc > cols; });
    if (clash.length) {
      setSummary(XIO.formatColumns(clash) + ' 行不能當作各月產量（和 WIP 或 CT. 重複，或超出報表範圍）。', true);
      return;
    }

    var headers = monthCols.map(function (mc) { return S.headerEdits[mc] || ''; });
    var emptyHead = monthCols.filter(function (mc, i) { return !XIO.normalizeText(headers[i]); });
    if (emptyHead.length) {
      setSummary('在上面「各月的表頭」輸入 ' + XIO.formatColumns(emptyHead) + ' 行的月份。');
      return;
    }

    var keyCol = src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;
    var opts = {
      dataStartRow: sr, keyCol: keyCol, wipCol: wipCol, monthCols: monthCols, ctCol: ctCol,
      title: TITLE, headers: headers
    };
    var plan = XIO.planDemand(work.main, opts);
    if (!plan.rows.length) {
      setSummary('第 ' + sr + ' 列之後沒有料號。請確認上傳的檔案。', true);
      return;
    }

    var firstL = XIO.numToCol(plan.startCol);
    var lastL = XIO.numToCol(plan.startCol + monthCols.length - 1);
    setSummary('需求工時會填在 ' + work.main.name + ' 工作表的 ' + firstL + ' 到 ' + lastL + ' 行' +
      (plan.isNew ? '（原本最後一行之後空一行）' : '（原本就有「需求工時」，會直接覆蓋）') +
      '，共 ' + plan.rows.length.toLocaleString('en-US') + ' 個料號、' + monthCols.length + ' 個月。',
      false, plan.noCt > 0 || plan.errors > 0);

    renderNotes(plan, monthCols);
    renderResult(plan, opts);
    els.exportBtn.disabled = false;

    // 給下一步用
    IEApp.state.withDemand = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: sr,
      keyCol: keyCol,
      demandCol: plan.startCol,
      monthCount: monthCols.length,
      build: function () { return buildWorkbook(src, opts); }
    };
    els.nextBtn.disabled = false;
  }

  // 連續的行寫成 E:J，其餘用逗號分開
  function compact(cols) {
    if (!cols.length) return '';
    var contiguous = cols.every(function (c, i) { return i === 0 || c === cols[i - 1] + 1; });
    if (contiguous && cols.length > 1) return XIO.numToCol(cols[0]) + ':' + XIO.numToCol(cols[cols.length - 1]);
    return XIO.formatColumns(cols);
  }

  function toggleMonth(c) {
    var text = XIO.normalizeText(settings.monthCols) ? settings.monthCols : els.monthCols.value;
    var cols = XIO.parseColumns(text).cols;
    var i = cols.indexOf(c);
    if (i >= 0) cols.splice(i, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    settings.monthCols = compact(cols);
    saveSettings();
    els.exportNote.textContent = '';
    var scroller = els.mainPreview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    refresh();
    scroller = els.mainPreview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;
    var btn = els.mainPreview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  function clearChoices() {
    els.wipCol.textContent = '';
    els.wipCol.appendChild(h('option', { value: '', text: '請選擇' }));
    els.ctInfo.textContent = '';
    els.mainPreview.textContent = '';
    els.headers.textContent = '';
    els.headers.removeAttribute('data-cols');
  }

  // 每個月一個輸入框，由使用者輸入月份，記在 S.headerEdits。
  // 月份的行沒變就不重畫，打字時才不會被打斷。
  function renderHeaderInputs(monthCols) {
    var key = monthCols.join(',');
    if (els.headers.getAttribute('data-cols') === key && els.headers.childNodes.length) return;
    els.headers.setAttribute('data-cols', key);
    els.headers.textContent = '';
    monthCols.forEach(function (mc, i) {
      var id = 'f4-head-' + mc;
      var input = h('input', {
        id: id, type: 'text', class: 'in in-month', autocomplete: 'off',
        value: S.headerEdits[mc] || '', placeholder: i === 0 ? '例如 Oct' : null
      });
      input.addEventListener('input', function () {
        S.headerEdits[mc] = input.value;
        els.exportNote.textContent = '';
        clearTimeout(timer);
        timer = setTimeout(function () { refresh(true); }, 250);
      });
      els.headers.appendChild(h('div', { class: 'month-head' },
        h('label', { for: id, text: XIO.numToCol(mc) + ' 行' + (i === 0 ? '（當月）' : '') }), input));
    });
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function listText(labels, max) {
    return labels.slice(0, max).join('、') + (labels.length > max ? ' …等 ' + labels.length + ' 個' : '');
  }

  function renderNotes(plan, monthCols) {
    if (plan.noCt) {
      var none = plan.rows.filter(function (it) { return it.noCt; }).map(function (it) { return it.label; });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + plan.noCt + ' 個料號沒有 CT.（第 4 步在 P003 找不到）：' + listText(none, 12) +
        '。這些列一樣會填公式，但需求工時會算成 0，並標上淺紅色底。'));
    }
    if (plan.errors) {
      var bad = plan.rows.filter(function (it) {
        return it.values.some(function (v) { return v && typeof v === 'object'; });
      }).map(function (it) { return it.label; });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + plan.errors + ' 列算不出來（CT. 是錯誤值，或數量不是數字）：' + listText(bad, 8) +
        '。Excel 裡這些格會顯示 #DIV/0! 或 #VALUE!。'));
    }
    if (monthCols.length !== MONTH_COUNT) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '目前選了 ' + monthCols.length + ' 個月份，通常是 ' + MONTH_COUNT + ' 個（當月加往後五個月）。'));
    }
  }

  function showNumber(v) {
    if (v && typeof v === 'object') return v.error;
    if (typeof v !== 'number') return v === null || v === undefined ? '' : String(v);
    return Math.round(v).toLocaleString('en-US');          // 和 Excel 一樣只顯示整數
  }

  // 前幾筆的計算結果，以及第一筆每一格的公式
  function renderResult(plan, opts) {
    els.result.textContent = '';
    var SHOW = 6;
    var n = opts.monthCols.length;
    var top = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: 'MPS 列' }),
      h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: '料號' }),
      h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: 'CT.' }),
      h('th', { class: 'colhead', scope: 'colgroup', colspan: String(n), text: TITLE }));
    var sub = h('tr');
    plan.headers.forEach(function (text, i) {
      sub.appendChild(h('th', { class: 'colhead', scope: 'col', text: XIO.numToCol(plan.startCol + i) + '　' + text }));
    });
    var body = h('tbody');
    plan.rows.slice(0, SHOW).forEach(function (it) {
      var tr = h('tr', { class: it.noCt ? 'is-missing' : null },
        h('th', { class: 'rowhead', scope: 'row', text: String(it.row) }),
        h('td', { text: it.label }),
        h('td', { class: 'is-num', text: it.noCt ? '沒有' : (typeof it.ct === 'number' ? String(Math.round(it.ct * 1e6) / 1e6) : String(it.ct)) }));
      it.values.forEach(function (v) { tr.appendChild(h('td', { class: 'is-num is-kept', text: showNumber(v) })); });
      body.appendChild(tr);
    });
    var more = plan.rows.length - SHOW;
    els.result.appendChild(h('p', {
      class: 'sheet-caption',
      text: '計算結果（前 ' + Math.min(SHOW, plan.rows.length) + ' 筆）。' + (more > 0 ? '還有 ' + more.toLocaleString('en-US') + ' 筆沒顯示。' : '')
    }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '計算結果' },
      h('table', { class: 'sheet is-result' }, h('thead', null, top, sub), body)));

    var first = plan.rows[0];
    var formulas = first.formulas.map(function (f, i) {
      return XIO.numToCol(plan.startCol + i) + first.row + ' =' + f;
    });
    els.result.appendChild(h('p', { class: 'hint', text: '第 ' + first.row + ' 列填入的公式：' + formulas.join('　') }));
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      var main = pickMain(wb, src);
      XIO.addDemandColumns(main, opts);
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
    var ready = IEApp.state.withDemand;
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

  IEApp.registerStep({ id: 'demand-hours', render: render });
})(window);
