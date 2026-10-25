/*
 * 第 5 步：計算需求工時（Dipping_Milling）
 *   每個料號各月的計畫產量，分別乘上「機台 CT.」和「目檢 CT.」，算出兩塊需求工時：
 *
 *     機台需求工時   各月 = 數量 * 機台 CT.                 例如 =E6*L6、=F6*L6
 *     目檢需求工時   各月 = 數量 * 目檢 CT. * 目檢-檢驗率   例如 =E6*M6*N6、=F6*M6*N6
 *   當月也只用當月產量，不加 WIP。
 *
 *   位置：「目檢-檢驗率」之後空一行開始放機台需求工時（六個月），再空一行放目檢需求工時（六個月）。
 *   只算 Cutting、Dipping 兩段的料號列（標題、空白列、第 2 步的 Total 列都跳過）。每格都是公式。
 *   表頭：第 2 列是跨所有月份置中的「機台需求工時」／「目檢需求工時」，第 3 列到表頭最後一列是各月的月份。
 *        月份由使用者自己輸入，兩塊用同一組月份。
 *   每一塊在每一段（Milling = Cutting、Dipping）最後一列下面空一列加一列 Total（和第 2 步的 Total 同一列）：
 *     「Total」放在這一塊左邊那一行，各月 =SUM(P6:P8)；淺綠底，外框粗線、裡面細線。
 *   數值只顯示整數（儲存格裡仍保留小數，加總時用的是完整的數值）。
 *
 *   預設的行（照第 1 步保留的順序）：第 5 到第 10 行是當月和往後五個月。
 *
 *   來源有兩種：
 *     a. 剛才第 4 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 4 步匯出的檔案（MPS 工作表裡要有「機台 CT.」「目檢 CT.」這兩行）
 *
 * 做完之後把結果放在 IEApp.state.withDemand，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = 'step5';             // 共用設定的組名（IEApp.settings）
  var STEP1_KEY = 'step1';
  var REF_SHEET = 'P003';
  var CT_HEADERS = XIO.CT_HEADERS;          // { machine: '機台 CT.', labor: '目檢 CT.' }
  var TITLES = XIO.DEMAND_TITLES;           // { machine: '機台需求工時', labor: '目檢需求工時' }
  var SUFFIX = '_需求工時';
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

  // 月份的行和第 2 步（匯總）共用同一份設定（wipCol 已不使用，只是原樣保留）；資料起始列是這一步自己的。
  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var shared = ui.columns.get();
    return {
      wipCol: shared.wipCol,
      monthCols: shared.monthCols,
      startRow: s.startRow === undefined ? DEFAULTS.startRow : s.startRow
    };
  }
  function saveSettings() {
    try { IEApp.settings.write(STORE_KEY, { startRow: settings.startRow }); } catch (e) { /* 存不了就算了 */ }
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
      var s1 = IEApp.settings.read(STEP1_KEY) || {};
      n = parseInt(XIO.normalizeText(s1.startRow === undefined ? '' : s1.startRow), 10);
    } catch (e) { n = 0; }
    return n >= 1 ? n : 5;
  }


  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};
    settings = loadSettings();        // 每次進來都重讀：第 2 步可能改過月份的行

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f4-base',
      title: '上傳第 4 步匯出的 Excel',
      sub: '或把檔案拖進這裡。MPS 工作表裡要有機台 CT.、目檢 CT. 這兩行',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 4 步帶入 CT.',
      onclick: function () { IEApp.go('ct-lookup'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

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
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '計算需求工時' }),
      h('p', { text: '把每個料號各月的計畫產量乘上機台 CT. 算出「機台需求工時」；乘上目檢 CT. 再乘目檢-檢驗率算出「目檢需求工時」。當月也只用當月產量，不加 WIP。目檢-檢驗率之後空一行放機台需求工時，再空一行放目檢需求工時。只算 Cutting、Dipping 的料號列，每格都是公式；每一段下面空一列加一列 Total。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 4 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: '各月產量在哪幾行' }),
        h('div', { class: 'field-grid' },
          h('label', { for: 'f4-months', text: '各月產量在' }), els.monthCols,
          h('span', { class: 'field-label', text: 'CT. 在' }), els.ctInfo),
        h('p', { class: 'hint', text: '各月產量的第一行是當月；連續的行可以寫成 E:J，也可以點下面表格最上方的行號來選。這裡的設定和第 2 步（匯總）共用。CT. 和檢驗率會自動找表頭是「機台 CT.」「目檢 CT.」「目檢-檢驗率」的三行。' }),
        els.mainPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f4-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: '各月的表頭' }),
        h('p', { class: 'hint', text: '當月用下拉選單選月份，再按「自動帶出後面月份」就會依序填好後面幾個月；後面的月份也可以自己改。機台需求工時和目檢需求工時用同一組。' }),
        els.headers),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.result,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

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

    // 機台 CT.、目檢 CT. 在哪一行：找表頭
    var machineCtCol = XIO.findHeaderColumn(work.main, headerEnd, CT_HEADERS.machine);
    var laborCtCol = XIO.findHeaderColumn(work.main, headerEnd, CT_HEADERS.labor);
    var laborRateCol = XIO.findHeaderColumn(work.main, headerEnd, CT_HEADERS.rate);
    els.ctInfo.textContent = '機台 CT. ' + (machineCtCol ? XIO.numToCol(machineCtCol) + ' 行' : '找不到') +
      '、目檢 CT. ' + (laborCtCol ? XIO.numToCol(laborCtCol) + ' 行' : '找不到') +
      '、目檢-檢驗率 ' + (laborRateCol ? XIO.numToCol(laborRateCol) + ' 行' : '找不到');
    var ctCol = machineCtCol && laborCtCol && laborRateCol ? Math.min(machineCtCol, laborCtCol, laborRateCol) : 0;

    // 月份：先用記住的，沒有就用預設的位置（第 5 到 10 行）
    var limit = ctCol ? ctCol - 1 : cols;        // 數量都在 CT. 左邊
    var parsed = XIO.parseColumns(settings.monthCols);
    var monthCols = parsed.cols;
    if (!XIO.normalizeText(settings.monthCols)) {
      monthCols = [];
      for (var c = MONTH_FIRST; c < MONTH_FIRST + MONTH_COUNT && c <= limit; c++) monthCols.push(c);
    }

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
      var lack = [machineCtCol ? '' : '「機台 CT.」', laborCtCol ? '' : '「目檢 CT.」', laborRateCol ? '' : '「目檢-檢驗率」']
        .filter(Boolean).join('、');
      setSummary('MPS 工作表裡找不到表頭是' + lack + '的行。請先做第 4 步帶入 CT.。', true);
      return;
    }
    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 E:J。', true);
      return;
    }
    if (!monthCols.length) { setSummary('輸入各月產量在哪幾行。'); return; }
    var clash = monthCols.filter(function (mc) { return mc === machineCtCol || mc === laborCtCol || mc === laborRateCol || mc > cols; });
    if (clash.length) {
      setSummary(XIO.formatColumns(clash) + ' 行不能當作各月產量（和 CT. 重複，或超出報表範圍）。', true);
      return;
    }

    var headers = monthCols.map(function (mc) { return S.headerEdits[mc] || ''; });
    var emptyHead = monthCols.filter(function (mc, i) { return !XIO.normalizeText(headers[i]); });
    if (emptyHead.length) {
      setSummary('在上面「各月的表頭」輸入 ' + XIO.formatColumns(emptyHead) + ' 行的月份。');
      return;
    }

    var keyCol = src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;
    var rows = XIO.sectionDataRows(work.main, headerEnd);
    if (!rows.length) {
      setSummary('MPS 工作表找不到 Cutting、Dipping 的料號列（第 1 步寫的「Cutting（…）」「Dipping（…）」標題下面）。請確認上傳的是這一版第 4 步匯出的檔案。', true);
      return;
    }
    var opts = {
      dataStartRow: sr, keyCol: keyCol, monthCols: monthCols,
      machineCtCol: machineCtCol, laborCtCol: laborCtCol, laborRateCol: laborRateCol, headers: headers, rows: rows
    };
    var plans = XIO.planDemandBlocks(work.main, opts);
    var plan = plans.machine;
    if (!plan.rows.length) {
      setSummary('Cutting、Dipping 兩段裡沒有料號。請確認上傳的檔案。', true);
      return;
    }

    var n = monthCols.length;
    function span(p) { return XIO.numToCol(p.startCol) + ' 到 ' + XIO.numToCol(p.startCol + n - 1) + ' 行'; }
    var overwrite = !plans.machine.isNew || !plans.labor.isNew;
    setSummary('機台需求工時會填在 ' + work.main.name + ' 工作表的 ' + span(plans.machine) +
      '、目檢需求工時在 ' + span(plans.labor) +
      (overwrite ? '（原本就有，會直接覆蓋）' : '（目檢-檢驗率之後空一行，兩塊中間再空一行）') +
      '，共 ' + plan.rows.length.toLocaleString('en-US') + ' 個料號、' + n + ' 個月。兩塊在 Milling、Dipping 各自最後一列下面空一列加一列 Total（=SUM(…)）。',
      false, plans.machine.noCt > 0 || plans.labor.noCt > 0 || plans.machine.errors > 0 || plans.labor.errors > 0);

    renderNotes(plans, monthCols);
    renderResult(plans, opts);
    els.exportBtn.disabled = false;

    // 給下一步用
    IEApp.state.withDemand = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: sr,
      keyCol: keyCol,
      machineDemandCol: plans.machine.startCol,
      laborDemandCol: plans.labor.startCol,
      monthCount: n,
      rows: rows,
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
    els.ctInfo.textContent = '';
    els.mainPreview.textContent = '';
    els.headers.textContent = '';
    els.headers.removeAttribute('data-cols');
  }

  // 每個月一個輸入框，記在 S.headerEdits。
  // 第一個（當月）是下拉選單，選 Jan～Dec；其餘是文字框，可以自己改。
  // 最後一個框後面有「自動帶出後面月份」按鈕：依第一個的月份往後排，Dec 之後接 Jan。
  // 月份的行沒變就不重畫，打字時才不會被打斷。
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function onHeaderChange(mc, value) {
    S.headerEdits[mc] = value;
    els.exportNote.textContent = '';
    clearTimeout(timer);
    timer = setTimeout(function () { refresh(true); }, 250);
  }

  function renderHeaderInputs(monthCols) {
    var key = monthCols.join(',');
    if (els.headers.getAttribute('data-cols') === key && els.headers.childNodes.length) return;
    els.headers.setAttribute('data-cols', key);
    els.headers.textContent = '';
    var fields = [];
    var fillBtn = null;
    monthCols.forEach(function (mc, i) {
      var id = 'f4-head-' + mc;
      var value = S.headerEdits[mc] || '';
      var field;
      if (i === 0) {
        field = h('select', { id: id, class: 'in in-month' });
        field.appendChild(h('option', { value: '', text: '請選擇' }));
        if (value && MONTHS.indexOf(value) < 0) field.appendChild(h('option', { value: value, text: value }));
        MONTHS.forEach(function (m) { field.appendChild(h('option', { value: m, text: m })); });
        field.value = value;
        field.addEventListener('change', function () {
          onHeaderChange(mc, field.value);
          if (fillBtn) fillBtn.disabled = MONTHS.indexOf(field.value) < 0;
        });
      } else {
        field = h('input', {
          id: id, type: 'text', class: 'in in-month', autocomplete: 'off', value: value
        });
        field.addEventListener('input', function () { onHeaderChange(mc, field.value); });
      }
      fields.push(field);
      els.headers.appendChild(h('div', { class: 'month-head' },
        h('label', { for: id, text: XIO.numToCol(mc) + ' 行' + (i === 0 ? '（當月）' : '') }), field));
    });

    if (monthCols.length > 1) {
      fillBtn = h('button', {
        type: 'button', class: 'btn-secondary btn-fill-months', id: 'f4-head-fill',
        text: '自動帶出後面月份', disabled: MONTHS.indexOf(fields[0].value) < 0,
        onclick: function () {
          var first = MONTHS.indexOf(fields[0].value);
          if (first < 0) return;
          for (var i = 1; i < fields.length; i++) {
            var m = MONTHS[(first + i) % 12];
            fields[i].value = m;
            S.headerEdits[monthCols[i]] = m;
          }
          els.exportNote.textContent = '';
          clearTimeout(timer);
          refresh(true);
        }
      });
      els.headers.appendChild(h('div', { class: 'month-head month-head-btn' }, fillBtn));
    }
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function listText(labels, max) {
    return labels.slice(0, max).join('、') + (labels.length > max ? ' …等 ' + labels.length + ' 個' : '');
  }

  function renderNotes(plans, monthCols) {
    [['machine', '機台 CT.', '機台需求工時'], ['labor', '目檢 CT.', '目檢需求工時']].forEach(function (k) {
      var plan = plans[k[0]];
      if (!plan.noCt) return;
      var none = plan.rows.filter(function (it) { return it.noCt; }).map(function (it) { return it.label; });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + plan.noCt + ' 個料號沒有' + k[1] + '（第 4 步在 P003 找不到）：' + listText(none, 12) +
        '。這些列一樣會填公式，但' + k[2] + '會算成 0，並標上淺紅色底。'));
    });
    var badKeys = {};
    [plans.machine, plans.labor].forEach(function (plan) {
      plan.rows.forEach(function (it) {
        if (it.values.some(function (v) { return v && typeof v === 'object'; })) badKeys[it.row] = it.label;
      });
    });
    var bad = Object.keys(badKeys).map(function (k) { return badKeys[k]; });
    if (bad.length) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + bad.length + ' 列算不出來（CT. 或檢驗率是錯誤值，例如 Base Qty 是 0，或數量不是數字）：' + listText(bad, 8) +
        '。Excel 裡這些格會顯示 #DIV/0! 或 #VALUE!。'));
    }
    var noRate = plans.labor.rows.filter(function (it) { return it.noRate; }).map(function (it) { return it.label; });
    if (noRate.length) {
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + noRate.length + ' 個料號的目檢-檢驗率是空白：' + listText(noRate, 8) +
        '。目檢需求工時會算成 0。請回到第 4 步輸入檢驗率。'));
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

  function rateText(it) {
    if (it.noRate) return '空白';
    return typeof it.rate === 'number' ? Math.round(it.rate * 10000) / 100 + '%' : String(it.rate);
  }

  function ctText(it) {
    return it.noCt ? '沒有' : (typeof it.ct === 'number' ? String(Math.round(it.ct * 1e6) / 1e6) : String(it.ct));
  }

  // 前幾筆的計算結果（機台、目檢兩塊各一張表），以及第一筆每一格的公式
  function renderResult(plans, opts) {
    els.result.textContent = '';
    var SHOW = 6;
    var n = opts.monthCols.length;
    [['machine', '機台 CT.'], ['labor', '目檢 CT.']].forEach(function (k) {
      var plan = plans[k[0]];
      var title = TITLES[k[0]];
      var withRate = k[0] === 'labor';
      var top = h('tr', null,
        h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: 'MPS 列' }),
        h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: '料號' }),
        h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: k[1] }),
        withRate ? h('th', { class: 'colhead', scope: 'col', rowspan: '2', text: '目檢-檢驗率' }) : null,
        h('th', { class: 'colhead', scope: 'colgroup', colspan: String(n), text: title }));
      var sub = h('tr');
      plan.headers.forEach(function (text, i) {
        sub.appendChild(h('th', { class: 'colhead', scope: 'col', text: XIO.numToCol(plan.startCol + i) + '　' + text }));
      });
      var body = h('tbody');
      plan.rows.slice(0, SHOW).forEach(function (it) {
        var tr = h('tr', { class: it.noCt ? 'is-missing' : null },
          h('th', { class: 'rowhead', scope: 'row', text: String(it.row) }),
          h('td', { text: it.label }),
          h('td', { class: 'is-num', text: ctText(it) }),
          withRate ? h('td', { class: 'is-num', text: rateText(it) }) : null);
        it.values.forEach(function (v) { tr.appendChild(h('td', { class: 'is-num is-kept', text: showNumber(v) })); });
        body.appendChild(tr);
      });
      var more = plan.rows.length - SHOW;
      els.result.appendChild(h('p', {
        class: 'sheet-caption',
        text: title + '（前 ' + Math.min(SHOW, plan.rows.length) + ' 筆）。' + (more > 0 ? '還有 ' + more.toLocaleString('en-US') + ' 筆沒顯示。' : '')
      }));
      els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': title },
        h('table', { class: 'sheet is-result' }, h('thead', null, top, sub), body)));

      var first = plan.rows[0];
      var formulas = first.formulas.slice(0, 2).map(function (f, i) {
        return XIO.numToCol(plan.startCol + i) + first.row + ' =' + f;
      });
      els.result.appendChild(h('p', { class: 'hint', text: '第 ' + first.row + ' 列的公式：' + formulas.join('　') + (n > 2 ? '　…' : '') }));
    });
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      var main = pickMain(wb, src);
      XIO.addDemandBlocks(main, opts);
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
