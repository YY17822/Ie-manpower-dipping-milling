/*
 * 第 4 步：帶入標準工時 CT.
 *   拿 MPS 工作表的料號（某一行），到 P003 工作表的某一行找相同的料號，
 *   找到後取那一列的 Labor ÷ Base Qty（P003 的 Labor 是做 Base Qty 個的工時，
 *   除過之後才是單件工時），填到 MPS 工作表「原本最後一行之後空一行」的位置，表頭是「CT.」。
 *
 *   填入的是公式（例如 ='P003'!M5/'P003'!K5），之後在 Excel 點儲存格就能看到數值是從 P003 哪兩格算出來的。
 *   同一個料號在 P003 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）。
 *   找不到的料號留空白，並標上淺紅色底。
 *
 *   來源有兩種：
 *     a. 剛才第 3 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 3 步匯出的檔案（裡面要有 MPS 和 P003 兩個工作表）
 *
 * 做完之後把結果放在 IEApp.state.withCT，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = IEApp.storeKey('step3');       // 鍵名是舊的步驟編號，改掉會讓記住的設定不見，所以沿用
  var STEP1_KEY = IEApp.storeKey('step1');
  var REF_SHEET = 'P003';
  var HEADER = 'CT.';
  var SUFFIX = '_CT';
  var KEY_HINT = /p\s*\/?\s*n|料號|part|material|型號/i;   // 猜哪一行是料號
  var NOT_KEY = /desc|說明|品名/i;
  var LABOR_HINT = /labor|人工|工時/i;                      // 猜哪一行是 Labor
  var BASE_HINT = /base\s*qty|base\s*quantity|基準|基本數量/i;  // 猜哪一行是 Base Qty

  var DEFAULTS = { keyCol: '', refKeyCol: '', refValueCol: '', refDivCol: '', startRow: '' };

  // 這一步的狀態（切到別的步驟再回來時還在）
  var S = {
    upload: null,         // 上傳的第 3 步檔案 { name, buffer }
    preferUpload: false,  // 兩種來源都有時，用哪一個
    work: null,           // 目前拿來預覽、統計的活頁簿 { src, loading, error, wb, main, ref, mainExt, refExt, refHeaderRow }
    index: null,          // P003 料號索引
    indexKey: null
  };
  var settings = loadSettings();
  var els = {};

  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    return out;
  }
  function saveSettings() {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(settings)); } catch (e) { /* 存不了就算了 */ }
  }

  function isRefName(name) { return String(name).toLowerCase() === REF_SHEET.toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.withP003;
    if (S.upload && (S.preferUpload || !prev)) return { kind: 'file', key: S.upload, name: S.upload.name };
    if (prev) return { kind: 'prev', key: prev, name: prev.outName, prev: prev };
    return null;
  }

  // 來源對應的全新活頁簿（Promise）
  function buildFrom(src) {
    return src.kind === 'prev' ? src.prev.build() : XIO.load(src.key.buffer.slice(0));
  }

  function pickSheets(wb, src) {
    var ref = wb.worksheets.filter(function (w) { return isRefName(w.name); })[0] || null;
    var others = wb.worksheets.filter(function (w) { return !isRefName(w.name); });
    var wanted = src.kind === 'prev' ? src.prev.mainSheet : null;
    var main = others.filter(function (w) { return w.name === wanted; })[0] || others[0] || null;
    return { main: main, ref: ref };
  }

  // MPS 的資料從第幾列開始：接著第 1、3 步做下來的話沿用第 1 步的設定
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

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f3-base',
      title: '上傳第 3 步匯出的 Excel',
      sub: '或把檔案拖進這裡。裡面要有 MPS 和 P003 兩個工作表',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 3 步匯入 P003',
      onclick: function () { IEApp.go('p003-import'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.keyCol = h('select', { id: 'f3-key', class: 'in-select' });
    els.mainPreview = h('div', { class: 'sheet-area' });
    els.startRow = h('input', { id: 'f3-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric' });
    els.startNote = h('span');

    els.refKeyCol = h('select', { id: 'f3-refkey', class: 'in-select' });
    els.refValueCol = h('select', { id: 'f3-refval', class: 'in-select' });
    els.refDivCol = h('select', { id: 'f3-refdiv', class: 'in-select' });
    els.refPreview = h('div', { class: 'sheet-area' });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.result = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：計算需求工時',
      onclick: function () { IEApp.go('demand-hours'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '帶入標準工時 CT.' }),
      h('p', { text: '拿 MPS 的料號到 P003 找同一個料號，把那一列的 Labor ÷ Base Qty（單件工時）帶回 MPS，放在最後一行之後空一行的位置，表頭是「CT.」。填入的是公式，在 Excel 點儲存格就能看到數值是從 P003 哪兩格算出來的。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 3 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg),
      h('li', { class: 'task' },
        h('h2', { text: 'MPS 用哪一行對照' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f3-key', text: '料號在' }), els.keyCol),
        els.mainPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f3-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: 'P003 的料號、Labor、Base Qty' }),
        h('div', { class: 'field-grid' },
          h('label', { for: 'f3-refkey', text: '料號在' }), els.refKeyCol,
          h('label', { for: 'f3-refval', text: 'Labor 在' }), els.refValueCol,
          h('label', { for: 'f3-refdiv', text: 'Base Qty 在' }), els.refDivCol),
        h('p', { class: 'hint', text: 'CT. = Labor ÷ Base Qty。P003 的 Labor 是做 Base Qty 個的工時，除過之後才是單件工時。' }),
        els.refPreview),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes, els.result,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    bindSelect(els.keyCol, 'keyCol');
    bindSelect(els.refKeyCol, 'refKeyCol');
    bindSelect(els.refValueCol, 'refValueCol');
    bindSelect(els.refDivCol, 'refDivCol');
    els.startRow.addEventListener('change', function () {
      settings.startRow = els.startRow.value;
      saveSettings();
      refresh();
    });

    refresh();
  }

  function bindSelect(select, key) {
    select.addEventListener('change', function () {
      settings[key] = select.value;
      saveSettings();
      els.exportNote.textContent = '';
      refresh();
    });
  }

  /* ---------- 來源 ---------- */
  function openUpload(file) {
    els.srcMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (w) { return w.name; });
        if (!names.some(isRefName)) {
          throw new Error('這個檔案沒有 P003 工作表。請上傳第 3 步匯出的檔案，或先回到第 3 步匯入 P003。');
        }
        if (names.length < 2) {
          throw new Error('這個檔案只有 P003 工作表，沒有 MPS。請上傳第 3 步匯出的檔案。');
        }
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
    S.index = null;
    buildFrom(src).then(function (wb) {
      if (S.work !== work) return;               // 這段時間來源又換了
      var sheets = pickSheets(wb, src);
      if (!sheets.main || !sheets.ref) throw new Error('找不到 MPS 或 P003 工作表。');
      work.wb = wb;
      work.main = sheets.main;
      work.ref = sheets.ref;
      work.mainExt = XIO.extent(sheets.main);
      work.refExt = XIO.extent(sheets.ref);
      work.refHeaderRow = XIO.guessHeaderRow(sheets.ref, 10);
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  /* ---------- 下拉選單 ---------- */
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

  function guessKeyColumn(labels, colCount) {
    for (var c = 1; c <= colCount; c++) {
      if (labels[c] && KEY_HINT.test(labels[c]) && !NOT_KEY.test(labels[c])) return c;
    }
    return 0;
  }
  function guessColumn(labels, colCount, hint) {
    for (var c = 1; c <= colCount; c++) {
      if (labels[c] && hint.test(labels[c])) return c;
    }
    return 0;
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = IEApp.state.withP003;
    IEApp.state.withCT = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.result.textContent = '';

    // 來源
    if (!src) {
      els.srcInfo.textContent = '還沒有第 3 步的結果。回到第 3 步匯入 P003，或上傳之前匯出的檔案。';
      els.drop.setIdle();
    } else if (src.kind === 'prev') {
      els.srcInfo.textContent = '沿用第 3 步剛做好的結果：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.upload) els.drop.setFile(S.upload.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.drop.setIdle();
    } else {
      els.srcInfo.textContent = '使用上傳的檔案。';
      els.drop.setFile(S.upload.name);
    }
    els.backBtn.hidden = !!src;
    els.switchBtn.hidden = !(S.upload && prev);
    if (src) els.switchBtn.textContent = src.kind === 'file' ? '改用第 3 步剛做好的結果' : '改用上傳的檔案';

    var chain = chainStartRow();
    els.startRow.value = String(startRow());
    els.startRow.disabled = !!chain;
    els.startNote.textContent = chain ? '列開始（沿用第 1 步的設定）' : '列開始';

    if (!src) {
      clearChoices();
      setSummary('先準備好第 3 步的 Excel。');
      return;
    }

    ensureWork(src);
    var work = S.work;
    if (work.loading) { clearChoices(); setSummary('準備資料中…'); return; }
    if (work.error) { clearChoices(); setSummary(work.error, true); return; }

    var sr = startRow();
    var headerEnd = sr - 1;
    var mainCols = Math.max(work.mainExt.colCount, 1);
    var refCols = Math.max(work.refExt.colCount, 1);
    var mainLabels = XIO.columnLabels(work.main, headerEnd >= 2 ? 2 : 1, Math.max(headerEnd, 1), mainCols);
    var refLabels = XIO.columnLabels(work.ref, work.refHeaderRow, work.refHeaderRow, refCols);

    var keyCol = colOf(settings.keyCol, mainCols) || guessKeyColumn(mainLabels, mainCols);
    var refKeyCol = colOf(settings.refKeyCol, refCols) || guessKeyColumn(refLabels, refCols);
    var refValueCol = colOf(settings.refValueCol, refCols) || guessColumn(refLabels, refCols, LABOR_HINT);
    var refDivCol = colOf(settings.refDivCol, refCols) || guessColumn(refLabels, refCols, BASE_HINT);

    fillSelect(els.keyCol, mainLabels, mainCols, keyCol);
    fillSelect(els.refKeyCol, refLabels, refCols, refKeyCol);
    fillSelect(els.refValueCol, refLabels, refCols, refValueCol);
    fillSelect(els.refDivCol, refLabels, refCols, refDivCol);

    var keepMain = {}; if (keyCol) keepMain[keyCol] = true;
    ui.renderSheet(els.mainPreview, work.main, {
      headerEnd: headerEnd, lastRow: work.mainExt.lastRow, colCount: mainCols, rows: 5, keep: keepMain,
      caption: work.main.name + ' 工作表（前 5 列資料）。共 ' +
        Math.max(0, work.mainExt.lastRow - headerEnd).toLocaleString('en-US') + ' 列資料。'
    });
    var keepRef = {};
    [refKeyCol, refValueCol, refDivCol].forEach(function (c) { if (c) keepRef[c] = true; });
    ui.renderSheet(els.refPreview, work.ref, {
      headerEnd: work.refHeaderRow, lastRow: work.refExt.lastRow, colCount: refCols, rows: 6, keep: keepRef,
      caption: work.ref.name + ' 工作表（前 6 列資料）。共 ' + work.refExt.lastRow.toLocaleString('en-US') + ' 列。'
    });

    if (!keyCol) { setSummary('選擇 MPS 的料號在哪一行。'); return; }
    if (!refKeyCol) { setSummary('選擇 P003 的料號在哪一行。'); return; }
    if (!refValueCol) { setSummary('選擇 P003 的 Labor 在哪一行。'); return; }
    if (!refDivCol) { setSummary('選擇 P003 的 Base Qty 在哪一行。'); return; }
    if (refDivCol === refValueCol) { setSummary('Labor 和 Base Qty 選到同一行了。請確認兩個各是哪一行。', true); return; }

    var indexKey = refKeyCol;
    if (!S.index || S.indexFor !== work || S.indexKey !== indexKey) {
      S.index = XIO.buildIndex(work.ref, refKeyCol);
      S.indexFor = work;
      S.indexKey = indexKey;
    }
    var opts = {
      dataStartRow: sr, keyCol: keyCol, refKeyCol: refKeyCol,
      refValueCol: refValueCol, refDivCol: refDivCol, header: HEADER
    };
    var plan = XIO.planLookup(work.main, work.ref, Object.assign({ index: S.index }, opts));

    var keyL = XIO.numToCol(keyCol), refKeyL = XIO.numToCol(refKeyCol);
    var targetL = XIO.numToCol(plan.targetCol);

    if (!plan.rows.length) {
      setSummary('MPS 的 ' + keyL + ' 行從第 ' + sr + ' 列起沒有料號。請確認選的行對不對。', true);
      return;
    }
    if (!plan.found) {
      setSummary('MPS 的 ' + keyL + ' 行和 P003 的 ' + refKeyL + ' 行沒有任何相同的料號。請確認兩邊選的行對不對。', true);
      renderResult(plan, work.ref.name, opts);
      return;
    }

    setSummary('CT. 會填在 ' + work.main.name + ' 工作表的 ' + targetL + ' 行' +
      (plan.isNew ? '（原本最後一行之後空一行）' : '（這一行原本就是 CT.，會直接覆蓋）') +
      '。' + plan.rows.length.toLocaleString('en-US') + ' 個料號裡找到 ' +
      plan.found.toLocaleString('en-US') + ' 個' +
      (plan.missing ? '，找不到 ' + plan.missing.toLocaleString('en-US') + ' 個。' : '，全部都有找到。'),
      false, plan.missing > 0);

    renderNotes(plan);
    renderResult(plan, work.ref.name, opts);
    els.exportBtn.disabled = false;

    // 給下一步用
    IEApp.state.withCT = {
      outName: outputName(src),
      mainSheet: work.main.name,
      dataStartRow: sr,
      keyCol: keyCol,
      ctCol: plan.targetCol,
      build: function () { return buildWorkbook(src, opts); }
    };
    els.nextBtn.disabled = false;
  }

  function clearChoices() {
    [els.keyCol, els.refKeyCol, els.refValueCol, els.refDivCol].forEach(function (sel) {
      sel.textContent = '';
      sel.appendChild(h('option', { value: '', text: '請選擇' }));
    });
    els.mainPreview.textContent = '';
    els.refPreview.textContent = '';
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function uniqueLabels(rows, test) {
    var seen = {}, out = [];
    rows.forEach(function (it) {
      if (!test(it) || seen[it.key]) return;
      seen[it.key] = true;
      out.push(it.label);
    });
    return out;
  }
  function listText(labels, max) {
    return labels.slice(0, max).join('、') + (labels.length > max ? ' …等 ' + labels.length + ' 個' : '');
  }

  function renderNotes(plan) {
    if (plan.missing) {
      var miss = uniqueLabels(plan.rows, function (it) { return !it.refRow; });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        'P003 找不到的料號：' + listText(miss, 12) + '。這些列的 CT. 會留空白，並標上淺紅色底。'));
    }
    if (plan.multi) {
      var text = '有 ' + plan.multi + ' 個料號在 P003 出現不只一列，一律取最上面那一列。';
      if (plan.multiDiff) {
        var diff = uniqueLabels(plan.rows, function (it) { return it.differs; });
        text += '其中 ' + plan.multiDiff + ' 個各列算出來的 CT. 不一樣，請留意：' + listText(diff, 8) + '。';
      } else {
        text += '這些料號各列算出來的 CT. 都相同，不影響結果。';
      }
      els.notes.appendChild(h('li', { class: plan.multiDiff ? 'is-warn' : null }, text));
    }
    if (plan.errors) {
      var errs = uniqueLabels(plan.rows, function (it) { return !!it.error; });
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + plan.errors + ' 列算不出 CT.（Base Qty 是空白或 0，或 Labor、Base Qty 不是數字）：' +
        listText(errs, 8) + '。Excel 裡這些格會顯示 #DIV/0! 或 #VALUE!。'));
    }
    if (plan.blank) {
      els.notes.appendChild(h('li', null, '另有 ' + plan.blank + ' 列的料號是空白，略過不填。'));
    }
  }

  function showNumber(v) {
    if (v === null || v === undefined) return '';
    if (typeof v !== 'number') return String(v);
    return String(Math.round(v * 1e6) / 1e6);
  }

  // 前幾筆的對照結果：哪個料號、對到 P003 第幾列、Labor、Base Qty、公式和算出來的 CT.
  function renderResult(plan, refName, opts) {
    els.result.textContent = '';
    var SHOW = 8;
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: 'MPS 列' }),
      h('th', { class: 'colhead', scope: 'col', text: '料號' }),
      h('th', { class: 'colhead', scope: 'col', text: 'P003 列' }),
      h('th', { class: 'colhead', scope: 'col', text: 'Labor' }),
      h('th', { class: 'colhead', scope: 'col', text: 'Base Qty' }),
      h('th', { class: 'colhead', scope: 'col', text: '填入的公式' }),
      h('th', { class: 'colhead', scope: 'col', text: 'CT.' }));
    var body = h('tbody');
    plan.rows.slice(0, SHOW).forEach(function (it) {
      body.appendChild(h('tr', { class: it.refRow ? null : 'is-missing' },
        h('th', { class: 'rowhead', scope: 'row', text: String(it.row) }),
        h('td', { text: it.label }),
        h('td', { class: 'is-num', text: it.refRow ? String(it.refRow) + (it.multi > 1 ? '（共 ' + it.multi + ' 列）' : '') : '找不到' }),
        h('td', { class: 'is-num', text: it.refRow ? it.top : '' }),
        h('td', { class: 'is-num', text: it.refRow ? it.bottom : '' }),
        h('td', { text: it.refRow ? '=' + XIO.lookupFormula(refName, it.refRow, opts) : '' }),
        h('td', { class: 'is-num is-kept', text: it.error ? it.error : showNumber(it.value) })));
    });
    var more = plan.rows.length - SHOW;
    els.result.appendChild(h('p', {
      class: 'sheet-caption',
      text: '對照結果（前 ' + Math.min(SHOW, plan.rows.length) + ' 筆）。' + (more > 0 ? '還有 ' + more.toLocaleString('en-US') + ' 筆沒顯示。' : '')
    }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '對照結果' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));
  }

  /* ---------- 產生活頁簿 ---------- */
  // 每次都從頭產生一份新的，再把 CT. 填進去。回傳 Promise<Workbook>
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      var sheets = pickSheets(wb, src);
      XIO.addLookupColumn(sheets.main, sheets.ref, opts);
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
    var ready = IEApp.state.withCT;
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

  IEApp.registerStep({ id: 'ct-lookup', render: render });
})(window);
