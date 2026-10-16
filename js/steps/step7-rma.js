/*
 * 第 7 步：RMA 需求工時
 *   前面算的都是正常生產（Production）的需求工時。這一步在匯總那一塊的下面空一列，
 *   加上重工（RMA）的需求工時。項目由使用者在畫面上自己輸入，列數可以增減：
 *
 *     CT.   RMA總需求工時   =SUM(N81:N83)     底下各項目的合計
 *     2.5   項目 A          =120*L81          這個月的重工數量 × 左邊那一格的 CT.
 *     1.2   項目 B          =80*L82
 *
 *   每個項目要輸入：名稱、CT.、各月的重工數量。
 *   數量直接寫在公式裡，CT. 放在項目名稱左邊那一格；之後在 Excel 改 CT.，工時會跟著變。
 *   外觀和匯總那一塊一樣：外框粗線、裡面細格線。
 *
 *   項目名稱和 CT. 會記在這台電腦的瀏覽器裡（每個月大致相同）；數量每次都要重新輸入。
 *
 *   來源有兩種：
 *     a. 剛才第 6 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 6 步匯出的檔案（MPS 工作表裡要有匯總那三列）
 *
 * 做完之後把結果放在 IEApp.state.withRma，下一步（第 8 步匯入人力報表）直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = IEApp.storeKey('step6');       // 鍵名是舊的步驟編號，改掉會讓記住的設定不見，所以沿用
  var STEP1_KEY = IEApp.storeKey('step1');
  var REF_SHEET = 'P003';
  var TITLE = '需求工時';
  var SUMMARY_MARK = '總需求工時';       // 匯總第一列的標籤裡一定有這幾個字
  var LABELS = { title: 'RMA總需求工時', ct: 'CT.' };
  var SUFFIX = '_RMA';
  var DEFAULT_MONTHS = 6;

  var S = {
    upload: null,         // 上傳的第 6 步檔案 { name, buffer }
    preferUpload: false,
    work: null,           // 拿來預覽的活頁簿 { src, loading, error, wb, main }
    items: null           // 使用者輸入的項目 [{ name, ct, qty: ['', '', …] }]（都是輸入框裡的文字）
  };
  var settings = loadSettings();
  var els = {};
  var timer = null;

  function loadSettings() {
    var s = {};
    try { s = JSON.parse(localStorage.getItem(STORE_KEY) || '{}') || {}; } catch (e) { s = {}; }
    return {
      startRow: s.startRow === undefined ? '' : s.startRow,
      items: Array.isArray(s.items) ? s.items : []      // 只記名稱和 CT.
    };
  }
  function saveSettings() {
    settings.items = (S.items || []).map(function (it) { return { name: it.name, ct: it.ct }; })
      .filter(function (it) { return XIO.normalizeText(it.name) !== '' || XIO.normalizeText(it.ct) !== ''; });
    try { localStorage.setItem(STORE_KEY, JSON.stringify(settings)); } catch (e) { /* 存不了就算了 */ }
  }

  function blankItem(months) {
    var qty = [];
    for (var i = 0; i < months; i++) qty.push('');
    return { name: '', ct: '', qty: qty };
  }
  function initItems(months) {
    if (!S.items) {
      S.items = settings.items.map(function (it) {
        var item = blankItem(months);
        item.name = it.name || '';
        item.ct = it.ct === undefined ? '' : String(it.ct);
        return item;
      });
      if (!S.items.length) S.items.push(blankItem(months));
    }
    // 月份數變了就補齊或截掉
    S.items.forEach(function (it) {
      while (it.qty.length < months) it.qty.push('');
      if (it.qty.length > months) it.qty.length = months;
    });
  }

  function isRefName(name) { return String(name).toLowerCase() === REF_SHEET.toLowerCase(); }

  function activeSource() {
    var prev = IEApp.state.withSummary;
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

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.srcInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.drop = ui.fileDrop({
      id: 'f6-base',
      title: '上傳第 6 步匯出的 Excel',
      sub: '或把檔案拖進這裡。MPS 工作表裡要有匯總那三列',
      onFile: openUpload
    });
    els.switchBtn = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 6 步匯總',
      onclick: function () { IEApp.go('summary'); }
    });
    els.srcMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });
    els.startRow = h('input', { id: 'f6-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric' });
    els.startNote = h('span');

    els.entry = h('div', { class: 'entry-area' });
    els.addBtn = h('button', { type: 'button', class: 'btn-secondary btn-small', text: '新增一個項目', onclick: addItem });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.result = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯入人力報表',
      onclick: function () { IEApp.go('manpower'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: 'RMA 需求工時' }),
      h('p', { text: '前面算的是正常生產的需求工時。這一步在匯總下面空一列，加上重工（RMA）的需求工時：每個項目輸入名稱、CT. 和各月的重工數量，工時用公式算（數量 × CT.），第一列是各項目的合計。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 6 步的 Excel' }),
        els.srcInfo, els.drop.el,
        h('div', { class: 'link-row' }, els.switchBtn, els.backBtn),
        els.srcMsg,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f6-start', text: 'MPS 的資料從第' }), els.startRow, els.startNote))),
      h('li', { class: 'task' },
        h('h2', { text: 'RMA 項目' }),
        h('p', { class: 'hint', text: '一列一個項目，可以自己增減。需求工時 = 重工數量 × CT.。項目名稱和 CT. 下次打開還在，數量每次都要重新輸入。' }),
        els.entry,
        h('div', { class: 'actions' }, els.addBtn)),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.result,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

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
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表。請上傳第 6 步匯出的檔案。');
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
      work.loading = false;
      refresh();
    }).catch(function (err) {
      if (S.work !== work) return;
      work.loading = false;
      work.error = err && err.message ? err.message : String(err);
      refresh();
    });
  }

  /* ---------- 項目輸入表 ---------- */
  function addItem() {
    S.items.push(blankItem(S.items.length ? S.items[0].qty.length : DEFAULT_MONTHS));
    els.entry.removeAttribute('data-key');
    refresh();
    var rows = els.entry.querySelectorAll('tbody tr');
    var input = rows.length ? rows[rows.length - 1].querySelector('input') : null;
    if (input) input.focus();
  }

  function removeItem(index) {
    S.items.splice(index, 1);
    if (!S.items.length) S.items.push(blankItem(DEFAULT_MONTHS));
    saveSettings();
    els.entry.removeAttribute('data-key');
    els.exportNote.textContent = '';
    refresh();
  }

  function onType() {
    saveSettings();
    els.exportNote.textContent = '';
    clearTimeout(timer);
    timer = setTimeout(refresh, 250);
  }

  // 列數或月份沒變就不重畫，打字時才不會被打斷
  function renderEntry(headers) {
    var key = S.items.length + '|' + headers.join('\u0001');
    if (els.entry.getAttribute('data-key') === key) return;
    els.entry.setAttribute('data-key', key);
    els.entry.textContent = '';

    var head = h('tr', null,
      h('th', { scope: 'col', text: '項目' }),
      h('th', { scope: 'col', text: 'CT.' }));
    headers.forEach(function (text, i) {
      head.appendChild(h('th', { scope: 'col', text: (text || '第 ' + (i + 1) + ' 個月') + ' 數量' }));
    });
    head.appendChild(h('th', { scope: 'col' }, h('span', { class: 'visually-hidden', text: '刪除' })));

    var body = h('tbody');
    S.items.forEach(function (it, r) {
      var n = r + 1;
      var name = h('input', {
        type: 'text', class: 'in entry-name', autocomplete: 'off', value: it.name,
        'aria-label': '第 ' + n + ' 個項目的名稱', placeholder: '項目名稱'
      });
      name.addEventListener('input', function () { it.name = name.value; onType(); });
      var ct = h('input', {
        type: 'text', class: 'in entry-num', inputmode: 'decimal', autocomplete: 'off', value: it.ct,
        'aria-label': '第 ' + n + ' 個項目的 CT.'
      });
      ct.addEventListener('input', function () { it.ct = ct.value; onType(); });
      var tr = h('tr', null, h('td', null, name), h('td', null, ct));
      it.qty.forEach(function (q, i) {
        var input = h('input', {
          type: 'text', class: 'in entry-num', inputmode: 'decimal', autocomplete: 'off', value: q,
          'aria-label': '第 ' + n + ' 個項目，' + (headers[i] || '第 ' + (i + 1) + ' 個月') + ' 的數量'
        });
        input.addEventListener('input', function () { it.qty[i] = input.value; onType(); });
        tr.appendChild(h('td', null, input));
      });
      tr.appendChild(h('td', null, h('button', {
        type: 'button', class: 'btn-link', text: '刪除', 'aria-label': '刪除第 ' + n + ' 個項目',
        onclick: function () { removeItem(r); }
      })));
      body.appendChild(tr);
    });

    els.entry.appendChild(h('div', { class: 'entry-scroll' },
      h('table', { class: 'entry' }, h('thead', null, head), body)));
  }

  // 把輸入框裡的文字整理成可以寫進 Excel 的項目；有問題時回傳 { error }
  function readItems() {
    var out = [];
    for (var r = 0; r < S.items.length; r++) {
      var it = S.items[r];
      var name = XIO.normalizeText(it.name);
      var touched = name !== '' || XIO.normalizeText(it.ct) !== '' ||
        it.qty.some(function (q) { return XIO.normalizeText(q) !== ''; });
      if (!touched) continue;                    // 整列都沒填，略過
      if (name === '') return { error: '第 ' + (r + 1) + ' 個項目還沒有輸入名稱。' };
      var ct = XIO.toNumber(it.ct);
      if (ct === null) return { error: '「' + name + '」還沒有輸入 CT.。' };
      if (isNaN(ct)) return { error: '「' + name + '」的 CT. 不是數字。' };
      var qty = [];
      for (var i = 0; i < it.qty.length; i++) {
        var q = XIO.toNumber(it.qty[i]);
        if (q !== null && isNaN(q)) return { error: '「' + name + '」第 ' + (i + 1) + ' 個月的數量不是數字。' };
        qty.push(q === null ? 0 : q);
      }
      out.push({ name: it.name.trim(), ct: ct, qty: qty });
    }
    return { items: out };
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = IEApp.state.withSummary;
    IEApp.state.withRma = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.result.textContent = '';

    if (!src) {
      els.srcInfo.textContent = '還沒有第 6 步的結果。回到第 6 步匯總，或上傳之前匯出的檔案。';
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

    var chain = chainStartRow();
    els.startRow.value = String(startRow());
    els.startRow.disabled = !!chain;
    els.startNote.textContent = chain ? '列開始（沿用第 1 步的設定）' : '列開始';

    // 還不知道有幾個月之前，先用 6 個月把輸入表畫出來
    var plan = null;
    var problem = '';
    var sr = startRow();
    var keyCol = src && src.kind === 'prev' && src.prev.keyCol ? src.prev.keyCol : 1;
    var base = { dataStartRow: sr, keyCol: keyCol, title: TITLE, summaryMark: SUMMARY_MARK };

    if (!src) {
      problem = '先準備好第 6 步的 Excel。';
    } else {
      ensureWork(src);
      if (S.work.loading) problem = '準備資料中…';
      else if (S.work.error) problem = S.work.error;
      else {
        plan = XIO.planRma(S.work.main, base);
        if (!plan.found) {
          problem = plan.reason === 'summary'
            ? 'MPS 工作表裡找不到匯總那三列。請先做第 6 步匯總。'
            : plan.reason === 'rows'
              ? '第 ' + sr + ' 列之後沒有料號。請確認上傳的檔案。'
              : 'MPS 工作表裡找不到「需求工時」。請先做第 5、6 步。';
          plan = null;
        }
      }
    }

    var months = plan ? plan.width : (S.items && S.items.length ? S.items[0].qty.length : DEFAULT_MONTHS);
    var headers = [];
    for (var i = 0; i < months; i++) headers.push(plan ? plan.headers[i] : '');
    initItems(months);
    renderEntry(headers);

    if (!plan) {
      var isError = !!src && !(S.work && S.work.loading);
      setSummary(problem, isError);
      return;
    }

    var read = readItems();
    if (read.error) { setSummary(read.error, true); return; }
    if (!read.items.length) { setSummary('輸入至少一個 RMA 項目。'); return; }

    var first = plan.startRow + 1;
    var last = plan.startRow + read.items.length;
    setSummary('RMA 會填在 ' + S.work.main.name + ' 工作表的第 ' + plan.startRow + ' 到 ' + last +
      ' 列（Others 總需求工時在第 ' + plan.restRow + ' 列，空一列），共 ' + read.items.length + ' 個項目。');

    renderResult(plan, read.items, first, last);
    els.exportBtn.disabled = false;

    var opts = Object.assign({ items: read.items, labels: LABELS }, base);
    IEApp.state.withRma = {
      outName: outputName(src),
      mainSheet: S.work.main.name,
      dataStartRow: sr,
      keyCol: keyCol,
      rmaRow: plan.startRow,
      rmaLastRow: last,
      build: function () { return buildWorkbook(src, opts); }
    };
    els.nextBtn.disabled = false;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  function showNumber(v) {
    return Math.round(v).toLocaleString('en-US');             // 和 Excel 一樣只顯示整數
  }

  // 算出來的 RMA 需求工時，以及第一個項目填入的公式
  function renderResult(plan, items, first, last) {
    var ctL = XIO.numToCol(plan.ctCol);
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      h('th', { class: 'colhead', scope: 'col', text: ctL + '　CT.' }),
      h('th', { class: 'colhead', scope: 'col', text: XIO.numToCol(plan.labelCol) + '　項目' }));
    plan.headers.forEach(function (text, i) {
      head.appendChild(h('th', { class: 'colhead', scope: 'col', text: XIO.numToCol(plan.startCol + i) + (text ? '　' + text : '') }));
    });

    var body = h('tbody');
    var totalRow = h('tr', null,
      h('th', { class: 'rowhead', scope: 'row', text: String(plan.startRow) }),
      h('td', { text: LABELS.ct }),
      h('td', { text: LABELS.title }));
    for (var i = 0; i < plan.width; i++) {
      var sum = 0;
      items.forEach(function (it) { sum += it.qty[i] * it.ct; });
      totalRow.appendChild(h('td', { class: 'is-num is-kept', text: showNumber(sum) }));
    }
    body.appendChild(totalRow);
    items.forEach(function (it, k) {
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(first + k) }),
        h('td', { class: 'is-num', text: String(it.ct) }),
        h('td', { text: it.name }));
      it.qty.forEach(function (q) { tr.appendChild(h('td', { class: 'is-num', text: showNumber(q * it.ct) })); });
      body.appendChild(tr);
    });

    els.result.appendChild(h('p', { class: 'sheet-caption', text: 'RMA 需求工時（數量 × CT.）。' }));
    els.result.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': 'RMA 需求工時' },
      h('table', { class: 'sheet is-result' }, h('thead', null, head), body)));

    var L = XIO.numToCol(plan.startCol);
    els.result.appendChild(h('p', { class: 'hint', text: L + ' 行填入的公式：' +
      L + plan.startRow + ' =SUM(' + L + first + ':' + L + last + ')　' +
      L + first + ' =' + String(items[0].qty[0]) + '*' + ctL + first }));
  }

  /* ---------- 產生活頁簿 ---------- */
  function buildWorkbook(src, opts) {
    return buildFrom(src).then(function (wb) {
      XIO.addRmaRows(pickMain(wb, src), opts);
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
    var ready = IEApp.state.withRma;
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

  IEApp.registerStep({ id: 'rma', render: render });
})(window);
