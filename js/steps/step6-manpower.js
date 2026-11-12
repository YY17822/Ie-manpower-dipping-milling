/*
 * 第 6 步：匯入人力報表（Dipping_Milling 現有人力表）
 *   把人力報表放進前面做好的 Excel，成為同一個檔案裡名為「人力」的工作表（整張照抄）。
 *
 *   人力報表的樣子（例如「Dipping / Milling 現有人力表」）：
 *     一行是項目（操機、鏡檢、支援入…），旁邊是 Dipping、Milling 兩行的人數，再過去是「總現有人力」；
 *     項目下面有一列 TTL（各行的合計）。只有一組人數（這個月的人數），總表六個月都用同一個值。
 *
 *   匯入之後請使用者確認：
 *     1. 項目、Dipping、Milling、總現有人力各在哪一行
 *     2. 項目從第幾列到第幾列、TTL 在第幾列
 *   程式會先依格子的文字猜（Dipping、Mill…、總現有人力、TTL），使用者核對、不對就改。
 *   確認的位置會記住（共用設定 step6），下個月猜不到時拿來當預設值；第 7 步上傳第 6 步的檔案時也用它找資料。
 *
 *   之後總表：
 *     人力表整張放到總表 O 行起（最左邊有內容的那一行對到 O，列不變），原本的「人力」工作表不保留
 *     現有人力 Milling   =<人力表在總表裡的 Milling 行><TTL 列>，例如 =Q12（六個月相同），Notes 寫 Milling 各項目的人數，例如「操機 3 人、鏡檢 3 人」
 *     現有人力 Dipping   同上，用 Dipping 那一行
 *
 *   來源有兩種：
 *     a. 剛才第 5 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 5 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withManpower，下一步直接拿來用：
 *   manpowerSheet  '人力'
 *   layout         { itemCol, dipCol, millCol, totalCol, firstItemRow, lastItemRow, ttlRow }（數字，A=1）
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = XIO.DM_MANPOWER.sheet;    // '人力'
  var REF_SHEET = 'P003';
  var SUFFIX = '_人力';
  var STORE_KEY = 'step6';             // 共用設定的組名（IEApp.settings）
  var MIN_PREVIEW_ROWS = 15;
  var MAX_PREVIEW_ROWS = 200;
  var COLS = XIO.DM_MANPOWER.cols;     // [{ key: 'itemCol', label: '項目' }, …]
  var ROWS = XIO.DM_MANPOWER.rows;     // [{ key: 'firstItemRow', label: '項目從第' }, …]
  var SIDES = XIO.DM_MANPOWER.sides;   // [{ key: 'milling', label: 'Milling', col: 'millCol' }, …]

  var S = {
    base: null,           // 上傳的第 5 步檔案 { name, buffer, sheets, hasManpower }
    preferUpload: false,
    mp: null,             // 人力報表 { name, workbook, ws, ext, rows: [{ row, text }] }
    pick: null            // 使用者確認的位置 { itemCol, dipCol, millCol, totalCol, firstItemRow, lastItemRow, ttlRow }（還沒選的是 0）
  };
  var settings = loadSettings();
  var els = {};

  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    return XIO.dmManpowerLayout(s);
  }
  // 記住：行存成字母（B、C…），列存成數字
  function saveSettings() {
    if (!S.pick) return;
    var out = {};
    COLS.forEach(function (x) { out[x.key] = S.pick[x.key] ? XIO.numToCol(S.pick[x.key]) : ''; });
    ROWS.forEach(function (x) { out[x.key] = S.pick[x.key] || ''; });
    settings = XIO.dmManpowerLayout(out);
    try { IEApp.settings.write(STORE_KEY, out); } catch (e) { /* 存不了就算了 */ }
  }

  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  function isExtra(name) { return sameName(name, REF_SHEET) || sameName(name, SHEET_NAME); }

  // 前面步驟的結果：第 5 步（計算需求工時）
  function previous() {
    return IEApp.state.withDemand || null;
  }
  function previousLabel() {
    return '第 5 步剛做好的結果';
  }

  function activeSource() {
    var prev = previous();
    if (S.base && (S.preferUpload || !prev)) return { kind: 'file', base: S.base, name: S.base.name };
    if (prev) return { kind: 'prev', prev: prev, name: prev.outName };
    return null;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = { colSelects: {}, rowSelects: {} };

    els.baseInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.baseDrop = ui.fileDrop({
      id: 'f8-base',
      title: '上傳第 5 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openBase
    });
    els.baseSwitch = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 5 步計算需求工時',
      onclick: function () { IEApp.go('demand-hours'); }
    });
    els.baseMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.mpDrop = ui.fileDrop({
      id: 'f8-file',
      title: '選擇人力報表的 Excel 檔',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openManpower
    });
    els.sheetSelect = h('select', { id: 'f8-sheet', 'aria-label': '人力報表的工作表' });
    els.sheetRow = h('div', { class: 'field-row', hidden: true },
      h('label', { for: 'f8-sheet', text: '工作表' }), els.sheetSelect);
    els.mpMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    var colGrid = h('div', { class: 'field-grid' });
    COLS.forEach(function (x) {
      var id = 'f8-' + x.key;
      var sel = h('select', { id: id, class: 'in-select' });
      sel.addEventListener('change', function () { onPick(x.key, sel.value); });
      els.colSelects[x.key] = sel;
      colGrid.appendChild(h('label', { for: id, text: x.label + ' 在' }));
      colGrid.appendChild(sel);
    });
    var rowGrid = h('div', { class: 'field-grid' });
    ROWS.forEach(function (x) {
      var id = 'f8-' + x.key;
      var sel = h('select', { id: id, class: 'in-select' });
      sel.addEventListener('change', function () { onPick(x.key, sel.value); });
      els.rowSelects[x.key] = sel;
      rowGrid.appendChild(h('label', { for: id, text: x.label }));
      rowGrid.appendChild(sel);
    });

    els.preview = h('div', { class: 'sheet-area' });
    els.check = h('div', { class: 'sheet-area' });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.notes = h('ul', { class: 'notes' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：總表',
      onclick: function () { IEApp.go('report'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '匯入人力報表' }),
      h('p', { text: '把現有人力表放進前面做好的 Excel，成為同一個檔案裡名為「人力」的工作表。匯入後確認項目、Dipping、Milling、總現有人力各在哪一行，項目和 TTL 在哪幾列。總表的現有人力 Milling、Dipping 會抓該行的 TTL（六個月都用同一個值），Notes 寫各項目的人數。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 5 步的 Excel' }),
        els.baseInfo, els.baseDrop.el,
        h('div', { class: 'link-row' }, els.baseSwitch, els.backBtn),
        els.baseMsg),
      h('li', { class: 'task' },
        h('h2', { text: '上傳人力報表' }),
        els.mpDrop.el, els.sheetRow, els.mpMsg),
      h('li', { class: 'task' },
        h('h2', { text: '確認各欄、各列的位置' }),
        h('p', { class: 'hint', text: '程式會先依格子的文字帶入（Dipping、Milling、總現有人力、TTL），請對照下面的表格確認，不對就改。確認的位置會記住，下個月照樣用。' }),
        colGrid, rowGrid, els.preview, els.check),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.sheetSelect.addEventListener('change', function () {
      pickSheet(parseInt(els.sheetSelect.value, 10));
    });

    if (S.mp) showManpowerFile();
    refresh();
  }

  /* ---------- 第 5 步的 Excel ---------- */
  function openBase(file) {
    els.baseMsg.hidden = true;
    els.baseDrop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (ws) { return ws.name; });
        var others = names.filter(function (n) { return !isExtra(n); });
        if (!others.length) throw new Error('這個檔案沒有 MPS 工作表，看起來不是第 5 步匯出的檔案。請上傳第 5 步匯出的 Excel。');
        S.base = {
          name: file.name, buffer: res.buffer, sheets: names,
          hasManpower: names.some(function (n) { return sameName(n, SHEET_NAME); })
        };
        S.preferUpload = true;
        els.exportNote.textContent = '';
        refresh();
      })
      .catch(function (err) {
        els.baseMsg.textContent = err && err.message ? err.message : String(err);
        els.baseMsg.hidden = false;
        refresh();
      });
  }

  function switchSource() {
    S.preferUpload = !S.preferUpload;
    els.exportNote.textContent = '';
    refresh();
  }

  /* ---------- 人力報表 ---------- */
  function openManpower(file) {
    els.mpMsg.hidden = true;
    els.mpDrop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        S.mp = { name: file.name, workbook: res.workbook, ws: null, ext: null, rows: [] };
        showManpowerFile();
        // 檔案裡有叫「人力」的工作表就先選那一張，沒有就選第一張
        var sheets = res.workbook.worksheets;
        var named = sheets.filter(function (ws) { return sameName(ws.name, SHEET_NAME); })[0];
        pickSheet(named ? sheets.indexOf(named) : 0);
      })
      .catch(function (err) {
        els.mpMsg.textContent = err && err.message ? err.message : String(err);
        els.mpMsg.hidden = false;
        if (S.mp) showManpowerFile(); else els.mpDrop.setIdle();
        refresh();
      });
  }

  function showManpowerFile() {
    els.mpDrop.setFile(S.mp.name);
    var sheets = S.mp.workbook.worksheets;
    els.sheetSelect.textContent = '';
    sheets.forEach(function (ws, i) {
      els.sheetSelect.appendChild(h('option', { value: String(i), text: ws.name }));
    });
    els.sheetRow.hidden = sheets.length < 2;
    if (S.mp.ws) els.sheetSelect.value = String(sheets.indexOf(S.mp.ws));
  }

  function pickSheet(index) {
    var mp = S.mp;
    mp.ws = mp.workbook.worksheets[index];
    mp.ext = XIO.extent(mp.ws);
    mp.rows = XIO.manpowerRows(mp.ws, mp.ext);
    els.sheetSelect.value = String(index);
    els.exportNote.textContent = '';

    // 預設值：先依文字猜；猜不到的，用上次記住的位置（還在工作表範圍內才用）
    var guess = XIO.guessDmManpower(mp.ws, mp.ext);
    var pick = {};
    COLS.forEach(function (x) {
      var saved = settings[x.key];
      pick[x.key] = guess[x.key] || (saved && saved <= mp.ext.colCount ? saved : 0);
    });
    ROWS.forEach(function (x) {
      var saved = settings[x.key];
      pick[x.key] = guess[x.key] || (saved && saved <= mp.ext.lastRow ? saved : 0);
    });
    S.pick = pick;
    refresh();
  }

  function shorten(text, max) {
    return text.length > max ? text.slice(0, max - 1) + '…' : text;
  }

  function fillSelects(labels) {
    var ready = !!(S.mp && S.mp.ws);
    COLS.forEach(function (x) {
      var sel = els.colSelects[x.key];
      sel.textContent = '';
      if (!ready) { sel.appendChild(h('option', { value: '', text: '先上傳人力報表' })); sel.disabled = true; return; }
      sel.disabled = false;
      sel.appendChild(h('option', { value: '', text: '請選擇' }));
      for (var c = 1; c <= S.mp.ext.colCount; c++) {
        sel.appendChild(h('option', {
          value: String(c),
          text: XIO.numToCol(c) + ' 行' + (labels[c] ? '　' + shorten(labels[c], 30) : '')
        }));
      }
      sel.value = S.pick[x.key] ? String(S.pick[x.key]) : '';
    });
    ROWS.forEach(function (x) {
      var sel = els.rowSelects[x.key];
      sel.textContent = '';
      if (!ready) { sel.appendChild(h('option', { value: '', text: '先上傳人力報表' })); sel.disabled = true; return; }
      sel.disabled = false;
      sel.appendChild(h('option', { value: '', text: '請選擇' }));
      S.mp.rows.forEach(function (rr) {
        sel.appendChild(h('option', {
          value: String(rr.row),
          text: '第 ' + rr.row + ' 列　' + (rr.text ? shorten(rr.text, 40) : '（這一列沒有文字）')
        }));
      });
      // 選到的列如果不在清單裡（例如整列空白），也加進去
      var v = S.pick[x.key];
      if (v && !S.mp.rows.some(function (rr) { return rr.row === v; })) {
        sel.appendChild(h('option', { value: String(v), text: '第 ' + v + ' 列　（空白）' }));
      }
      sel.value = v ? String(v) : '';
    });
  }

  function onPick(key, value) {
    if (!S.pick) return;
    S.pick[key] = parseInt(value, 10) || 0;
    els.exportNote.textContent = '';
    refresh();
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = previous();
    IEApp.state.withManpower = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    els.notes.textContent = '';
    els.preview.textContent = '';
    els.check.textContent = '';

    // 第 5 步的 Excel
    if (!src) {
      els.baseInfo.textContent = '還沒有第 5 步的結果。回到第 5 步，或上傳之前匯出的檔案。';
      els.baseDrop.setIdle();
    } else if (src.kind === 'prev') {
      els.baseInfo.textContent = '沿用' + previousLabel() + '：' + prev.outName + '。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.base) els.baseDrop.setFile(S.base.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.baseDrop.setIdle();
    } else {
      els.baseInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.base.sheets.join('、') + '。';
      els.baseDrop.setFile(S.base.name);
    }
    els.backBtn.hidden = !!src;
    els.baseSwitch.hidden = !(S.base && prev);
    if (src) els.baseSwitch.textContent = src.kind === 'file' ? '改用' + previousLabel() : '改用上傳的檔案';

    var mp = S.mp;
    var hasSheet = !!(mp && mp.ws && mp.ext.lastRow);
    var pick = S.pick;

    // 每一行的表頭文字：從第一個項目的列往上找
    var labels = [];
    if (hasSheet) {
      var topRow = pick.firstItemRow || pick.ttlRow || mp.ext.lastRow + 1;
      labels = XIO.labelsAbove(mp.ws, topRow, mp.ext.colCount);
    }
    fillSelects(labels);

    // 人力報表預覽：選到的行、項目列和 TTL 列畫螢光
    if (hasSheet) {
      var keep = {};
      COLS.forEach(function (x) { if (pick[x.key]) keep[pick[x.key]] = true; });
      var keepRows = {};
      if (pick.firstItemRow && pick.lastItemRow) {
        for (var r = pick.firstItemRow; r <= pick.lastItemRow; r++) keepRows[r] = true;
      }
      if (pick.ttlRow) keepRows[pick.ttlRow] = true;
      var show = Math.max(MIN_PREVIEW_ROWS, (pick.ttlRow || pick.lastItemRow || 0) + 2);
      ui.renderSheet(els.preview, mp.ws, {
        headerEnd: 0,
        lastRow: mp.ext.lastRow,
        colCount: mp.ext.colCount,
        rows: Math.min(show, MAX_PREVIEW_ROWS),
        keep: keep,
        keepRows: keepRows,
        caption: '人力報表。共 ' + mp.ext.lastRow.toLocaleString('en-US') + ' 列、' + mp.ext.colCount +
          ' 行。螢光的是選到的行和列。'
      });
    }

    // 匯出前的檢查
    if (!src) { setSummary('先準備好第 5 步的 Excel。'); return; }
    if (!mp || !mp.ws) { setSummary('上傳人力報表。'); return; }
    if (!hasSheet) { setSummary('人力報表的這個工作表沒有資料。請確認檔案，或換一個工作表。', true); return; }
    var missing = COLS.concat(ROWS).filter(function (x) { return !pick[x.key]; })
      .map(function (x) { return x.name; });
    if (missing.length) { setSummary('還沒有選：' + missing.join('、') + '。'); return; }
    if (pick.firstItemRow > pick.lastItemRow) {
      setSummary('項目的第一列（第 ' + pick.firstItemRow + ' 列）在最後一列（第 ' + pick.lastItemRow + ' 列）下面了，請確認。', true);
      return;
    }
    if (pick.dipCol === pick.millCol) {
      setSummary('Dipping 和 Milling 選到同一行了，請確認。', true);
      return;
    }

    var read = XIO.readDmManpower(mp.ws, pick);
    var offset = XIO.dmManpowerOffset(mp.ws);    // 第 7 步把人力表放進總表 O 行起要往右移幾行
    var have = XIO.dmHaveInfo(read, XIO.shiftDmLayout(pick, offset), null);
    renderCheck(read, have);

    // 注意事項
    var warn = false;
    function note(text, isWarn) {
      if (isWarn) warn = true;
      els.notes.appendChild(h('li', { class: isWarn ? 'is-warn' : null }, text));
    }
    COLS.forEach(function (x, i) {
      COLS.slice(i + 1).forEach(function (y) {
        if (pick[x.key] === pick[y.key]) note(x.label + ' 和 ' + y.label + ' 選了同一行（' + XIO.numToCol(pick[x.key]) + ' 行），請確認。', true);
      });
    });
    if (pick.ttlRow >= pick.firstItemRow && pick.ttlRow <= pick.lastItemRow) {
      note('TTL（第 ' + pick.ttlRow + ' 列）在項目的範圍裡，請確認項目到第幾列。', true);
    }
    if (!read.items.length) note('第 ' + pick.firstItemRow + ' 到 ' + pick.lastItemRow + ' 列沒有任何項目，Notes 會是空白。', true);
    SIDES.forEach(function (sd) {
      var t = read.ttl[sd.key];
      var at = XIO.numToCol(pick[sd.col]) + pick.ttlRow;
      if (t.blank) note(sd.label + ' 的 TTL（' + at + '）是空白，總表會當成 0。', true);
      else if (t.bad) note(sd.label + ' 的 TTL（' + at + '）不是數字：「' + t.text + '」。', true);
      else if (!near(read.sum[sd.key], t.number)) {
        note(sd.label + ' 各項目合計 ' + fmt(read.sum[sd.key]) + ' 人，TTL 是 ' + fmt(t.number) + ' 人，對不起來。請確認項目的列有沒有選完整。', true);
      }
      var bad = read.items.filter(function (it) { return it[sd.key].bad; }).map(function (it) { return it.name; });
      if (bad.length) note(sd.label + ' 有不是數字的格子：' + bad.join('、') + '，Notes 不會列出。', true);
    });
    var tot = read.ttl.total;
    var both = (read.ttl.milling.number || 0) + (read.ttl.dipping.number || 0);
    if (tot.blank || tot.bad) {
      note('總現有人力的 TTL（' + XIO.numToCol(pick.totalCol) + pick.ttlRow + '）是空白或不是數字，請確認。', true);
    } else if (!near(tot.number, both)) {
      note('總現有人力 TTL 是 ' + fmt(tot.number) + ' 人，但 Milling + Dipping 是 ' + fmt(both) + ' 人，對不起來，請確認。', true);
    }

    var replacing = src.kind === 'file' && S.base.hasManpower;
    var text = '匯出的檔案會多一個「' + SHEET_NAME + '」工作表（' + mp.ext.lastRow.toLocaleString('en-US') + ' 列、' +
      mp.ext.colCount + ' 行）；第 7 步做總表時會把它整張放到總表的 ' + XIO.numToCol(XIO.DM_MANPOWER.reportCol) +
      ' 行起（列不變），原本的「' + SHEET_NAME + '」工作表不保留。總表的現有人力 Milling 會抓 =' + have.milling.formula + '（' + fmt(have.milling.value) + ' 人）、' +
      'Dipping 抓 =' + have.dipping.formula + '（' + fmt(have.dipping.value) + ' 人），六個月都用同一個值。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會換成這次上傳的。';
    setSummary(text, false, warn || replacing);
    els.exportBtn.disabled = false;

    // 給下一步用
    var sheets = src.kind === 'prev'
      ? [prev.mainSheet, REF_SHEET]
      : S.base.sheets.filter(function (n) { return !sameName(n, SHEET_NAME); });
    sheets = sheets.concat([SHEET_NAME]);
    var layout = {};
    COLS.concat(ROWS).forEach(function (x) { layout[x.key] = pick[x.key]; });
    var sheet = mp.ws;
    IEApp.state.withManpower = {
      outName: outputName(src),
      sheets: sheets,
      mainSheet: src.kind === 'prev' ? prev.mainSheet : S.base.sheets.filter(function (n) { return !isExtra(n); })[0],
      dataStartRow: src.kind === 'prev' ? prev.dataStartRow : 0,
      keyCol: src.kind === 'prev' && prev.keyCol ? prev.keyCol : 1,
      manpowerSheet: SHEET_NAME,
      layout: layout,
      build: function () { return buildWorkbook(src, sheet); }
    };
    saveSettings();       // 全部選好了，記下來給下個月和第 7 步用
    els.nextBtn.disabled = false;
  }

  function near(a, b) { return Math.abs((a || 0) - (b || 0)) < 1e-9; }
  function fmt(v) { return typeof v === 'number' && !isNaN(v) ? String(Math.round(v * 1e6) / 1e6) : '0'; }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  // 核對表：每個項目一列（Milling、Dipping 的人數），最後是 TTL；下面寫總表會填的公式和 Notes
  function renderCheck(read, have) {
    var pick = S.pick;
    function colHead(text, c) {
      return h('th', { class: 'colhead', scope: 'col', text: text + (c ? '（' + XIO.numToCol(c) + ' 行）' : '') });
    }
    function numCell(info) {
      return h('td', { class: info.bad ? 'is-missing' : 'is-num', text: info.blank ? '' : info.text });
    }
    function ttlCell(info) {
      var off = info.blank || info.bad;
      return h('td', { class: (off ? 'is-missing' : 'is-num') + ' is-kept', text: info.blank ? '（空白）' : info.text });
    }
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '列' }),
      colHead('項目', pick.itemCol),
      colHead('Milling', pick.millCol),
      colHead('Dipping', pick.dipCol),
      colHead('總現有人力', pick.totalCol));
    var body = h('tbody');
    read.items.forEach(function (it) {
      body.appendChild(h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: String(it.row) }),
        h('td', { text: it.name }),
        numCell(it.milling), numCell(it.dipping), h('td')));
    });
    var t = read.ttl;
    body.appendChild(h('tr', null,
      h('th', { class: 'rowhead', scope: 'row', text: String(pick.ttlRow) }),
      h('td', { text: 'TTL' }),
      ttlCell(t.milling), ttlCell(t.dipping), ttlCell(t.total)));

    els.check.appendChild(h('p', { class: 'sheet-caption', text: '照上面選的位置讀到的人數（螢光是 TTL）：' }));
    els.check.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '現有人力' },
      h('table', { class: 'sheet is-result is-check' }, h('thead', null, head), body)));
    SIDES.forEach(function (sd) {
      var x = have[sd.key];
      els.check.appendChild(h('p', { class: 'hint',
        text: '總表 現有人力 ' + sd.label + '：=' + x.formula + '（六個月相同）　Notes：' + (x.note || '（沒有人數，留空白）') }));
    });
  }

  /* ---------- 產生活頁簿 ---------- */
  // 每次都從頭產生一份新的：前面步驟的內容 + 人力工作表。回傳 Promise<Workbook>
  function buildWorkbook(src, sheet) {
    var base = src.kind === 'prev'
      ? Promise.resolve(src.prev.build())
      : XIO.load(src.base.buffer.slice(0));
    return base.then(function (wb) {
      XIO.copySheet(sheet, wb, SHEET_NAME);
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
    var ready = IEApp.state.withManpower;
    if (!ready) return;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    ready.build()
      .then(function (wb) { return XIO.toBlob(wb); })
      .then(function (blob) {
        ui.download(blob, ready.outName);
        els.exportNote.textContent = '已匯出 ' + ready.outName + '（工作表：' + ready.sheets.join('、') + '），在瀏覽器的下載資料夾。';
      })
      .catch(function (err) {
        els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
      })
      .then(function () {
        els.exportBtn.disabled = false;
      });
  }

  IEApp.registerStep({ id: 'manpower', render: render });
})(window);
