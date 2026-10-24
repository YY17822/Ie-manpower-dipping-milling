/*
 * 第 6 步：匯入人力報表
 *   把人力報表（未來六個月的人力預計）放進前面做好的 Excel，
 *   成為同一個檔案裡名為「人力」的工作表（整張照抄）。
 *
 *   匯入之後請使用者確認兩件事，之後做總表時才知道要去哪幾格抓資料：
 *     1. 這六個項目各在第幾列：
 *        Online、Online（生產）、Online（重工）、Off line（直接）、Off line（間接）、TTL（On + Off直接）
 *     2. 當月（第一個月）在哪一行。六個月就是從這一行開始往右連續六行。
 *
 *   程式會先依列的文字猜各項目在哪一列，使用者核對、不對就改。
 *   當月在哪一行：選好的各列裡的數字剛好落在六個連續的行時才會自動帶入，否則要自己選
 *   （報表如果連過去的月份也列出來，程式沒辦法知道哪一行是當月）。
 *   六個項目都選好時，各項目的列會記在這台電腦的瀏覽器裡，下個月列的文字對不上時拿來當預設值。
 *
 *   來源有兩種：
 *     a. 剛才第 5 步做好的結果（不用先匯出，直接沿用）
 *     b. 上傳之前第 5 步匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withManpower，下一步直接拿來用：
 *   rows           { online: 5, onlineProd: 6, … } 各項目在「人力」工作表的第幾列
 *   firstMonthCol  當月在第幾行（A=1）
 *   monthCount     6
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = '人力';
  var REF_SHEET = 'P003';
  var SUFFIX = '_人力';
  var MONTHS = 6;
  var STORE_KEY = 'step6';             // 共用設定的組名（IEApp.settings）
  var MIN_PREVIEW_ROWS = 15;
  var MAX_PREVIEW_ROWS = 200;
  var ITEMS = XIO.MANPOWER_ITEMS;

  var S = {
    base: null,           // 上傳的第 5 步檔案 { name, buffer, sheets, hasManpower }
    preferUpload: false,
    mp: null,             // 人力報表 { name, workbook, ws, ext, rows: [{ row, text, key }] }
    pick: null            // 使用者確認的位置 { rows: { online: 5, … }, firstCol: 5 }（還沒選的是 0）
  };
  var settings = loadSettings();
  var els = {};

  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    // lastFirstCol：上次確認的當月行，只給第 7 步（上傳第 6 步檔案時）用，這一步不拿來當預設值
    return { rows: s.rows && typeof s.rows === 'object' ? s.rows : {}, lastFirstCol: parseInt(s.lastFirstCol, 10) || 0 };
  }
  function saveSettings() {
    if (!S.pick) return;
    ITEMS.forEach(function (item) {
      if (S.pick.rows[item.key]) settings.rows[item.key] = S.pick.rows[item.key];
    });
    if (S.pick.firstCol) settings.lastFirstCol = S.pick.firstCol;
    try { IEApp.settings.write(STORE_KEY, settings); } catch (e) { /* 存不了就算了 */ }
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
    els = { rowSelects: {} };

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

    var grid = h('div', { class: 'field-grid' });
    ITEMS.forEach(function (item) {
      var id = 'f8-row-' + item.key;
      var sel = h('select', { id: id, class: 'in-select' });
      sel.addEventListener('change', function () { onRowChange(item.key, sel.value); });
      els.rowSelects[item.key] = sel;
      grid.appendChild(h('label', { for: id, text: item.label }));
      grid.appendChild(sel);
    });
    els.colSelect = h('select', { id: 'f8-col', class: 'in-select' });
    els.colSelect.addEventListener('change', onColChange);
    grid.appendChild(h('label', { for: 'f8-col', text: '當月（第一個月）' }));
    grid.appendChild(els.colSelect);

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
      h('p', { text: '把人力報表（未來六個月的人力預計）放進前面做好的 Excel，成為同一個檔案裡名為「人力」的工作表。匯入後確認各項目在哪一列、當月在哪一行，之後做總表時就照這些位置抓資料。' })));

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
        h('h2', { text: '確認各項目的位置' }),
        h('p', { class: 'hint', text: '選每個項目在人力報表的第幾列，以及當月（第一個月）在哪一行；六個月是從當月那一行開始往右連續六行。程式會先依列的文字帶入，請對照下面的表格確認。' }),
        grid, els.preview, els.check),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.notes,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.sheetSelect.addEventListener('change', function () {
      pickSheet(parseInt(els.sheetSelect.value, 10));
    });

    if (S.mp) showManpowerFile();
    fillRowSelects();
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

    // 預設值：先依列的文字猜；猜不到的，上次選的那一列如果還有內容就先帶入
    var guess = XIO.guessManpowerRows(mp.rows);
    var known = {};
    mp.rows.forEach(function (x) { known[x.row] = true; });
    var rows = {};
    ITEMS.forEach(function (item) {
      var remembered = parseInt(settings.rows[item.key], 10);
      rows[item.key] = guess[item.key] || (known[remembered] ? remembered : 0);
    });
    S.pick = { rows: rows, firstCol: 0 };
    guessColumn();

    fillRowSelects();
    refresh();
  }

  // 當月在哪一行還沒選時，用已經選好的列試著猜（各列的數字合起來剛好六個連續的行時才猜得出來）
  function guessColumn() {
    if (!S.mp || !S.pick || S.pick.firstCol) return;
    var rows = ITEMS.map(function (item) { return S.pick.rows[item.key]; }).filter(Boolean);
    if (rows.length) S.pick.firstCol = XIO.guessFirstMonthCol(S.mp.ws, rows, S.mp.ext.colCount, MONTHS);
  }

  function shorten(text, max) {
    return text.length > max ? text.slice(0, max - 1) + '…' : text;
  }

  function fillRowSelects() {
    ITEMS.forEach(function (item) {
      var sel = els.rowSelects[item.key];
      sel.textContent = '';
      if (!S.mp || !S.mp.ws) {
        sel.appendChild(h('option', { value: '', text: '先上傳人力報表' }));
        sel.disabled = true;
        return;
      }
      sel.disabled = false;
      sel.appendChild(h('option', { value: '', text: '請選擇' }));
      S.mp.rows.forEach(function (x) {
        sel.appendChild(h('option', {
          value: String(x.row),
          text: '第 ' + x.row + ' 列　' + (x.text ? shorten(x.text, 40) : '（這一列沒有文字）')
        }));
      });
      sel.value = S.pick.rows[item.key] ? String(S.pick.rows[item.key]) : '';
    });
  }

  function fillColSelect(labels) {
    var sel = els.colSelect;
    sel.textContent = '';
    if (!S.mp || !S.mp.ws) {
      sel.appendChild(h('option', { value: '', text: '先上傳人力報表' }));
      sel.disabled = true;
      return;
    }
    sel.disabled = false;
    sel.appendChild(h('option', { value: '', text: '請選擇' }));
    for (var c = 1; c <= S.mp.ext.colCount; c++) {
      sel.appendChild(h('option', {
        value: String(c),
        text: XIO.numToCol(c) + ' 行' + (labels[c] ? '　' + shorten(labels[c], 30) : '')
      }));
    }
    sel.value = S.pick.firstCol ? String(S.pick.firstCol) : '';
  }

  function onRowChange(key, value) {
    if (!S.pick) return;
    S.pick.rows[key] = parseInt(value, 10) || 0;
    guessColumn();
    els.exportNote.textContent = '';
    refresh();
  }

  function onColChange() {
    if (!S.pick) return;
    S.pick.firstCol = parseInt(els.colSelect.value, 10) || 0;
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

    // 選好的列、六個月的行
    var chosen = [];
    var missing = [];
    if (hasSheet) {
      ITEMS.forEach(function (item) {
        if (pick.rows[item.key]) chosen.push(pick.rows[item.key]); else missing.push(item.label);
      });
    }
    var monthCols = [];
    if (hasSheet && pick.firstCol) {
      for (var i = 0; i < MONTHS; i++) monthCols.push(pick.firstCol + i);
    }

    // 每一行的表頭文字：從最上面那個項目的列往上找
    var labels = [];
    if (hasSheet) {
      var topRow = chosen.length ? Math.min.apply(null, chosen) : (mp.rows.length ? mp.rows[0].row + 1 : 1);
      labels = XIO.labelsAbove(mp.ws, topRow, Math.max(mp.ext.colCount, pick.firstCol + MONTHS - 1));
    }
    fillColSelect(labels);

    // 人力報表預覽：選到的列和六個月的行畫螢光
    if (hasSheet) {
      var keep = {};
      monthCols.forEach(function (c) { keep[c] = true; });
      var keepRows = {};
      chosen.forEach(function (r) { keepRows[r] = true; });
      var show = Math.max(MIN_PREVIEW_ROWS, chosen.length ? Math.max.apply(null, chosen) + 2 : 0);
      ui.renderSheet(els.preview, mp.ws, {
        headerEnd: 0,
        lastRow: mp.ext.lastRow,
        colCount: mp.ext.colCount,
        rows: Math.min(show, MAX_PREVIEW_ROWS),
        keep: keep,
        keepRows: keepRows,
        caption: '人力報表。共 ' + mp.ext.lastRow.toLocaleString('en-US') + ' 列、' + mp.ext.colCount +
          ' 行。螢光的格子是之後總表要抓的資料。'
      });
    }

    // 核對表
    var read = null;
    if (hasSheet && monthCols.length && chosen.length) {
      read = XIO.readManpower(mp.ws, pick.rows, pick.firstCol, MONTHS);
      renderCheck(read, labels);
    }

    // 匯出
    if (!src) { setSummary('先準備好第 5 步的 Excel。'); return; }
    if (!mp || !mp.ws) { setSummary('上傳人力報表。'); return; }
    if (!hasSheet) { setSummary('人力報表的這個工作表沒有資料。請確認檔案，或換一個工作表。', true); return; }
    if (missing.length) { setSummary('還沒有選在哪一列：' + missing.join('、') + '。'); return; }
    if (!pick.firstCol) { setSummary('還沒有選當月（第一個月）在哪一行。'); return; }

    // 注意事項
    var warn = false;
    var byRow = {};
    ITEMS.forEach(function (item) {
      var r = pick.rows[item.key];
      (byRow[r] = byRow[r] || []).push(item.label);
    });
    Object.keys(byRow).forEach(function (r) {
      if (byRow[r].length < 2) return;
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' },
        byRow[r].join('、') + ' 選了同一列（第 ' + r + ' 列），請確認。'));
    });
    if (read.bad) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + read.bad + ' 格不是數字（上面表格標成淺紅色）。請確認列和行有沒有選對。'));
    }
    if (read.blank) {
      warn = true;
      els.notes.appendChild(h('li', { class: 'is-warn' },
        '有 ' + read.blank + ' 格是空白（上面表格標成淺紅色）。空白之後會當成 0；如果不該是空白，請確認列和行有沒有選對。'));
    }

    var firstL = XIO.numToCol(pick.firstCol);
    var lastL = XIO.numToCol(pick.firstCol + MONTHS - 1);
    var replacing = src.kind === 'file' && S.base.hasManpower;
    var text = '匯出的檔案會多一個「' + SHEET_NAME + '」工作表（' + mp.ext.lastRow.toLocaleString('en-US') + ' 列、' +
      mp.ext.colCount + ' 行）。當月是 ' + firstL + ' 行，六個月是 ' + firstL + ' 到 ' + lastL + ' 行。';
    if (replacing) text += ' 檔案裡原本的「' + SHEET_NAME + '」工作表會換成這次上傳的。';
    setSummary(text, false, warn || replacing);
    els.exportBtn.disabled = false;

    // 給下一步用
    var sheets = src.kind === 'prev'
      ? [prev.mainSheet, REF_SHEET]
      : S.base.sheets.filter(function (n) { return !sameName(n, SHEET_NAME); });
    sheets = sheets.concat([SHEET_NAME]);
    var rowsCopy = {};
    ITEMS.forEach(function (item) { rowsCopy[item.key] = pick.rows[item.key]; });
    var sheet = mp.ws;
    IEApp.state.withManpower = {
      outName: outputName(src),
      sheets: sheets,
      mainSheet: src.kind === 'prev' ? prev.mainSheet : S.base.sheets.filter(function (n) { return !isExtra(n); })[0],
      dataStartRow: src.kind === 'prev' ? prev.dataStartRow : 0,
      keyCol: src.kind === 'prev' && prev.keyCol ? prev.keyCol : 1,
      manpowerSheet: SHEET_NAME,
      rows: rowsCopy,
      firstMonthCol: pick.firstCol,
      monthCount: MONTHS,
      build: function () { return buildWorkbook(src, sheet); }
    };
    saveSettings();       // 六個項目都選好了，記下來給下個月當預設值
    els.nextBtn.disabled = false;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  // 核對表：六個項目 × 六個月，照選的列和行從人力報表讀出來的數值
  function renderCheck(read, labels) {
    var pick = S.pick;
    var head = h('tr', null,
      h('th', { class: 'colhead', scope: 'col', text: '項目' }),
      h('th', { class: 'colhead', scope: 'col', text: '列' }));
    read.cols.forEach(function (c, i) {
      head.appendChild(h('th', { class: 'colhead', scope: 'col',
        text: XIO.numToCol(c) + (labels[c] ? '　' + labels[c] : '') + (i === 0 ? '（當月）' : '') }));
    });

    var body = h('tbody');
    var example = '';
    ITEMS.forEach(function (item) {
      var r = pick.rows[item.key];
      var tr = h('tr', null,
        h('th', { class: 'rowhead', scope: 'row', text: item.label }),
        h('td', { text: r ? '第 ' + r + ' 列' : '還沒選' }));
      read.values[item.key].forEach(function (info) {
        var off = !r || info.blank || info.bad;
        tr.appendChild(h('td', {
          class: (off ? 'is-missing' : 'is-num'),
          text: !r ? '' : info.blank ? '（空白）' : info.text
        }));
      });
      body.appendChild(tr);
      if (r && !example) example = "='" + SHEET_NAME + "'!" + XIO.numToCol(read.cols[0]) + r;
    });

    els.check.appendChild(h('p', { class: 'sheet-caption', text: '照上面選的位置，之後總表會抓到的數值：' }));
    els.check.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '各項目六個月的人力' },
      h('table', { class: 'sheet is-result is-check' }, h('thead', null, head), body)));
    if (example) {
      els.check.appendChild(h('p', { class: 'hint',
        text: '之後總表會用公式抓這些格子，例如 ' + example + '，點儲存格就能看到是從人力報表哪一格來的。' }));
    }
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
