/*
 * 第 3 步：匯入 P003 標準工時
 *   把 P003 報表（公司所有產品型號的標準工時）放進前面做好的 MPS Excel，
 *   成為同一個檔案裡名為「P003」的工作表。
 *
 *   MPS 的 Excel 有兩種來源：
 *     a. 剛才第 2 步（匯總）做好的結果，不用先匯出，直接沿用。
 *        如果跳過第 2 步，就沿用第 1 步篩選好的結果。
 *     b. 上傳之前第 2 步（或第 1 步）匯出的檔案
 *
 * 做完之後把結果放在 IEApp.state.withP003，下一步直接拿來用。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var SHEET_NAME = 'P003';
  var SUFFIX = '_含P003';
  var PREVIEW_ROWS = 8;

  // 這一步的狀態（切到別的步驟再回來時還在）
  var S = {
    base: null,          // 上傳的 MPS 檔案 { name, buffer, sheets: [工作表名稱], hasP003 }
    preferUpload: false, // 兩種來源都有時，用哪一個
    p003: null           // { name, workbook, ws, ext: { lastRow, colCount } }
  };
  var els = {};

  function isP003(name) { return String(name).toLowerCase() === SHEET_NAME.toLowerCase(); }

  // 前面步驟的結果：做過第 2 步（匯總）就用第 2 步的，沒有就用第 1 步的
  function previous() {
    return IEApp.state.mpsTotals || IEApp.state.mpsFiltered || null;
  }
  function previousLabel() {
    return IEApp.state.mpsTotals ? '第 2 步剛做好的結果' : '第 1 步剛篩選好的結果（沒有做第 2 步匯總）';
  }

  // 目前要用哪一個來源當作 MPS 的 Excel
  function activeSource() {
    var prev = previous();
    if (S.base && (S.preferUpload || !prev)) return { kind: 'file', base: S.base };
    if (prev) return { kind: 'prev', prev: prev };
    return null;
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.baseInfo = h('p', { class: 'source-line', 'aria-live': 'polite' });
    els.baseDrop = ui.fileDrop({
      id: 'f2-base',
      title: '上傳第 2 步匯出的 Excel',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openBase
    });
    els.baseSwitch = h('button', { type: 'button', class: 'btn-link', hidden: true, onclick: switchSource });
    els.backBtn = h('button', {
      type: 'button', class: 'btn-link', hidden: true, text: '回到第 2 步匯總',
      onclick: function () { IEApp.go('mps-totals'); }
    });
    els.baseMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    els.p003Drop = ui.fileDrop({
      id: 'f2-p003',
      title: '選擇 P003 報表的 Excel 檔',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openP003
    });
    els.sheetSelect = h('select', { id: 'f2-sheet', 'aria-label': 'P003 的工作表' });
    els.sheetRow = h('div', { class: 'field-row', hidden: true },
      h('label', { for: 'f2-sheet', text: '工作表' }), els.sheetSelect);
    els.p003Msg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });
    els.preview = h('div', { class: 'sheet-area' });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：帶入 CT.',
      onclick: function () { IEApp.go('ct-lookup'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '匯入 P003 標準工時' }),
      h('p', { text: '把 P003 報表（各產品型號的標準工時）放進前面做好的 MPS Excel，成為同一個檔案裡名為「P003」的工作表。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '第 2 步的 Excel' }),
        els.baseInfo, els.baseDrop.el,
        h('div', { class: 'link-row' }, els.baseSwitch, els.backBtn),
        els.baseMsg),
      h('li', { class: 'task' },
        h('h2', { text: '上傳 P003 報表' }),
        els.p003Drop.el, els.sheetRow, els.p003Msg, els.preview),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.sheetSelect.addEventListener('change', function () {
      pickSheet(parseInt(els.sheetSelect.value, 10));
    });

    if (S.p003) showP003File();
    refresh();
  }

  /* ---------- MPS 的 Excel ---------- */
  function openBase(file) {
    els.baseMsg.hidden = true;
    els.baseDrop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        var names = res.workbook.worksheets.map(function (ws) { return ws.name; });
        var others = names.filter(function (n) { return !isP003(n); });
        if (!others.length) {
          var e = new Error('這個檔案只有 P003 工作表，看起來不是第 2 步匯出的檔案。請上傳第 2 步匯出的 Excel。');
          throw e;
        }
        S.base = { name: file.name, buffer: res.buffer, sheets: names, hasP003: others.length < names.length };
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

  /* ---------- P003 報表 ---------- */
  function openP003(file) {
    els.p003Msg.hidden = true;
    els.p003Drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        S.p003 = { name: file.name, workbook: res.workbook, ws: null, ext: null };
        showP003File();
        pickSheet(0);
      })
      .catch(function (err) {
        els.p003Msg.textContent = err && err.message ? err.message : String(err);
        els.p003Msg.hidden = false;
        if (S.p003) showP003File(); else els.p003Drop.setIdle();
        refresh();
      });
  }

  function showP003File() {
    els.p003Drop.setFile(S.p003.name);
    var sheets = S.p003.workbook.worksheets;
    els.sheetSelect.textContent = '';
    sheets.forEach(function (ws, i) {
      els.sheetSelect.appendChild(h('option', { value: String(i), text: ws.name }));
    });
    els.sheetRow.hidden = sheets.length < 2;
    if (S.p003.ws) els.sheetSelect.value = String(sheets.indexOf(S.p003.ws));
  }

  function pickSheet(index) {
    S.p003.ws = S.p003.workbook.worksheets[index];
    S.p003.ext = XIO.extent(S.p003.ws);
    els.sheetSelect.value = String(index);
    els.exportNote.textContent = '';
    refresh();
  }

  /* ---------- 更新畫面 ---------- */
  function refresh() {
    var src = activeSource();
    var prev = previous();
    IEApp.state.withP003 = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;

    // MPS 的 Excel
    if (!src) {
      els.baseInfo.textContent = '還沒有第 2 步的結果。回到第 2 步匯總，或上傳之前匯出的檔案。';
      els.baseDrop.setIdle();
    } else if (src.kind === 'prev') {
      els.baseInfo.textContent = '沿用' + previousLabel() + '：' + prev.outName + '（' +
        prev.matchedRows.toLocaleString('en-US') + ' 列資料）。不用先匯出；要改用之前匯出的檔案，再從下面上傳。';
      if (S.base) els.baseDrop.setFile(S.base.name, '已上傳，目前沒有使用。點這裡換一個檔案');
      else els.baseDrop.setIdle();
    } else {
      els.baseInfo.textContent = '使用上傳的檔案，裡面的工作表：' + S.base.sheets.join('、') + '。';
      els.baseDrop.setFile(S.base.name);
    }
    els.backBtn.hidden = !!src;
    els.baseSwitch.hidden = !(S.base && prev);
    if (src) {
      els.baseSwitch.textContent = src.kind === 'file' ? '改用' + previousLabel() : '改用上傳的檔案';
    }

    // P003 預覽
    els.preview.textContent = '';
    var p = S.p003;
    if (p && p.ws && p.ext.lastRow) {
      ui.renderSheet(els.preview, p.ws, {
        headerEnd: 0,
        lastRow: p.ext.lastRow,
        colCount: p.ext.colCount,
        rows: PREVIEW_ROWS,
        caption: 'P003 報表（前 ' + PREVIEW_ROWS + ' 列）。共 ' + p.ext.lastRow.toLocaleString('en-US') +
          ' 列、' + p.ext.colCount + ' 行。'
      });
    }

    // 匯出
    if (!src) { setSummary('先準備好第 2 步的 Excel。'); return; }
    if (!p || !p.ws) { setSummary('上傳 P003 報表。'); return; }
    if (!p.ext.lastRow) { setSummary('P003 的這個工作表沒有資料。請確認檔案，或換一個工作表。', true); return; }

    var sheets = src.kind === 'prev'
      ? [prev.sheetName]
      : S.base.sheets.filter(function (n) { return !isP003(n); });
    sheets = sheets.concat([SHEET_NAME]);
    var text = '匯出的檔案有 ' + sheets.length + ' 個工作表：' + sheets.join('、') + '。P003 共 ' +
      p.ext.lastRow.toLocaleString('en-US') + ' 列、' + p.ext.colCount + ' 行。';
    var replacing = src.kind === 'file' && S.base.hasP003;
    if (replacing) text += ' 檔案裡原本的 P003 工作表會換成這次上傳的。';
    setSummary(text, false, replacing);
    els.exportBtn.disabled = false;

    // 給下一步用
    IEApp.state.withP003 = {
      outName: outputName(src),
      sheets: sheets,
      mainSheet: sheets[0],                                        // MPS 那張工作表的名稱
      dataStartRow: src.kind === 'prev' ? prev.headerEnd + 1 : 0, // MPS 資料從第幾列開始（上傳的檔案不知道，給 0）
      build: function () { return buildWorkbook(src, p.ws); }
    };
    els.nextBtn.disabled = false;
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 產生活頁簿 ---------- */
  // 每次都從頭產生一份新的：前面步驟的內容 + P003 工作表。回傳 Promise<Workbook>
  function buildWorkbook(src, p003Sheet) {
    var base = src.kind === 'prev'
      ? Promise.resolve(src.prev.build())
      : XIO.load(src.base.buffer.slice(0));
    return base.then(function (wb) {
      XIO.copySheet(p003Sheet, wb, SHEET_NAME);
      return wb;
    });
  }

  function outputName(src) {
    var name = src.kind === 'prev' ? src.prev.outName : src.base.name;
    var stem = name.replace(/\.[^.]+$/, '');
    if (stem.slice(-SUFFIX.length) === SUFFIX) stem = stem.slice(0, -SUFFIX.length);
    return stem + SUFFIX + '.xlsx';
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.withP003;
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

  IEApp.registerStep({ id: 'p003-import', render: render });
})(window);
