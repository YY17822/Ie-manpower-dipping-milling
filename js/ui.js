/*
 * ui.js — 各步驟共用的畫面元件：上傳區、試算表預覽、讀檔、下載，以及 WIP／月份行的共用設定。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var h = IEApp.h;

  /* ---------- 上傳區 ---------- */
  /*
   * o.id      input 的 id（每個上傳區要不同）
   * o.title   還沒選檔時的大字
   * o.sub     還沒選檔時的小字
   * o.onFile  選好檔案後呼叫 onFile(file)
   * 回傳 { el, setIdle(), setBusy(檔名), setFile(檔名, 小字) }
   */
  function fileDrop(o) {
    var input = h('input', { type: 'file', accept: o.accept || '.xlsx,.xlsm', class: 'visually-hidden', id: o.id });
    var main = h('span', { class: 'drop-main', text: o.title });
    var sub = h('span', { class: 'drop-sub', text: o.sub });
    var el = h('label', { class: 'drop', for: o.id }, input, main, sub);

    input.addEventListener('change', function () {
      var f = input.files[0];
      input.value = '';
      if (f) o.onFile(f);
    });
    ['dragenter', 'dragover'].forEach(function (ev) {
      el.addEventListener(ev, function (e) { e.preventDefault(); el.classList.add('is-over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      el.addEventListener(ev, function (e) { e.preventDefault(); el.classList.remove('is-over'); });
    });
    el.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) o.onFile(f);
    });

    return {
      el: el,
      setIdle: function () {
        el.classList.remove('has-file');
        main.textContent = o.title;
        sub.textContent = o.sub;
      },
      setBusy: function (name) {
        main.textContent = '讀取中…';
        sub.textContent = name;
      },
      setFile: function (name, note) {
        el.classList.add('has-file');
        main.textContent = name;
        sub.textContent = note || '點這裡換一個檔案';
      }
    };
  }

  /* ---------- 讀 Excel 檔 ---------- */
  // 回傳 Promise<{ workbook, buffer }>；失敗時 err.message 是可以直接顯示給使用者的中文訊息
  function readWorkbook(file) {
    return file.arrayBuffer().then(function (buffer) {
      return IEApp.loadExcelJS()
        .then(function () { return XIO.load(buffer); })
        .then(function (workbook) { return { workbook: workbook, buffer: buffer }; });
    });
  }

  /* ---------- 下載 ---------- */
  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = h('a', { href: url, download: name });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
  }

  /* ---------- 試算表預覽 ---------- */
  /*
   * o.headerEnd   表頭最後一列（沒有表頭就給 0）
   * o.lastRow     最後一列資料
   * o.colCount    要畫幾行
   * o.rows        表頭之後顯示幾列（預設 8）
   * o.caption     表格上方的說明文字
   * o.keep        { 行號: true }，要用螢光筆標示的行（可省略）
   * o.keepRows    { 列號: true }，要標示的列（可省略）。有給的話，只有「標示的列 × 標示的行」交叉的格子畫螢光
   * o.onToggle    點行號時呼叫（可省略；有給的話行號會變成按鈕）
   * o.selectorCol / o.wanted  用來把不符合條件的列變淡（可省略）
   */
  function renderSheet(container, ws, o) {
    container.textContent = '';
    var rows = o.rows || 8;
    var lastShown = Math.min(o.lastRow, o.headerEnd + rows);
    var colCount = o.colCount;
    var keep = o.keep || {};
    var keepRows = o.keepRows || null;

    // 合併儲存格 -> colspan / rowspan
    var span = {};
    var covered = {};
    XIO.listMerges(ws).forEach(function (m) {
      if (m.top > lastShown || m.left > colCount) return;
      var bottom = Math.min(m.bottom, lastShown);
      var right = Math.min(m.right, colCount);
      span[m.top + ':' + m.left] = { rows: bottom - m.top + 1, cols: right - m.left + 1 };
      for (var r = m.top; r <= bottom; r++) {
        for (var c = m.left; c <= right; c++) {
          if (r !== m.top || c !== m.left) covered[r + ':' + c] = true;
        }
      }
    });

    var headRow = h('tr', null, h('th', { class: 'corner', scope: 'col' }));
    for (var c = 1; c <= colCount; c++) {
      var letter = XIO.numToCol(c);
      var on = !!keep[c];
      var label = o.onToggle
        ? h('button', {
            type: 'button', 'data-col': String(c), 'aria-pressed': on ? 'true' : 'false',
            title: (on ? '取消保留 ' : '保留 ') + letter + ' 行',
            onclick: (function (col) { return function () { o.onToggle(col); }; })(c)
          }, letter)
        : letter;
      headRow.appendChild(h('th', { class: 'colhead' + (on ? ' is-kept' : ''), scope: 'col' }, label));
    }

    var body = h('tbody');
    for (var r = 1; r <= lastShown; r++) {
      var row = ws.findRow(r);
      var isHeader = r <= o.headerEnd;
      var out = false;
      if (!isHeader && o.wanted) {
        var k = XIO.matchKey(valueOf(row && row.findCell(o.selectorCol)));
        out = !(k !== '' && o.wanted[k]);
      }
      var tr = h('tr', { class: (isHeader ? 'is-header' : '') + (out ? ' is-out' : '') },
        h('th', { class: 'rowhead' + (keepRows && keepRows[r] ? ' is-kept' : ''), scope: 'row', text: String(r) }));
      for (c = 1; c <= colCount; c++) {
        if (covered[r + ':' + c]) continue;
        var cell = row ? row.findCell(c) : null;
        var sp = span[r + ':' + c];
        // 合併儲存格：全部的行都保留才整格畫螢光，只保留一部分就用淡色
        var width = sp ? sp.cols : 1;
        var keptCount = 0;
        for (var cc = c; cc < c + width; cc++) if (keep[cc]) keptCount++;
        if (keepRows && !keepRows[r]) keptCount = 0;
        var raw = XIO.plainValue(valueOf(cell));
        var cls = (keptCount === width ? 'is-kept' : keptCount ? 'is-part' : '') +
          (typeof raw === 'number' ? ' is-num' : '') + (sp ? ' is-merged' : '') + (width > 6 ? ' is-wide' : '');
        var td = h('td', { class: cls.trim() || null, text: cell ? XIO.displayText(cell) : '' });
        if (sp && sp.cols > 1) td.colSpan = sp.cols;
        if (sp && sp.rows > 1) td.rowSpan = sp.rows;
        tr.appendChild(td);
      }
      body.appendChild(tr);
    }

    var more = o.lastRow - lastShown;
    container.appendChild(h('p', {
      class: 'sheet-caption',
      text: (o.caption || '') + (more > 0 ? ' 還有 ' + more.toLocaleString('en-US') + ' 列沒顯示。' : '')
    }));
    container.appendChild(h('div', { class: 'sheet-scroll', tabindex: '0', role: 'group', 'aria-label': '試算表預覽' },
      h('table', { class: 'sheet' }, h('thead', null, headRow), body)));
  }

  function valueOf(cell) {
    if (!cell) return null;
    return cell.isMerged ? cell.master.value : cell.value;
  }

  /* ---------- WIP、各月產量在哪幾行（第 2 步匯總和第 5 步計算需求工時共用） ---------- */
  // 兩個步驟讀寫的是同一份設定，在其中一步改了，另一步也會跟著變，避免兩邊對不起來。
  var COLUMNS_KEY = IEApp.storeKey('columns');
  var COLUMNS_OLD_KEY = IEApp.storeKey('step4');       // 舊版存在這裡，第一次讀的時候搬過來

  function readJson(key) {
    try { return JSON.parse(localStorage.getItem(key) || 'null'); } catch (e) { return null; }
  }
  function getColumns() {
    var s = readJson(COLUMNS_KEY) || readJson(COLUMNS_OLD_KEY) || {};
    return {
      wipCol: s.wipCol === undefined ? '' : String(s.wipCol),
      monthCols: s.monthCols === undefined ? '' : String(s.monthCols)
    };
  }
  function setColumns(value) {
    try {
      localStorage.setItem(COLUMNS_KEY, JSON.stringify({ wipCol: value.wipCol || '', monthCols: value.monthCols || '' }));
    } catch (e) { /* 存不了就算了 */ }
  }

  IEApp.ui = {
    fileDrop: fileDrop,
    readWorkbook: readWorkbook,
    download: download,
    renderSheet: renderSheet,
    columns: { get: getColumns, set: setColumns }
  };
})(window);
