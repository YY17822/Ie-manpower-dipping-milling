/*
 * 第 1 步：篩選 MPS 報表（Dipping_Milling）
 *   1. 上傳 MPS 的 Excel
 *   2. 選篩選方式，留下 Cutting 和 Dipping 兩種列：
 *        a. 依製程（預設）：程式在料號那一行自動找「TTL (製程名)」列，
 *           TTL (Plastic Lens-Cutting) 上面那一塊是 Cutting，TTL (Plastic Lens-Dipping) 上面那一塊是 Dipping。
 *           每個製程最後一列料號的下面是「x」，再下面才是 TTL 列。
 *        b. 依 Status：使用者輸入 A 行等於多少代表 Cutting、等於多少代表 Dipping。
 *   3. 輸入要保留哪些行
 *   4. 匯出新的 Excel：第 1–4 列表頭原樣保留，下面分兩段——
 *        第 5 列     標題「Cutting（Plastic Lens-Cutting）」，跨所有保留的行置中
 *        接著        Cutting 的列
 *        空 10 列    留給第 2 步的匯總表格
 *        接著        標題「Dipping（Plastic Lens-Dipping）」，下面是 Dipping 的列
 *
 * 做完之後把結果放在 IEApp.state.mpsFiltered（sections 是兩段各在第幾列），第 2 步直接拿來用。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var XIO = global.ExcelIO;
  var ui = IEApp.ui;
  var h = IEApp.h;

  var STORE_KEY = 'step1';             // 共用設定的組名（IEApp.settings）
  var PREVIEW_ROWS = 8;       // 原始報表預覽顯示幾列資料
  var OUT_PREVIEW_MAX = 40;   // 匯出預覽最多顯示幾列（要看得到兩段）
  var MAX_CHIPS = 24;

  // 兩段的名稱。製程名稱可以在畫面的「報表版面不一樣時」改，標題會跟著變。
  var SECTIONS = [
    { key: 'cutting', label: 'Cutting', procSetting: 'cutProc', valueSetting: 'cutValues' },
    { key: 'dipping', label: 'Dipping', procSetting: 'dipProc', valueSetting: 'dipValues' }
  ];

  var DEFAULTS = {
    mode: 'process',            // 'process' 依製程、'status' 依 Status
    cutValues: '', dipValues: '',
    cols: '', startRow: 5, selCol: 'A', pnCol: 'B',
    cutProc: 'Plastic Lens-Cutting', dipProc: 'Plastic Lens-Dipping'
  };

  // 這一步的狀態（切到別的步驟再回來時還在）
  var S = { fileName: '', workbook: null, ws: null, info: null, infoKey: '', blocks: null, blocksKey: '', result: null };
  var settings = loadSettings();
  var els = {};
  var timer = null;

  /* ---------- 記住上次輸入的條件（每個月做的事情差不多） ---------- */
  function loadSettings() {
    var s = {};
    try { s = IEApp.settings.read(STORE_KEY) || {}; } catch (e) { s = {}; }
    var out = {};
    Object.keys(DEFAULTS).forEach(function (k) { out[k] = s[k] === undefined ? DEFAULTS[k] : s[k]; });
    if (out.mode !== 'status') out.mode = 'process';
    return out;
  }
  function saveSettings() {
    try { IEApp.settings.write(STORE_KEY, settings); } catch (e) { /* 存不了就算了 */ }
  }

  /* ---------- 讀取目前的設定值 ---------- */
  function startRow() {
    var n = parseInt(XIO.normalizeText(settings.startRow), 10);
    return n >= 1 ? n : DEFAULTS.startRow;
  }
  function oneCol(text, fallback) {
    var p = XIO.parseColumns(text);
    return p.cols.length === 1 && !p.bad.length ? p.cols[0] : fallback;
  }
  function selCol() { return oneCol(settings.selCol, 1); }
  function pnCol() { return oneCol(settings.pnCol, 2); }
  function valueList(text) {
    return XIO.normalizeText(text || '').split(/[,、;]+/)
      .map(function (v) { return v.trim(); })
      .filter(function (v) { return v !== ''; });
  }
  function procName(sec) {
    return XIO.normalizeText(settings[sec.procSetting] || '') || DEFAULTS[sec.procSetting];
  }
  function sectionTitle(sec) { return sec.label + '（' + procName(sec) + '）'; }
  function outputName() {
    return S.fileName.replace(/\.[^.]+$/, '') + '_篩選.xlsx';
  }

  /* ---------- 畫面 ---------- */
  function render(panel) {
    els = {};

    els.drop = ui.fileDrop({
      id: 'f-file',
      title: '選擇 MPS 的 Excel 檔',
      sub: '或把檔案拖進這裡。支援 .xlsx',
      onFile: openFile
    });
    els.sheetSelect = h('select', { id: 'f-sheet', 'aria-label': '工作表' });
    els.sheetRow = h('div', { class: 'field-row', hidden: true },
      h('label', { for: 'f-sheet', text: '工作表' }), els.sheetSelect);
    els.fileMsg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });

    // 篩選方式
    els.modeProcess = h('input', { type: 'radio', name: 'f-mode', id: 'f-mode-process', value: 'process' });
    els.modeStatus = h('input', { type: 'radio', name: 'f-mode', id: 'f-mode-status', value: 'status' });
    els.modeProcess.checked = settings.mode === 'process';
    els.modeStatus.checked = settings.mode === 'status';

    // a. 依製程
    els.pnCol = h('input', {
      id: 'f-pncol', type: 'text', class: 'in in-tiny', autocomplete: 'off', value: String(settings.pnCol)
    });
    els.procChips = h('div', { class: 'chips' });
    els.procPane = h('div', { class: 'mode-pane' },
      h('div', { class: 'field-row' },
        h('label', { for: 'f-pncol', text: '料號（P/N）在' }), els.pnCol, h('span', { text: '行' })),
      h('p', { class: 'hint', text: '程式會在這一行找「TTL (製程名)」：TTL (Plastic Lens-Cutting) 上面那一塊是 Cutting，' +
        'TTL (Plastic Lens-Dipping) 上面那一塊是 Dipping。「x」和 TTL 這兩列不會留下。' }),
      els.procChips);

    // b. 依 Status
    els.selLabels = [];
    els.valueInputs = {};
    var statusRows = SECTIONS.map(function (sec) {
      var lab = h('span', { text: 'A' });
      els.selLabels.push(lab);
      var input = h('input', {
        id: 'f-val-' + sec.key, type: 'text', class: 'in in-short', autocomplete: 'off',
        placeholder: '例如 1', value: settings[sec.valueSetting]
      });
      els.valueInputs[sec.key] = input;
      return h('div', { class: 'field-row' },
        h('label', { for: 'f-val-' + sec.key, class: 'label-fixed' }, sec.label + '：', lab, ' 行等於'), input);
    });
    els.statusChips = h('div', { class: 'chips' });
    els.statusPane = h('div', { class: 'mode-pane' },
      statusRows,
      h('p', { class: 'hint', text: '一種有多個數值時用逗號分開，例如 1, 3。同一個數值不能同時算 Cutting 和 Dipping。' }),
      els.statusChips);

    els.cols = h('input', {
      id: 'f-cols', type: 'text', class: 'in in-wide', autocomplete: 'off', spellcheck: 'false',
      placeholder: '例如 B, C, U, Z, AE', value: settings.cols
    });
    els.srcPreview = h('div', { class: 'sheet-area' });

    els.startRow = h('input', {
      id: 'f-start', type: 'text', class: 'in in-tiny', inputmode: 'numeric', value: String(settings.startRow)
    });
    els.selCol = h('input', {
      id: 'f-selcol', type: 'text', class: 'in in-tiny', autocomplete: 'off', value: String(settings.selCol)
    });
    els.cutProc = h('input', {
      id: 'f-cutproc', type: 'text', class: 'in in-short', autocomplete: 'off', spellcheck: 'false', value: settings.cutProc
    });
    els.dipProc = h('input', {
      id: 'f-dipproc', type: 'text', class: 'in in-short', autocomplete: 'off', spellcheck: 'false', value: settings.dipProc
    });

    els.summary = h('p', { class: 'summary', 'aria-live': 'polite' });
    els.outPreview = h('div', { class: 'sheet-area' });
    els.exportBtn = h('button', { type: 'button', class: 'btn-primary', disabled: true, text: '匯出 Excel', onclick: onExport });
    els.nextBtn = h('button', {
      type: 'button', class: 'btn-secondary', disabled: true, text: '下一步：匯總',
      onclick: function () { IEApp.go('mps-totals'); }
    });
    els.exportNote = h('span', { class: 'export-note', 'aria-live': 'polite' });

    panel.appendChild(h('div', { class: 'step-head' },
      h('h1', { text: '篩選 MPS 報表' }),
      h('p', { text: '留下 Cutting 和 Dipping 的列，以及你要的行，匯出成新的 Excel。' +
        '表頭（第 1–4 列）原樣保留，下面先放 Cutting、空 10 列，再放 Dipping，每一段上面有標題。' })));

    panel.appendChild(h('ol', { class: 'tasks' },
      h('li', { class: 'task' },
        h('h2', { text: '上傳 MPS 報表' }),
        els.drop.el, els.sheetRow, els.fileMsg),
      h('li', { class: 'task' },
        h('h2', { text: '篩選方式' }),
        h('div', { class: 'mode-pick', role: 'radiogroup', 'aria-label': '篩選方式' },
          h('label', { class: 'mode-option', for: 'f-mode-process' }, els.modeProcess,
            h('span', null, h('strong', { text: '依製程' }), h('span', { class: 'mode-sub', text: '自動找 TTL 列，分出 Cutting、Dipping' }))),
          h('label', { class: 'mode-option', for: 'f-mode-status' }, els.modeStatus,
            h('span', null, h('strong', { text: '依 Status' }), h('span', { class: 'mode-sub', text: '輸入 A 行等於多少代表 Cutting、Dipping' })))),
        els.procPane, els.statusPane),
      h('li', { class: 'task' },
        h('h2', { text: '要保留的行' }),
        h('div', { class: 'field-row' },
          h('label', { for: 'f-cols', text: '行號' }), els.cols),
        h('p', { class: 'hint', text: '連續的行可以寫成 U:Z。上傳後也可以直接點下面表格最上方的行號來選。' }),
        els.srcPreview,
        h('details', { class: 'advanced' },
          h('summary', { text: '報表版面不一樣時' }),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-start', text: '資料從第' }), els.startRow, h('span', { text: '列開始（之前的列都當作表頭保留）' })),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-selcol', text: 'Status 在' }), els.selCol, h('span', { text: '行（依 Status 篩選時用）' })),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-cutproc', text: 'Cutting 的製程名稱' }), els.cutProc),
          h('div', { class: 'field-row' },
            h('label', { for: 'f-dipproc', text: 'Dipping 的製程名稱' }), els.dipProc),
          h('p', { class: 'hint', text: '製程名稱就是 TTL ( ) 括號裡的字，也會用在匯出檔案的標題。比對時不分大小寫，空白和橫線有沒有都可以。' }))),
      h('li', { class: 'task' },
        h('h2', { text: '匯出' }),
        els.summary, els.outPreview,
        h('div', { class: 'actions' }, els.exportBtn, els.nextBtn, els.exportNote))));

    els.sheetSelect.addEventListener('change', function () {
      pickSheet(parseInt(els.sheetSelect.value, 10));
    });
    [els.modeProcess, els.modeStatus].forEach(function (radio) {
      radio.addEventListener('change', function () {
        if (radio.checked) { settings.mode = radio.value; recompute(); }
      });
    });
    SECTIONS.forEach(function (sec) { bindInput(els.valueInputs[sec.key], sec.valueSetting); });
    bindInput(els.cols, 'cols');
    bindInput(els.startRow, 'startRow');
    bindInput(els.selCol, 'selCol');
    bindInput(els.pnCol, 'pnCol');
    bindInput(els.cutProc, 'cutProc');
    bindInput(els.dipProc, 'dipProc');

    if (S.workbook) showFile();
    recompute();
  }

  function bindInput(input, key) {
    input.addEventListener('input', function () {
      settings[key] = input.value;
      clearTimeout(timer);
      timer = setTimeout(recompute, 200);
    });
  }

  /* ---------- 開檔 ---------- */
  function openFile(file) {
    els.fileMsg.hidden = true;
    els.drop.setBusy(file.name);
    ui.readWorkbook(file)
      .then(function (res) {
        S.workbook = res.workbook;
        S.fileName = file.name;
        S.ws = null;
        showFile();
        pickSheet(0);
      })
      .catch(function (err) {
        els.fileMsg.textContent = err && err.message ? err.message : String(err);
        els.fileMsg.hidden = false;
        if (S.workbook) showFile(); else els.drop.setIdle();
      });
  }

  function showFile() {
    els.drop.setFile(S.fileName);
    var sheets = S.workbook.worksheets;
    els.sheetSelect.textContent = '';
    sheets.forEach(function (ws, i) {
      els.sheetSelect.appendChild(h('option', { value: String(i), text: ws.name }));
    });
    els.sheetRow.hidden = sheets.length < 2;
    if (S.ws) els.sheetSelect.value = String(sheets.indexOf(S.ws));
  }

  function pickSheet(index) {
    S.ws = S.workbook.worksheets[index];
    S.infoKey = '';
    S.blocksKey = '';
    els.sheetSelect.value = String(index);
    els.exportNote.textContent = '';
    recompute();
  }

  /* ---------- 找出 Cutting、Dipping 各是哪幾列 ---------- */

  // 依製程：回傳 { rows: { cutting: [...], dipping: [...] } } 或 { error }
  function rowsByProcess(info, sr) {
    var pc = pnCol();
    var key = sr + ':' + pc + ':' + info.lastRow;
    if (S.blocksKey !== key) {
      S.blocks = XIO.detectProcessBlocks(S.ws, { dataStartRow: sr, keyCol: pc, lastRow: info.lastRow });
      S.blocksKey = key;
    }
    var blocks = S.blocks.blocks;
    var rows = {};
    var used = {};
    SECTIONS.forEach(function (sec) {
      var want = XIO.processKey(procName(sec));
      rows[sec.key] = [];
      blocks.forEach(function (b, i) {
        if (b.key !== want) return;
        used[i] = sec.label;
        rows[sec.key] = rows[sec.key].concat(b.rows);
      });
    });
    renderProcessChips(blocks, used);

    if (!blocks.length) {
      return { error: XIO.numToCol(pc) + ' 行找不到「TTL (製程名)」的列。請確認料號在哪一行（上面的「料號（P/N）在」），或改用「依 Status」。' };
    }
    var missing = SECTIONS.filter(function (sec) {
      var want = XIO.processKey(procName(sec));
      return !blocks.some(function (b) { return b.key === want; });
    });
    if (missing.length) {
      return {
        error: '找不到 ' + missing.map(function (sec) { return 'TTL (' + procName(sec) + ')'; }).join('、') +
          '。這份報表有的製程列在上面；名稱不一樣的話，到「報表版面不一樣時」改製程名稱。'
      };
    }
    return { rows: rows };
  }

  // 依 Status：回傳 { rows, wanted } 或 { error, wanted }
  function rowsByStatus(info, sr) {
    var sc = selCol();
    var wanted = {};
    var conflict = [];
    var lists = {};
    SECTIONS.forEach(function (sec) {
      lists[sec.key] = valueList(settings[sec.valueSetting]);
      lists[sec.key].forEach(function (v) {
        var k = XIO.matchKey(v);
        if (wanted[k] && wanted[k] !== sec.key) conflict.push(v);
        wanted[k] = sec.key;
      });
    });
    renderStatusChips(info, wanted);

    var colName = XIO.numToCol(sc);
    var empty = SECTIONS.filter(function (sec) { return !lists[sec.key].length; });
    if (empty.length) {
      return { wanted: wanted, error: '輸入 ' + colName + ' 行等於多少代表 ' + empty.map(function (sec) { return sec.label; }).join('、') + '。' };
    }
    if (conflict.length) {
      return { wanted: wanted, error: conflict.join('、') + ' 同時寫在 Cutting 和 Dipping，請只留一邊。' };
    }
    var rows = {};
    SECTIONS.forEach(function (sec) { rows[sec.key] = []; });
    for (var r = sr; r <= info.lastRow; r++) {
      var row = S.ws.findRow(r);
      if (!row) continue;
      var cell = row.findCell(sc);
      var k = XIO.matchKey(cell ? (cell.isMerged ? cell.master.value : cell.value) : null);
      if (k !== '' && wanted[k]) rows[wanted[k]].push(r);
    }
    return { rows: rows, wanted: wanted };
  }

  /* ---------- 重新計算並更新畫面 ---------- */
  function recompute() {
    saveSettings();
    S.result = null;
    IEApp.state.mpsFiltered = null;
    els.exportBtn.disabled = true;
    els.nextBtn.disabled = true;
    var byProcess = settings.mode !== 'status';
    els.procPane.hidden = !byProcess;
    els.statusPane.hidden = byProcess;
    var scName = XIO.numToCol(selCol());
    els.selLabels.forEach(function (l) { l.textContent = scName; });
    els.procChips.textContent = '';
    els.statusChips.textContent = '';

    if (!S.ws) {
      els.srcPreview.textContent = '';
      els.outPreview.textContent = '';
      setSummary('先上傳 MPS 報表。');
      return;
    }

    var sr = startRow();
    var sc = selCol();
    var key = sr + ':' + sc;
    if (S.infoKey !== key) {
      S.info = XIO.inspect(S.ws, { dataStartRow: sr, selectorCol: sc });
      S.infoKey = key;
    }
    var info = S.info;
    var parsed = XIO.parseColumns(settings.cols);
    var keep = {};
    parsed.cols.forEach(function (c) { keep[c] = true; });

    var picked = !info.dataRows ? {} : byProcess ? rowsByProcess(info, sr) : rowsByStatus(info, sr);

    ui.renderSheet(els.srcPreview, S.ws, {
      headerEnd: sr - 1,
      lastRow: info.lastRow,
      colCount: Math.max(info.colCount, 1),
      rows: PREVIEW_ROWS,
      keep: keep,
      onToggle: toggleColumn,
      selectorCol: sc,
      wanted: !byProcess && picked.rows ? picked.wanted : null,
      caption: '原始報表（前 ' + PREVIEW_ROWS + ' 列資料）。共 ' + info.dataRows.toLocaleString('en-US') +
        ' 列資料、' + info.colCount + ' 行（到 ' + XIO.numToCol(info.colCount || 1) + ' 行）。'
    });

    els.outPreview.textContent = '';

    if (!info.dataRows) {
      setSummary('第 ' + sr + ' 列之後沒有資料。請確認上傳的檔案和工作表。', true);
      return;
    }
    if (picked.error) { setSummary(picked.error, !byProcess || !/^輸入/.test(picked.error)); return; }
    if (parsed.bad.length) {
      setSummary('看不懂「' + parsed.bad.join('、') + '」。行號請用英文字母，例如 B 或 U:Z。', true);
      return;
    }
    if (!parsed.cols.length) { setSummary('輸入要保留的行。'); return; }

    var counts = SECTIONS.map(function (sec) { return picked.rows[sec.key].length; });
    if (!counts[0] && !counts[1]) {
      var seen = info.values.slice(0, 8).map(function (v) { return v.label; }).join('、');
      setSummary(byProcess ? 'Cutting 和 Dipping 這兩個製程底下都沒有料號。' :
        scName + ' 行沒有符合的列。' + (seen ? '這份報表的 ' + scName + ' 行有：' + seen + '。' : scName + ' 行從第 ' + sr + ' 列起都是空白。'), true);
      return;
    }

    // 這組條件存起來：匯出、以及第 2 步要重新產生活頁簿時都用同一組
    var ws = S.ws;
    var opts = {
      dataStartRow: sr, keepCols: parsed.cols, lastRow: info.lastRow, gap: XIO.SECTION_GAP,
      sections: SECTIONS.map(function (sec) {
        return { key: sec.key, title: sectionTitle(sec), rows: picked.rows[sec.key] };
      })
    };
    var result = XIO.filter(ws, opts);
    S.result = result;

    var text = SECTIONS.map(function (sec, i) { return sec.label + ' ' + counts[i].toLocaleString('en-US') + ' 列'; }).join('、') +
      '（共 ' + result.totalRows.toLocaleString('en-US') + ' 列資料），保留 ' + parsed.cols.length + ' 行。' +
      result.sections.map(function (s) {
        return s.title + '：標題在第 ' + s.titleRow + ' 列' + (s.count ? '，資料在第 ' + s.firstRow + '–' + s.lastRow + ' 列' : '');
      }).join('；') + '。';
    var warn = [];
    SECTIONS.forEach(function (sec, i) { if (!counts[i]) warn.push(sec.label + ' 沒有任何列，只會有標題。'); });
    var beyond = parsed.cols.filter(function (c) { return c > info.colCount; });
    if (beyond.length) warn.push(XIO.formatColumns(beyond) + ' 行在報表裡沒有資料。');
    if (warn.length) text += ' 注意：' + warn.join(' ');
    setSummary(text, false, warn.length > 0);

    ui.renderSheet(els.outPreview, result.worksheet, {
      headerEnd: result.headerEnd,
      lastRow: result.lastRow,
      colCount: result.colCount,
      rows: Math.min(result.lastRow - result.headerEnd, OUT_PREVIEW_MAX),
      caption: '匯出後的樣子。'
    });
    els.exportBtn.disabled = false;
    els.nextBtn.disabled = false;

    // 給第 2 步用。build() 每次都產生一份全新的活頁簿，後面的步驟可以放心往裡面加工作表。
    IEApp.state.mpsFiltered = {
      sourceName: S.fileName,
      outName: outputName(),
      sheetName: result.worksheet.name,
      matchedRows: result.matchedRows,
      headerEnd: result.headerEnd,
      sections: result.sections,      // [{ key, title, titleRow, firstRow, lastRow, count }]
      build: function () { return XIO.filter(ws, opts).workbook; }
    };
  }

  function setSummary(text, isError, isWarn) {
    els.summary.textContent = text;
    els.summary.className = 'summary' + (isError ? ' is-error' : '') + (isWarn ? ' is-warn' : '');
  }

  /* ---------- 報表裡找到的製程（標出用的是哪兩塊） ---------- */
  function renderProcessChips(blocks, used) {
    els.procChips.textContent = '';
    if (!blocks.length) return;
    els.procChips.appendChild(h('span', { class: 'chips-label', text: '報表裡的製程' }));
    blocks.forEach(function (b, i) {
      var tag = used[i];
      els.procChips.appendChild(h('span', {
        class: 'chip is-static' + (tag ? ' is-on' : ''),
        title: 'TTL 在第 ' + b.ttlRow + ' 列'
      }, tag ? h('span', { class: 'chip-tag', text: tag }) : null,
         h('span', { class: 'chip-value', text: b.name }),
         h('span', { class: 'chip-count', text: b.rows.length.toLocaleString('en-US') + ' 列' })));
    });
  }

  /* ---------- Status 行出現過的值（標出算 Cutting 還是 Dipping） ---------- */
  function renderStatusChips(info, wanted) {
    els.statusChips.textContent = '';
    if (!info.values.length && !info.blankCount) return;
    var col = XIO.numToCol(selCol());
    var labelOf = {};
    SECTIONS.forEach(function (sec) { labelOf[sec.key] = sec.label; });
    els.statusChips.appendChild(h('span', { class: 'chips-label', text: '報表裡 ' + col + ' 行的值' }));
    info.values.slice(0, MAX_CHIPS).forEach(function (v) {
      var tag = wanted[v.key] ? labelOf[wanted[v.key]] : '';
      els.statusChips.appendChild(h('span', { class: 'chip is-static' + (tag ? ' is-on' : '') },
        tag ? h('span', { class: 'chip-tag', text: tag }) : null,
        h('span', { class: 'chip-value', text: v.label }),
        h('span', { class: 'chip-count', text: v.count.toLocaleString('en-US') + ' 列' })));
    });
    if (info.values.length > MAX_CHIPS) {
      els.statusChips.appendChild(h('span', { class: 'chips-more', text: '另有 ' + (info.values.length - MAX_CHIPS) + ' 種' }));
    }
    if (info.blankCount) {
      els.statusChips.appendChild(h('span', { class: 'chips-more', text: '空白 ' + info.blankCount.toLocaleString('en-US') + ' 列' }));
    }
  }

  function toggleColumn(c) {
    var cols = XIO.parseColumns(settings.cols).cols;
    var i = cols.indexOf(c);
    if (i >= 0) cols.splice(i, 1); else cols.push(c);
    cols.sort(function (a, b) { return a - b; });
    settings.cols = XIO.formatColumns(cols);
    els.cols.value = settings.cols;
    var scroller = els.srcPreview.querySelector('.sheet-scroll');
    var left = scroller ? scroller.scrollLeft : 0;
    recompute();
    scroller = els.srcPreview.querySelector('.sheet-scroll');
    if (scroller) scroller.scrollLeft = left;
    var btn = els.srcPreview.querySelector('button[data-col="' + c + '"]');
    if (btn) btn.focus({ preventScroll: true });
  }

  /* ---------- 匯出 ---------- */
  function onExport() {
    var ready = IEApp.state.mpsFiltered;
    if (!S.result || !ready) return;
    var rows = ready.matchedRows;
    var name = ready.outName;
    els.exportBtn.disabled = true;
    els.exportNote.textContent = '產生檔案中…';
    XIO.toBlob(ready.build()).then(function (blob) {
      ui.download(blob, name);
      els.exportNote.textContent = '已匯出 ' + name + '（' + rows.toLocaleString('en-US') + ' 列），在瀏覽器的下載資料夾。';
    }).catch(function (err) {
      els.exportNote.textContent = '匯出失敗：' + (err && err.message ? err.message : err);
    }).then(function () {
      els.exportBtn.disabled = false;
    });
  }

  IEApp.registerStep({ id: 'mps-filter', render: render });
})(window);
