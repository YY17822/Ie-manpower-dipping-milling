/*
 * excel-io.js — 所有 Excel 讀寫與排版都集中在這裡（使用 ExcelJS）。
 * 各步驟只呼叫 ExcelIO，不直接碰 ExcelJS，之後要調整匯出格式只需要改這個檔。
 *
 * 名詞：行 = 直的（A、B、C…），列 = 橫的（1、2、3…）。
 */
(function (global) {
  'use strict';

  /* ---------- 行號（字母）與數字互轉 ---------- */

  function colToNum(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  function numToCol(n) {
    var s = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // 全形轉半形、去頭尾空白（使用者常用全形輸入 Ａ、１）
  function normalizeText(s) {
    return String(s).normalize('NFKC').trim();
  }

  /*
   * 解析「B, C, U:Z, AE」這類輸入，回傳 { cols: [2,3,21,...], bad: ['看不懂的字'] }
   * 接受逗號、頓號、空白、分號分隔；範圍可用 : 或 - 或 ~
   */
  function parseColumns(text) {
    var cols = {};
    var bad = [];
    var tokens = normalizeText(text || '').toUpperCase().split(/[\s,、;]+/).filter(Boolean);
    tokens.forEach(function (tok) {
      var m = /^([A-Z]{1,3})(?:[:\-~]([A-Z]{1,3}))?$/.exec(tok);
      if (!m) { bad.push(tok); return; }
      var a = colToNum(m[1]);
      var b = m[2] ? colToNum(m[2]) : a;
      if (a > 16384 || b > 16384) { bad.push(tok); return; }
      if (a > b) { var t = a; a = b; b = t; }
      for (var c = a; c <= b; c++) cols[c] = true;
    });
    var list = Object.keys(cols).map(Number).sort(function (x, y) { return x - y; });
    return { cols: list, bad: bad };
  }

  function formatColumns(cols) {
    return cols.map(numToCol).join(', ');
  }

  /* ---------- 儲存格的值 ---------- */

  // 把 ExcelJS 的各種值型態（公式、RichText、超連結…）化成單純的值
  function plainValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v !== 'object') return v;
    if (v.formula !== undefined || v.sharedFormula !== undefined) return plainValue(v.result);
    if (Array.isArray(v.richText)) {
      return v.richText.map(function (t) { return t.text; }).join('');
    }
    if (v.error !== undefined) return v.error;
    if (v.text !== undefined) return plainValue(v.text);
    if (v.result !== undefined) return plainValue(v.result);
    return String(v);
  }

  // 合併儲存格裡不是左上角的格子，值要去找左上角那一格
  function cellRawValue(cell) {
    if (!cell) return null;
    var src = cell.isMerged ? cell.master : cell;
    var v = src.value;
    // ExcelJS 的 value 會把「計算結果剛好是 0」的公式結果丟掉（變成沒有結果），用 result 補回來
    if (v && typeof v === 'object' && !(v instanceof Date) && v.result === undefined &&
        (v.formula !== undefined || v.sharedFormula !== undefined)) {
      var res = src.result;
      if (res !== undefined && res !== null) v = Object.assign({}, v, { result: res });
    }
    return v;
  }

  // 比對用的鍵：1、"1"、"１"、"1.0" 都視為相同
  function matchKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'number') return String(v);
    var s = normalizeText(v);
    if (s !== '' && !isNaN(Number(s))) return String(Number(s));
    return s;
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  // 預覽畫面顯示用的文字
  function displayText(cell) {
    var v = plainValue(cellRawValue(cell));
    if (v === null || v === undefined) return '';
    if (v instanceof Date) {
      var d = v.getUTCFullYear() + '/' + (v.getUTCMonth() + 1) + '/' + v.getUTCDate();
      if (v.getUTCHours() || v.getUTCMinutes()) {
        d += ' ' + pad2(v.getUTCHours()) + ':' + pad2(v.getUTCMinutes());
      }
      return d;
    }
    if (typeof v === 'number') {
      var fmt = (cell && cell.numFmt) || '';
      if (fmt.indexOf('%') >= 0) return (Math.round(v * 10000) / 100) + '%';
      var rounded = Math.round(v * 10000) / 10000;
      if (fmt.indexOf('#,##') >= 0) return rounded.toLocaleString('en-US');
      return String(rounded);
    }
    return String(v);
  }

  // 匯出時寫入的值：公式換成計算結果（篩掉列與行之後，原公式的參照會錯位）
  function exportValue(cell) {
    var v = cellRawValue(cell);
    if (v && typeof v === 'object' && !(v instanceof Date) &&
        (v.formula !== undefined || v.sharedFormula !== undefined)) {
      return v.result === undefined ? null : v.result;
    }
    return v === undefined ? null : v;
  }

  /* ---------- 讀檔 ---------- */

  function ioError(code, message) {
    var e = new Error(message);
    e.code = code;
    return e;
  }

  /*
   * 讀檔前的修正：字型旗標。
   *
   * ExcelJS 4.x 讀字型設定時，只要看到 <b>、<i>、<strike> 這些標籤就當成「有」，
   * 不管後面是不是寫著 val="0"（＝沒有）。SAP 這類系統匯出的報表會把「沒有粗體、
   * 沒有刪除線」明寫成 <b val="0"/> <strike val="0"/>，結果整張表被讀成
   * 粗體＋斜體＋刪除線，匯出後就全部被畫上橫線。
   *
   * 做法：.xlsx 其實是一個 zip 壓縮檔。交給 ExcelJS 之前，先把裡面的樣式檔
   * （styles.xml）和文字檔（sharedStrings.xml）解開，拿掉這些「明寫為沒有」的標籤，
   * 再包回去。解壓縮用下面自己寫的 inflateRaw()，不靠瀏覽器的功能，所以 Chrome、
   * Edge、Safari 新舊版本的結果都一樣。任何一步不順利就原封不動交給 ExcelJS。
   */
  var OFF_FLAGS = /<(b|i|strike|outline|shadow|condense|extend)\s+val\s*=\s*(["'])(?:0|false)\2\s*(?:\/>|>\s*<\/\1>)/g;
  var FIX_ENTRIES = { 'xl/styles.xml': true, 'xl/sharedStrings.xml': true };

  /* 解開 zip 裡用 deflate 壓縮的資料。src：壓縮資料；size：解開後的大小。 */
  var LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  var CODE_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  function huffman(lengths, start, n) {
    var count = new Uint16Array(16);
    var symbol = new Uint16Array(n);
    var offs = new Uint16Array(16);
    var i;
    for (i = 0; i < n; i++) count[lengths[start + i]]++;
    for (i = 1; i < 15; i++) offs[i + 1] = offs[i] + count[i];
    for (i = 0; i < n; i++) {
      var len = lengths[start + i];
      if (len) symbol[offs[len]++] = i;
    }
    return { count: count, symbol: symbol };
  }

  function inflateRaw(src, size) {
    var out = new Uint8Array(size);
    var op = 0, ip = 0, bitBuf = 0, bitCnt = 0;

    function bits(n) {
      while (bitCnt < n) {
        if (ip >= src.length) throw new Error('inflate: 資料不完整');
        bitBuf |= src[ip++] << bitCnt;
        bitCnt += 8;
      }
      var v = bitBuf & ((1 << n) - 1);
      bitBuf >>>= n;
      bitCnt -= n;
      return v;
    }
    function decode(h) {
      var code = 0, first = 0, index = 0;
      for (var len = 1; len <= 15; len++) {
        code |= bits(1);
        var cnt = h.count[len];
        if (code - cnt < first) return h.symbol[index + (code - first)];
        index += cnt;
        first = (first + cnt) << 1;
        code <<= 1;
      }
      throw new Error('inflate: 編碼錯誤');
    }
    function block(lenCode, distCode) {
      for (;;) {
        var sym = decode(lenCode);
        if (sym < 256) {
          if (op >= size) throw new Error('inflate: 超出大小');
          out[op++] = sym;
        } else if (sym === 256) {
          return;
        } else {
          sym -= 257;
          if (sym >= 29) throw new Error('inflate: 長度錯誤');
          var len = LEN_BASE[sym] + bits(LEN_EXTRA[sym]);
          var ds = decode(distCode);
          if (ds >= 30) throw new Error('inflate: 距離錯誤');
          var dist = DIST_BASE[ds] + bits(DIST_EXTRA[ds]);
          if (dist > op || op + len > size) throw new Error('inflate: 距離錯誤');
          for (var k = 0; k < len; k++, op++) out[op] = out[op - dist];
        }
      }
    }

    var last, lengths, i;
    do {
      last = bits(1);
      var type = bits(2);
      if (type === 0) {                      // 沒壓縮的區塊
        bitBuf = 0; bitCnt = 0;
        if (ip + 4 > src.length) throw new Error('inflate: 資料不完整');
        var n = src[ip] | (src[ip + 1] << 8);
        ip += 4;
        if (ip + n > src.length || op + n > size) throw new Error('inflate: 資料不完整');
        out.set(src.subarray(ip, ip + n), op);
        ip += n; op += n;
      } else if (type === 1) {               // 固定編碼
        lengths = new Uint8Array(288 + 30);
        for (i = 0; i < 144; i++) lengths[i] = 8;
        for (; i < 256; i++) lengths[i] = 9;
        for (; i < 280; i++) lengths[i] = 7;
        for (; i < 288; i++) lengths[i] = 8;
        for (; i < 318; i++) lengths[i] = 5;
        block(huffman(lengths, 0, 288), huffman(lengths, 288, 30));
      } else if (type === 2) {               // 動態編碼
        var nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        if (nlen > 286 || ndist > 30) throw new Error('inflate: 編碼表錯誤');
        lengths = new Uint8Array(320);
        for (i = 0; i < ncode; i++) lengths[CODE_ORDER[i]] = bits(3);
        var codeLen = huffman(lengths, 0, 19);
        lengths = new Uint8Array(320);
        i = 0;
        while (i < nlen + ndist) {
          var sym = decode(codeLen);
          if (sym < 16) {
            lengths[i++] = sym;
          } else {
            var prev = 0, rep;
            if (sym === 16) {
              if (i === 0) throw new Error('inflate: 編碼表錯誤');
              prev = lengths[i - 1];
              rep = 3 + bits(2);
            } else if (sym === 17) {
              rep = 3 + bits(3);
            } else {
              rep = 11 + bits(7);
            }
            if (i + rep > nlen + ndist) throw new Error('inflate: 編碼表錯誤');
            while (rep--) lengths[i++] = prev;
          }
        }
        block(huffman(lengths, 0, nlen), huffman(lengths, nlen, ndist));
      } else {
        throw new Error('inflate: 區塊類型錯誤');
      }
    } while (!last);
    if (op !== size) throw new Error('inflate: 大小不符');
    return out;
  }

  var crcTable = null;
  function crc32(bytes) {
    if (!crcTable) {
      crcTable = new Uint32Array(256);
      for (var n = 0; n < 256; n++) {
        var c = n;
        for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        crcTable[n] = c >>> 0;
      }
    }
    var crc = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  // 讀出 zip 裡每個檔案的資訊與（仍是壓縮狀態的）內容。看不懂的格式回傳 null。
  function readZip(buffer) {
    var dv = new DataView(buffer);
    var size = buffer.byteLength;
    var eocd = -1;
    for (var i = size - 22; i >= 0 && i >= size - 22 - 65535; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) return null;
    var count = dv.getUint16(eocd + 10, true);
    var p = dv.getUint32(eocd + 16, true);
    if (count === 0xFFFF || p === 0xFFFFFFFF) return null;       // zip64，不處理
    var decoder = new TextDecoder();
    var entries = [];
    for (var k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) return null;
      var e = {
        flags: dv.getUint16(p + 8, true),
        method: dv.getUint16(p + 10, true),
        time: dv.getUint16(p + 12, true),
        date: dv.getUint16(p + 14, true),
        crc: dv.getUint32(p + 16, true),
        csize: dv.getUint32(p + 20, true),
        usize: dv.getUint32(p + 24, true),
        intAttr: dv.getUint16(p + 36, true),
        extAttr: dv.getUint32(p + 38, true)
      };
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var local = dv.getUint32(p + 42, true);
      if (e.flags & 1) return null;                              // 有加密，不處理
      if (e.csize === 0xFFFFFFFF || e.usize === 0xFFFFFFFF || local === 0xFFFFFFFF) return null;
      e.nameBytes = new Uint8Array(buffer, p + 46, nameLen);
      e.name = decoder.decode(e.nameBytes);
      if (dv.getUint32(local, true) !== 0x04034b50) return null;
      var dataStart = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      if (dataStart + e.csize > size) return null;
      e.data = new Uint8Array(buffer, dataStart, e.csize);
      entries.push(e);
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  // 把檔案清單重新包成 zip
  function writeZip(entries) {
    var total = 22;
    entries.forEach(function (e) { total += 30 + 46 + 2 * e.nameBytes.length + e.data.length; });
    var out = new Uint8Array(total);
    var dv = new DataView(out.buffer);
    var p = 0;
    entries.forEach(function (e) {
      e.offset = p;
      e.outFlags = e.flags & (e.method === 8 ? 0x0806 : 0x0800);  // 留下壓縮等級與 UTF-8 檔名旗標
      dv.setUint32(p, 0x04034b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, e.outFlags, true);
      dv.setUint16(p + 8, e.method, true);
      dv.setUint16(p + 10, e.time, true);
      dv.setUint16(p + 12, e.date, true);
      dv.setUint32(p + 14, e.crc, true);
      dv.setUint32(p + 18, e.data.length, true);
      dv.setUint32(p + 22, e.usize, true);
      dv.setUint16(p + 26, e.nameBytes.length, true);
      out.set(e.nameBytes, p + 30);
      out.set(e.data, p + 30 + e.nameBytes.length);
      p += 30 + e.nameBytes.length + e.data.length;
    });
    var cdStart = p;
    entries.forEach(function (e) {
      dv.setUint32(p, 0x02014b50, true);
      dv.setUint16(p + 4, 20, true);
      dv.setUint16(p + 6, 20, true);
      dv.setUint16(p + 8, e.outFlags, true);
      dv.setUint16(p + 10, e.method, true);
      dv.setUint16(p + 12, e.time, true);
      dv.setUint16(p + 14, e.date, true);
      dv.setUint32(p + 16, e.crc, true);
      dv.setUint32(p + 20, e.data.length, true);
      dv.setUint32(p + 24, e.usize, true);
      dv.setUint16(p + 28, e.nameBytes.length, true);
      dv.setUint16(p + 36, e.intAttr, true);
      dv.setUint32(p + 38, e.extAttr, true);
      dv.setUint32(p + 42, e.offset, true);
      out.set(e.nameBytes, p + 46);
      p += 46 + e.nameBytes.length;
    });
    dv.setUint32(p, 0x06054b50, true);
    dv.setUint16(p + 8, entries.length, true);
    dv.setUint16(p + 10, entries.length, true);
    dv.setUint32(p + 12, p - cdStart, true);
    dv.setUint32(p + 16, cdStart, true);
    return out.buffer;
  }

  // 回傳修正過的檔案內容（ArrayBuffer）；不需要修或修不了就回傳原本的
  function fixFontFlags(buffer) {
    try {
      var entries = readZip(buffer);
      if (!entries) return buffer;
      var changed = false;
      entries.forEach(function (e) {
        if (!FIX_ENTRIES[e.name] || (e.method !== 8 && e.method !== 0)) return;
        var bytes = e.method === 8 ? inflateRaw(e.data, e.usize) : e.data;
        if (crc32(bytes) !== e.crc) return;                       // 解出來的內容不對，不處理
        if (bytes[0] === 0xFF || bytes[0] === 0xFE) return;       // 不是 UTF-8，不處理
        var text = new TextDecoder('utf-8').decode(bytes);
        var fixed = text.replace(OFF_FLAGS, '');
        if (fixed === text) return;
        var plain = new TextEncoder().encode(fixed);
        e.data = plain;                                           // 改過的檔案不壓縮，直接存
        e.method = 0;
        e.crc = crc32(plain);
        e.usize = plain.length;
        changed = true;
      });
      return changed ? writeZip(entries) : buffer;
    } catch (err) {
      return buffer;
    }
  }

  function load(arrayBuffer) {
    var head = new Uint8Array(arrayBuffer.slice(0, 4));
    if (head[0] === 0xD0 && head[1] === 0xCF) {
      return Promise.reject(ioError('OLD_XLS',
        '這個檔案是舊版 .xls 格式，或有密碼保護。請在 Excel 用「另存新檔」存成 .xlsx（並取消密碼）後再上傳。'));
    }
    if (!(head[0] === 0x50 && head[1] === 0x4B)) {
      return Promise.reject(ioError('NOT_XLSX',
        '這不是 Excel 活頁簿（.xlsx）。請確認上傳的是 .xlsx 檔。'));
    }
    var wb = new global.ExcelJS.Workbook();
    return wb.xlsx.load(fixFontFlags(arrayBuffer)).then(function () {
      if (!wb.worksheets.length) throw ioError('NO_SHEET', '這個活頁簿裡沒有工作表。');
      return wb;
    }, function (err) {
      if (err && err.code) throw err;
      throw ioError('READ_FAIL', '讀不了這個 Excel 檔：' + (err && err.message ? err.message : err));
    });
  }

  /* ---------- 合併儲存格 ---------- */

  function listMerges(ws) {
    var out = [];
    var merges = ws._merges || {};
    Object.keys(merges).forEach(function (key) {
      var m = merges[key];
      if (m) out.push({ top: m.top, left: m.left, bottom: m.bottom, right: m.right });
    });
    return out;
  }

  /* ---------- 檢視工作表內容 ---------- */

  /*
   * 回傳：
   *   lastRow     最後一列有資料的列號
   *   colCount    最右邊有資料的行號
   *   dataRows    資料列數（從 dataStartRow 起）
   *   values      選取行出現過的值 [{ key, label, count }]，依出現順序
   *   blankCount  選取行是空白的列數
   */
  function inspect(ws, opts) {
    var dataStartRow = opts.dataStartRow;
    var selectorCol = opts.selectorCol;
    var lastRow = 0;
    var colCount = 0;
    var counts = {};
    var order = [];
    var nonBlank = 0;

    ws.eachRow({ includeEmpty: false }, function (row, r) {
      var rowHasValue = false;
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        var v = cellRawValue(cell);
        if (v === null || v === undefined || v === '') return;
        rowHasValue = true;
        if (c > colCount) colCount = c;
      });
      if (!rowHasValue) return;
      if (r > lastRow) lastRow = r;
      if (r < dataStartRow) return;
      var cell = row.findCell(selectorCol);
      var key = matchKey(cellRawValue(cell));
      if (key === '') return;
      nonBlank++;
      if (!counts[key]) {
        counts[key] = { key: key, label: displayText(cell), count: 0 };
        order.push(key);
      }
      counts[key].count++;
    });

    var dataRows = Math.max(0, lastRow - dataStartRow + 1);
    return {
      lastRow: lastRow,
      colCount: colCount,
      dataRows: dataRows,
      values: order.map(function (k) { return counts[k]; }),
      blankCount: dataRows - nonBlank
    };
  }

  /* ---------- 工作表的範圍 ---------- */

  // 逐列、逐格走訪「實際存在」的列與儲存格（包含只有格式、沒有值的格子）
  function eachStoredRow(ws, fn) {
    var rows = ws._rows;
    if (Array.isArray(rows)) {
      for (var i = 0; i < rows.length; i++) if (rows[i]) fn(rows[i], i + 1);
    } else {
      ws.eachRow({ includeEmpty: false }, fn);
    }
  }
  function eachStoredCell(row, fn) {
    var cells = row._cells;
    if (Array.isArray(cells)) {
      for (var i = 0; i < cells.length; i++) if (cells[i]) fn(cells[i], i + 1);
    } else {
      row.eachCell({ includeEmpty: false }, fn);
    }
  }

  // 回傳 { lastRow, colCount }：最後一列、最右一行「有值」的位置
  function extent(ws) {
    var lastRow = 0;
    var colCount = 0;
    eachStoredRow(ws, function (row, r) {
      var has = false;
      eachStoredCell(row, function (cell, c) {
        var v = cellRawValue(cell);
        if (v === null || v === undefined || v === '') return;
        has = true;
        if (c > colCount) colCount = c;
      });
      if (has && r > lastRow) lastRow = r;
    });
    return { lastRow: lastRow, colCount: colCount };
  }

  /* ---------- 表格外觀：框線、字型（MPS 工作表上各塊表格共用） ---------- */

  var BORDER_THIN = { style: 'thin' };
  var BORDER_THICK = { style: 'medium' };

  // 一塊表格（top..bottom 列 × left..right 行）：外框粗線、裡面細線。只改框線，其他樣式不動。
  function boxBorders(ws, top, left, bottom, right) {
    if (bottom < top || right < left) return;
    for (var r = top; r <= bottom; r++) {
      for (var c = left; c <= right; c++) {
        var cell = ws.getCell(r, c);
        var st = Object.assign({}, cell.style);
        st.border = {
          top: r === top ? BORDER_THICK : BORDER_THIN,
          bottom: r === bottom ? BORDER_THICK : BORDER_THIN,
          left: c === left ? BORDER_THICK : BORDER_THIN,
          right: c === right ? BORDER_THICK : BORDER_THIN
        };
        cell.style = st;
      }
    }
  }

  // 字型只留名稱、大小（和要加的粗體等），拿掉顏色、底線等（原報表有藍字之類的，抄過來會很雜）
  function plainFont(font, extra) {
    var f = {};
    if (font) {
      if (font.name) f.name = font.name;
      if (font.size) f.size = font.size;
      if (font.family) f.family = font.family;
      if (font.charset) f.charset = font.charset;
    }
    return Object.assign(f, extra || {});
  }

  // MPS 工作表的基本字型：第一段第一列料號那一格的字型（拿掉顏色、粗體）
  var DEFAULT_FONT = { name: 'Arial', size: 10 };
  function sheetFont(ws, headerEnd) {
    var secs = findSections(ws, { headerEnd: headerEnd, keyCol: 1 });
    var row = secs.length && secs[0].count ? secs[0].firstRow : headerEnd + 1;
    var cell = ws.findCell(row, 1);
    var f = plainFont(cell && cell.font);
    return f.name ? f : Object.assign({}, DEFAULT_FONT);
  }

  // 把一個區域的字型換成 font（保留粗體與否可以選），其他樣式不動
  function setFont(ws, top, left, bottom, right, font) {
    for (var r = top; r <= bottom; r++) {
      for (var c = left; c <= right; c++) {
        var cell = ws.getCell(r, c);
        cell.style = Object.assign({}, cell.style, { font: Object.assign({}, font) });
      }
    }
  }

  /*
   * 新增的一塊（CT.、需求工時等）統一樣式：
   *   表頭（第 2 列到表頭最後一列）：基本字型、置中、外框粗線裡面細線
   *   每一段（Cutting、Dipping）的資料列：基本字型、數字格式 numFmt、外框粗線裡面細線；
   *     原本標了淺紅色（找不到 CT.）的格子保留淺紅色，其他底色拿掉
   */
  function styleBlock(ws, o) {
    var font = o.font;
    var top = o.headerEnd >= 2 ? 2 : 1;
    var r, c, cell, old;
    for (r = top; r <= o.headerEnd; r++) {
      for (c = o.left; c <= o.right; c++) {
        cell = ws.getCell(r, c);
        cell.style = {
          font: Object.assign({}, font),
          alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
        };
      }
    }
    boxBorders(ws, top, o.left, o.headerEnd, o.right);
    o.sections.forEach(function (s) {
      if (!s.count) return;
      for (r = s.firstRow; r <= s.lastRow; r++) {
        for (c = o.left; c <= o.right; c++) {
          cell = ws.getCell(r, c);
          old = cell.style || {};
          var st = { font: Object.assign({}, font), alignment: { vertical: 'middle' } };
          var fmt = (o.numFmtByCol && o.numFmtByCol[c]) || o.numFmt || old.numFmt;
          if (fmt) st.numFmt = fmt;
          if (isMissingFill(old.fill)) st.fill = MISSING_FILL;
          cell.style = st;
        }
      }
      boxBorders(ws, s.firstRow, o.left, s.lastRow, o.right);
    });
  }

  /* ---------- 篩選並產生新的活頁簿 ---------- */

  function cleanObject(o) {
    var out = {};
    Object.keys(o).forEach(function (k) { if (o[k] !== undefined) out[k] = o[k]; });
    return out;
  }

  /*
   * opts:
   *   dataStartRow  資料從第幾列開始（之前的列是表頭，全部保留）
   *   selectorCol   選取條件所在的行（數字，A=1）
   *   values        要選取的值（字串陣列）。有給 sections 時不用
   *   keepCols      要保留的行（數字陣列，已排序）
   *   lastRow       最後一列資料的列號
   *   sections      分段輸出（可省略）：[{ key, title, rows: [原列號…] }]
   *                 每一段先寫一列標題（跨所有保留的行置中），下面接這一段的資料列；
   *                 段和段之間空 gap 列（預設 10 列，留給第 2 步的匯總表格）。
   *   gap           段和段之間空幾列
   *
   * 表頭列的樣式、合併儲存格、行寬、列高都會照原樣帶到新檔；
   * 合併儲存格如果只保留了其中幾行，會縮成只合併那幾行。
   */
  // 第 1 步篩選時把儲存格的填滿色彩拿掉（變透明）：原報表有些列塗了底色，
  // 留著的話會蓋過後面步驟標的前三大淺橘色，匯出的檔案也會一直帶著那些顏色。
  function withoutFill(style) {
    var out = Object.assign({}, style || {});
    delete out.fill;
    return out;
  }

  function filter(ws, opts) {
    var ExcelJS = global.ExcelJS;
    var dataStartRow = opts.dataStartRow;
    var headerEnd = dataStartRow - 1;
    var keepCols = opts.keepCols;
    var wanted = {};
    (opts.values || []).forEach(function (v) { wanted[matchKey(v)] = true; });

    var newIndex = {}; // 原行號 -> 新行號
    keepCols.forEach(function (c, i) { newIndex[c] = i + 1; });

    var out = new ExcelJS.Workbook();
    out.creator = 'IE 人力需求小程式';
    out.created = new Date();
    var props = ws.properties || {};
    var ows = out.addWorksheet(ws.name || 'MPS', {
      properties: cleanObject({
        defaultRowHeight: props.defaultRowHeight,
        defaultColWidth: props.defaultColWidth
      })
    });

    // 行寬
    keepCols.forEach(function (c, i) {
      var w = ws.getColumn(c).width;
      if (w) ows.getColumn(i + 1).width = w;
    });

    // 表頭列：每一格先照抄（值只有左上角那格有，其餘合併格只抄樣式）
    var r, i, src, dst, srcRow, dstRow;
    for (r = 1; r <= headerEnd; r++) {
      srcRow = ws.findRow(r);
      if (!srcRow) continue;
      dstRow = ows.getRow(r);
      if (srcRow.height) dstRow.height = srcRow.height;
      for (i = 0; i < keepCols.length; i++) {
        src = srcRow.findCell(keepCols[i]);
        if (!src) continue;
        dst = dstRow.getCell(i + 1);
        dst.style = withoutFill(src.style);
        if (!src.isMerged || src.master === src) dst.value = exportValue(src);
      }
    }

    // 表頭的合併儲存格：對應到新的行號後重新合併
    listMerges(ws).forEach(function (m) {
      if (m.top > headerEnd) return;
      var bottom = Math.min(m.bottom, headerEnd);
      var kept = keepCols.filter(function (c) { return c >= m.left && c <= m.right; });
      if (!kept.length) return;
      var n1 = newIndex[kept[0]];
      var n2 = newIndex[kept[kept.length - 1]];
      var master = ws.getCell(m.top, m.left);
      var target = ows.getCell(m.top, n1);
      target.value = exportValue(master);
      if (kept[0] !== m.left) target.style = withoutFill(master.style);
      if (n2 > n1 || bottom > m.top) {
        ows.mergeCellsWithoutStyle(m.top, n1, bottom, n2);
      }
    });

    // 第 1 列的大標題如果落在沒保留的行，補到新檔的左上角，避免標題不見
    srcRow = headerEnd >= 1 ? ws.findRow(1) : null;
    if (srcRow) {
      var titleCell = null;
      srcRow.eachCell({ includeEmpty: false }, function (cell) {
        if (titleCell || (cell.isMerged && cell.master !== cell)) return;
        var tv = cell.value;
        if (tv !== null && tv !== undefined && tv !== '') titleCell = cell;
      });
      var titleKept = false;
      for (i = 1; i <= keepCols.length; i++) {
        var cv = cellRawValue(ows.getCell(1, i));
        if (cv !== null && cv !== undefined && cv !== '') { titleKept = true; break; }
      }
      if (titleCell && !titleKept) {
        var a1 = ows.getCell(1, 1);
        a1.value = exportValue(titleCell);
        a1.style = withoutFill(titleCell.style);
      }
    }

    // 複製一列資料到新檔的下一列
    var outRow = dataStartRow;
    function copyDataRow(r) {
      var sRow = ws.findRow(r);
      var dRow = ows.getRow(outRow++);
      if (!sRow) return;
      if (sRow.height) dRow.height = sRow.height;
      for (var k = 0; k < keepCols.length; k++) {
        var sc = sRow.findCell(keepCols[k]);
        if (!sc) continue;
        var dc = dRow.getCell(k + 1);
        dc.style = withoutFill(sc.style);
        var val = exportValue(sc);
        if (val !== null) dc.value = val;
      }
    }

    var matched = 0;
    var total = Math.max(0, opts.lastRow - dataStartRow + 1);
    var outSections = null;

    if (opts.sections) {
      // 分段：標題列 + 這一段的資料列，段和段之間空 gap 列
      var gap = opts.gap === undefined ? SECTION_GAP : opts.gap;
      var baseFont = (ws.getCell(dataStartRow, keepCols[0]).font) || {};
      outSections = [];
      opts.sections.forEach(function (sec, si) {
        if (si > 0) outRow += gap;
        var titleRow = outRow++;
        var tcell = ows.getCell(titleRow, 1);
        tcell.value = sec.title;
        tcell.style = {
          font: cleanObject({ name: baseFont.name, size: baseFont.size, bold: true }),
          alignment: { horizontal: 'center', vertical: 'middle' }
        };
        if (keepCols.length > 1) ows.mergeCellsWithoutStyle(titleRow, 1, titleRow, keepCols.length);
        var first = outRow;
        sec.rows.forEach(copyDataRow);
        matched += sec.rows.length;
        outSections.push({
          key: sec.key, title: sec.title, titleRow: titleRow,
          firstRow: first, lastRow: outRow - 1, count: sec.rows.length
        });
      });

      // 外觀：表頭（第 2 列到表頭最後一列）一塊、每一段的資料一塊，外框粗線、裡面細線；字一律黑色
      var n = keepCols.length;
      var hTop = headerEnd >= 2 ? 2 : 1;
      var font = plainFont(baseFont);
      if (!font.name) font = Object.assign({}, DEFAULT_FONT);
      if (headerEnd >= 1) {
        for (r = hTop; r <= headerEnd; r++) {
          for (i = 1; i <= n; i++) {
            var hc = ows.getCell(r, i);
            hc.style = Object.assign({}, hc.style, { font: plainFont(hc.font && hc.font.name ? hc.font : font) });
          }
        }
        boxBorders(ows, hTop, 1, headerEnd, n);
      }
      outSections.forEach(function (s) {
        ows.getCell(s.titleRow, 1).style = {
          font: plainFont(font, { bold: true }),
          alignment: { horizontal: 'center', vertical: 'middle' }
        };
        if (!s.count) return;
        setFont(ows, s.firstRow, 1, s.lastRow, n, font);
        boxBorders(ows, s.firstRow, 1, s.lastRow, n);
      });
    } else {
      // 不分段：選取行的值符合才保留
      for (r = dataStartRow; r <= opts.lastRow; r++) {
        srcRow = ws.findRow(r);
        if (!srcRow) continue;
        var key = matchKey(cellRawValue(srcRow.findCell(opts.selectorCol)));
        if (key === '' || !wanted[key]) continue;
        matched++;
        copyDataRow(r);
      }
    }

    // 顯示比例、格線照原檔；凍結窗格不帶過去（打開檔案時畫面不要被凍住）
    var v = (ws.views && ws.views[0]) || {};
    var view = cleanObject({ showGridLines: v.showGridLines, zoomScale: v.zoomScale });
    ows.views = [Object.assign({ state: 'normal' }, view)];

    // 列印設定：紙張方向照原檔，每頁重複表頭
    var ps = ws.pageSetup || {};
    Object.assign(ows.pageSetup, cleanObject({
      orientation: ps.orientation,
      paperSize: ps.paperSize,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      printTitlesRow: headerEnd >= 1 ? '1:' + headerEnd : undefined
    }));

    return {
      workbook: out,
      worksheet: ows,
      matchedRows: matched,
      totalRows: total,
      headerEnd: headerEnd,
      lastRow: outRow - 1,
      colCount: keepCols.length,
      sections: outSections
    };
  }

  /* ---------- MPS 裡的製程區塊（TTL 列） ---------- */

  /*
   * MPS 報表裡有很多製程，在料號那一行，每個製程最後一列料號的下面是「x」，
   * 再下面是「TTL (製程名)」，例如 TTL (Plastic Lens-Cutting)。
   * 一個製程的資料列 = 上一個 TTL 列（或資料開始的那一列）之後、到這個 TTL 之前，
   * 料號那一格有值、而且不是「x」的列。
   *
   * o: { dataStartRow, keyCol, lastRow }
   * 回傳 { blocks: [{ name, key, ttlRow, rows: [列號…] }], trailing: [最後一個 TTL 之後的列號…] }
   */
  var SECTION_GAP = 10;   // 分段輸出時，段和段之間空幾列（留給第 2 步的匯總表格）
  var TTL_RE = /^TTL\s*[(（]\s*(.*?)\s*[)）]\s*$/i;

  // 製程名稱比對用：不分大小寫，忽略空白、橫線、底線、括號（Plastic Lens-Dipping = plastic lens dipping）
  function processKey(name) {
    return normalizeText(name == null ? '' : name).toLowerCase().replace(/[\s\-_－—–‐()（）]/g, '');
  }

  function detectProcessBlocks(ws, o) {
    var blocks = [];
    var cur = [];
    for (var r = o.dataStartRow; r <= o.lastRow; r++) {
      var row = ws.findRow(r);
      var v = row ? plainValue(cellRawValue(row.findCell(o.keyCol))) : null;
      var t = v === null || v === undefined ? '' : normalizeText(v instanceof Date ? '' : String(v));
      if (t === '') continue;
      if (/^TTL\b/i.test(t)) {
        var m = TTL_RE.exec(t);
        var name = m ? m[1] : t.replace(/^TTL\s*/i, '');
        blocks.push({ name: name, key: processKey(name), ttlRow: r, rows: cur });
        cur = [];
        continue;
      }
      if (t.toLowerCase() === 'x') continue;
      cur.push(r);
    }
    return { blocks: blocks, trailing: cur };
  }

  /* ---------- 把整張工作表複製到另一個活頁簿 ---------- */

  /*
   * 把 src 工作表原樣複製進 dstWorkbook，取名為 name。
   * 值、格式、合併儲存格、行寬、列高、隱藏的行列、凍結窗格、自動篩選都會帶過去；
   * 公式換成計算結果（跨檔案之後原公式的參照會失效）。
   * dstWorkbook 裡如果已經有同名的工作表，會先移除再放新的。
   * 回傳 { worksheet, lastRow, colCount }
   */
  function copySheet(src, dstWorkbook, name) {
    // Excel 的工作表名稱不分大小寫，p003 和 P003 算同名
    dstWorkbook.worksheets
      .filter(function (w) { return String(w.name).toLowerCase() === String(name).toLowerCase(); })
      .forEach(function (w) { dstWorkbook.removeWorksheet(w.id); });

    var props = src.properties || {};
    var ws = dstWorkbook.addWorksheet(name, {
      properties: cleanObject({
        defaultRowHeight: props.defaultRowHeight,
        defaultColWidth: props.defaultColWidth
      })
    });

    var ext = extent(src);
    var maxCol = ext.colCount;

    eachStoredRow(src, function (row, r) {
      var nrow = null;
      eachStoredCell(row, function (cell, c) {
        if (!nrow) nrow = ws.getRow(r);
        var ncell = nrow.getCell(c);
        ncell.style = cell.style;
        if (c > maxCol) maxCol = c;
        if (cell.isMerged && cell.master !== cell) return;
        var v = exportValue(cell);
        if (v !== null && v !== '') ncell.value = v;
      });
      if (row.height || row.hidden) {
        if (!nrow) nrow = ws.getRow(r);
        if (row.height) nrow.height = row.height;
        if (row.hidden) nrow.hidden = true;
      }
    });

    for (var c = 1; c <= maxCol; c++) {
      var col = src.getColumn(c);
      if (col.width) ws.getColumn(c).width = col.width;
      if (col.hidden) ws.getColumn(c).hidden = true;
    }

    listMerges(src).forEach(function (m) {
      ws.mergeCellsWithoutStyle(m.top, m.left, m.bottom, m.right);
    });

    var v = (src.views && src.views[0]) || {};
    var view = cleanObject({ showGridLines: v.showGridLines, zoomScale: v.zoomScale });
    if (v.state === 'frozen') {
      view.state = 'frozen';
      view.xSplit = v.xSplit || 0;
      view.ySplit = v.ySplit || 0;
    }
    if (Object.keys(view).length) ws.views = [view];

    if (src.autoFilter) ws.autoFilter = src.autoFilter;

    var ps = src.pageSetup || {};
    Object.assign(ws.pageSetup, cleanObject({
      orientation: ps.orientation,
      paperSize: ps.paperSize,
      fitToPage: ps.fitToPage,
      fitToWidth: ps.fitToWidth,
      fitToHeight: ps.fitToHeight,
      printTitlesRow: ps.printTitlesRow
    }));

    return { worksheet: ws, lastRow: ext.lastRow, colCount: ext.colCount };
  }

  /* ---------- 表頭文字 ---------- */

  // 猜表頭在哪一列：前幾列當中，有值的格子最多的那一列（大標題通常只有一兩格有值）
  function guessHeaderRow(ws, maxRows) {
    var best = 1, bestCount = -1;
    for (var r = 1; r <= (maxRows || 10); r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      var n = 0;
      eachStoredCell(row, function (cell) {
        if (cell.isMerged && cell.master !== cell) return;
        var v = cell.value;
        if (v !== null && v !== undefined && v !== '') n++;
      });
      if (n > bestCount) { bestCount = n; best = r; }
    }
    return best;
  }

  // 每一行的表頭文字（fromRow 到 toRow 之間，由上到下串起來；合併儲存格用左上角的字）
  // 回傳陣列，labels[行號] = '2027/01 月合計'
  function columnLabels(ws, fromRow, toRow, colCount) {
    var labels = [];
    for (var c = 1; c <= colCount; c++) {
      var parts = [];
      for (var r = fromRow; r <= toRow; r++) {
        var row = ws.findRow(r);
        var cell = row ? row.findCell(c) : null;
        var t = cell ? normalizeText(displayText(cell)).replace(/\s+/g, ' ') : '';
        if (t && parts.indexOf(t) < 0) parts.push(t);
      }
      labels[c] = parts.join(' ');
    }
    return labels;
  }

  /* ---------- 查表：用料號到另一張工作表找數值 ---------- */

  // 料號比對用的鍵：去頭尾空白、全形轉半形、不分大小寫。
  // 數字和文字的 123 視為相同，但 0012 和 12 不同（料號前面的 0 有意義）。
  function lookupKey(v) {
    v = plainValue(v);
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    return normalizeText(v).toUpperCase();
  }

  // 把參照工作表的某一行做成索引：{ 鍵: [列號, 列號…] }（由上到下）
  function buildIndex(ref, keyCol) {
    var index = {};
    eachStoredRow(ref, function (row, r) {
      var k = lookupKey(cellRawValue(row.findCell(keyCol)));
      if (k === '') return;
      if (!index[k]) index[k] = [];
      index[k].push(r);
    });
    return index;
  }

  function sheetRef(name) {
    return "'" + String(name).replace(/'/g, "''") + "'!";
  }

  // 在表頭列裡找文字是 header 的那一行（重做這一步時，覆蓋原本那一行而不是再加一行）
  function findHeaderColumn(ws, headerEnd, header) {
    var want = lookupKey(header);
    var found = 0;
    for (var r = 1; r <= headerEnd && !found; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      eachStoredCell(row, function (cell, c) {
        if (found || (cell.isMerged && cell.master !== cell)) return;
        if (lookupKey(cell.value) === want) found = c;
      });
    }
    return found;
  }

  // 把儲存格的值當成數字：數字照用、文字的 "100" 也算；空白回傳 null；不是數字回傳 NaN
  function toNumber(v) {
    v = plainValue(v);
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return NaN;
    var t = normalizeText(v).replace(/,/g, '');
    if (t === '') return null;
    var n = Number(t);
    return isNaN(n) ? NaN : n;
  }

  // 取出 ref 第 r 列要帶回去的數值。有給 o.refDivCol 時是「取值行 ÷ 除數行」。
  // 回傳 { value, error, numFmt, top, bottom }；error 是 Excel 會顯示的錯誤（例如 #DIV/0!）
  function lookupValue(ref, r, o) {
    var row = ref.findRow(r);
    var topCell = row ? row.findCell(o.refValueCol) : null;
    if (!o.refDivCol) {
      return { value: plainValue(cellRawValue(topCell)), error: null, numFmt: topCell ? topCell.numFmt : undefined };
    }
    var bottomCell = row ? row.findCell(o.refDivCol) : null;
    var top = toNumber(cellRawValue(topCell));
    var bottom = toNumber(cellRawValue(bottomCell));
    var out = {
      value: null, error: null, numFmt: 'General',
      top: topCell ? displayText(topCell) : '', bottom: bottomCell ? displayText(bottomCell) : ''
    };
    if (top !== null && isNaN(top)) out.error = '#VALUE!';
    else if (bottom !== null && isNaN(bottom)) out.error = '#VALUE!';
    else if (bottom === null || bottom === 0) out.error = '#DIV/0!';
    else out.value = (top === null ? 0 : top) / bottom;
    return out;
  }

  function sameResult(a, b) {
    if (a.error || b.error) return a.error === b.error;
    if (typeof a.value === 'number' && typeof b.value === 'number') {
      return Math.abs(a.value - b.value) <= 1e-12 * Math.max(1, Math.abs(a.value), Math.abs(b.value));
    }
    return lookupKey(a.value) === lookupKey(b.value);
  }

  /*
   * 先算出查表結果，不動工作表（給畫面預覽、統計用）。
   *
   * o.dataStartRow  main 的資料從第幾列開始
   * o.keyCol        main 拿來對照的行（放料號的那一行）
   * o.refKeyCol     ref 要在哪一行找料號
   * o.refValueCol   ref 找到之後取哪一行的數值（例如 Labor）
   * o.refDivCol     （可省略）要除以 ref 的哪一行（例如 Base Qty）；有給的話帶回的是 取值行 ÷ 這一行
   * o.header        新增那一行的表頭文字，例如 'CT.'
   * o.index         （可省略）buildIndex(ref, o.refKeyCol) 的結果，重複計算時可以沿用
   *
   * 同一個料號在 ref 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）。
   *
   * 回傳：
   *   targetCol   要填入的行（原本最後一行之後空一行；已經有同名表頭就用那一行）
   *   isNew       是不是新增的行
   *   rows        [{ row, key, label, refRow, value, error, top, bottom }]  refRow = 0 表示找不到
   *   found / missing / blank   找到、找不到、料號空白的列數
   *   errors      找到了、但算不出數值的列數（除數是 0 或空白、或不是數字）
   *   multi       找到的料號裡，有幾個在 ref 出現不只一列
   *   multiDiff   其中有幾個「各列算出來的數值不一樣」
   */
  function planLookup(main, ref, o) {
    var headerEnd = o.dataStartRow - 1;
    var ext = extent(main);
    var existing = findHeaderColumn(main, headerEnd, o.header);
    var lastCol = ext.colCount;
    var targetCol = existing || o.newCol || lastCol + 2;   // o.newCol：新增時指定放哪一行（例如緊接在另一個 CT. 右邊）
    var index = o.index || buildIndex(ref, o.refKeyCol);

    var rows = [];
    var found = 0, missing = 0, blank = 0, errors = 0;
    var multiKeys = {}, diffKeys = {};
    // 要查的列：有給 o.rows（例如 Cutting、Dipping 兩段的資料列）就只查那些列；
    // 沒給就從資料開始的那一列查到最後一個料號（下面的匯總列不算資料）
    var rowList = o.rows;
    var lastDataRow;
    if (rowList) {
      lastDataRow = rowList.length ? Math.max.apply(null, rowList) : 0;
    } else {
      lastDataRow = lastKeyRow(main, o.keyCol, o.dataStartRow);
      rowList = [];
      for (var rr = o.dataStartRow; rr <= lastDataRow; rr++) rowList.push(rr);
    }
    for (var ri = 0; ri < rowList.length; ri++) {
      var r = rowList[ri];
      var row = main.findRow(r);
      var cell = row ? row.findCell(o.keyCol) : null;
      var key = lookupKey(cellRawValue(cell));
      if (key === '') { blank++; continue; }
      var hits = index[key];
      var item = { row: r, key: key, label: displayText(cell), refRow: 0, value: null, error: null };
      if (hits && hits.length) {
        found++;
        var first = lookupValue(ref, hits[0], o);
        item.refRow = hits[0];
        item.value = first.value;
        item.error = first.error;
        item.numFmt = first.numFmt;
        item.top = first.top;
        item.bottom = first.bottom;
        if (first.error) errors++;
        if (hits.length > 1 && !multiKeys[key]) {
          multiKeys[key] = true;
          for (var i = 1; i < hits.length; i++) {
            if (!sameResult(first, lookupValue(ref, hits[i], o))) { diffKeys[key] = true; break; }
          }
        }
        item.multi = hits.length;
        item.differs = !!diffKeys[key];
      } else {
        missing++;
      }
      rows.push(item);
    }

    return {
      targetCol: targetCol,
      isNew: !existing,
      lastCol: existing ? existing : lastCol,
      lastDataRow: lastDataRow,
      rows: rows,
      found: found,
      missing: missing,
      blank: blank,
      errors: errors,
      multi: Object.keys(multiKeys).length,
      multiDiff: Object.keys(diffKeys).length
    };
  }

  // 某一列要填入的公式（不含開頭的 =），例如 'P003'!M5 或 'P003'!M5/'P003'!K5
  function lookupFormula(refName, r, o) {
    var prefix = sheetRef(refName);
    var f = prefix + numToCol(o.refValueCol) + r;
    if (o.refDivCol) f += '/' + prefix + numToCol(o.refDivCol) + r;
    return f;
  }

  /*
   * 把查表結果寫進 main 工作表：每一列填入公式（='P003'!M5 或 ='P003'!M5/'P003'!K5），
   * 點儲存格就看得到數值是從 ref 的哪一格來的。找不到的料號留空白並標上底色。
   * 參數和 planLookup 一樣，回傳的也是同一份結果。
   */
  var MISSING_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD9D9' } };
  function isMissingFill(fill) {
    return !!(fill && fill.fgColor && fill.fgColor.argb === MISSING_FILL.fgColor.argb);
  }

  function addLookupColumn(main, ref, o) {
    var plan = planLookup(main, ref, o);
    var headerEnd = o.dataStartRow - 1;
    var col = plan.targetCol;

    // 表頭：樣式照抄對照那一行（料號那一行）的表頭，大標題那一列不動
    if (headerEnd >= 1) {
      var top = headerEnd >= 2 ? 2 : 1;
      if (plan.isNew) {
        for (var r = top; r <= headerEnd; r++) {
          var src = main.findCell(r, o.keyCol);
          if (src) main.getCell(r, col).style = src.style;
        }
        if (headerEnd > top) main.mergeCellsWithoutStyle(top, col, headerEnd, col);
      }
      var headCell = main.getCell(top, col);
      (headCell.isMerged ? headCell.master : headCell).value = o.header;
    }

    // 資料：樣式照抄原本最後一行（重做時沿用這一行原本的樣式），數字格式改用 ref 那一格的
    var done = {};
    plan.rows.forEach(function (item) {
      var cell = main.getCell(item.row, col);
      var base = main.findCell(item.row, plan.isNew ? plan.lastCol : col);
      var style = Object.assign({}, base ? base.style : {});
      if (isMissingFill(style.fill)) delete style.fill;
      done[item.row] = true;
      if (item.refRow) {
        style.numFmt = item.numFmt || 'General';
        cell.style = style;
        cell.value = {
          formula: lookupFormula(ref.name, item.refRow, o),
          result: item.error ? { error: item.error }
            : (item.value === null || item.value === '' ? 0 : item.value)
        };
      } else {
        style.fill = MISSING_FILL;
        cell.style = style;
        cell.value = null;
      }
    });

    // 重做時，把這一行裡不再對得到的舊內容清掉
    if (!plan.isNew) {
      eachStoredRow(main, function (row, r) {
        if (r < o.dataStartRow || done[r]) return;
        if (!o.rows && r > plan.lastDataRow) return;
        var old = row.findCell(col);
        if (!old) return;
        old.value = null;
        if (o.rows && isMissingFill(old.fill)) {
          var st = Object.assign({}, old.style);
          delete st.fill;
          old.style = st;
        }
      });
    }

    if (plan.isNew) main.getColumn(col).width = o.width || 11;

    // 請 Excel 開檔時重新計算一次公式
    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  /* ---------- Dipping_Milling：機台 CT.、目檢 CT. 兩行 ---------- */

  var CT_HEADERS = { machine: '機台 CT.', labor: '目檢 CT.', rate: '目檢-檢驗率' };
  var CT_FORMAT = '0.000';     // CT. 顯示到小數三位（儲存格裡仍是完整的數值）
  var RATE_FORMAT = '0.0%';    // 目檢-檢驗率顯示成 %

  // 分段 MPS 裡要查 CT. 的列：Cutting、Dipping 兩段的資料列（標題、空白、Total 列都不算）
  function sectionDataRows(main, headerEnd) {
    var rows = [];
    findSections(main, { headerEnd: headerEnd, keyCol: 1 }).forEach(function (s) {
      for (var r = s.firstRow; r <= s.lastRow; r++) rows.push(r);
    });
    return rows;
  }

  /*
   * 先算出兩行 CT. 的查表結果，不動工作表。
   *   機台 CT. = P003 的 Machine ÷ Base Qty（o.refMachineCol ÷ o.refBaseCol）
   *   目檢 CT. = P003 的 Labor ÷ Base Qty（o.refLaborCol ÷ o.refBaseCol）
   * 兩行並排：機台 CT. 放在原本最後一行之後空一行，目檢 CT. 緊接在右邊；已經有同名表頭就覆蓋那一行。
   *
   * o: { dataStartRow, keyCol, refKeyCol, refMachineCol, refLaborCol, refBaseCol,
   *      rows（可省略，預設是兩段的資料列）, index }
   * 回傳 { machine: planLookup 的結果, labor: planLookup 的結果, rows }
   */
  function ctOptions(o, which, newCol) {
    return {
      dataStartRow: o.dataStartRow, keyCol: o.keyCol, refKeyCol: o.refKeyCol,
      refValueCol: which === 'machine' ? o.refMachineCol : o.refLaborCol,
      refDivCol: o.refBaseCol,
      header: CT_HEADERS[which], rows: o.rows, index: o.index, newCol: newCol, width: 11
    };
  }
  function planCtColumns(main, ref, o) {
    var opts = Object.assign({}, o);
    if (!opts.rows) opts.rows = sectionDataRows(main, o.dataStartRow - 1);
    if (!opts.index) opts.index = buildIndex(ref, o.refKeyCol);
    var machine = planLookup(main, ref, ctOptions(opts, 'machine'));
    var labor = planLookup(main, ref, ctOptions(opts, 'labor', machine.targetCol + 1));
    var rate = planRateColumn(main, o.dataStartRow - 1, labor.targetCol);
    return { machine: machine, labor: labor, rate: rate, rows: opts.rows };
  }

  /*
   * 目檢-檢驗率：放在目檢 CT. 右邊一行（已經有這個表頭就用原本那一行）。
   * 數值是使用者在畫面上輸入的，Cutting、Dipping 各一個，填到該段的每一個料號列（數值，不是公式）。
   */
  function planRateColumn(main, headerEnd, laborCol) {
    var existing = findHeaderColumn(main, headerEnd, CT_HEADERS.rate);
    return { targetCol: existing || laborCol + 1, isNew: !existing };
  }

  // o.rates = { cutting: 0.9, dipping: 1 }（小數，0.9 = 90%）；沒給的段留空白
  function addRateColumn(main, o, laborCol, keyCol) {
    var headerEnd = o.dataStartRow - 1;
    var plan = planRateColumn(main, headerEnd, laborCol);
    var col = plan.targetCol;
    var top = headerEnd >= 2 ? 2 : 1;
    // 表頭：第 2 列到表頭最後一列合併（和 CT. 一樣）
    if (headerEnd >= 1) {
      if (plan.isNew && headerEnd > top) main.mergeCellsWithoutStyle(top, col, headerEnd, col);
      var head = main.getCell(top, col);
      (head.isMerged ? head.master : head).value = CT_HEADERS.rate;
    }
    // 重做時先清掉這一行原本的數值
    if (!plan.isNew) {
      eachStoredRow(main, function (row, r) {
        if (r <= headerEnd) return;
        var old = row.findCell(col);
        if (old) old.value = null;
      });
    }
    findSections(main, { headerEnd: headerEnd, keyCol: 1 }).forEach(function (s) {
      var v = o.rates ? o.rates[s.key] : undefined;
      for (var r = s.firstRow; r <= s.lastRow; r++) {
        var key = main.findCell(r, keyCol || 1);
        if (!key || normalizeText(displayText(key)) === '') continue;
        main.getCell(r, col).value = typeof v === 'number' && !isNaN(v) ? v : null;
      }
    });
    if (plan.isNew) main.getColumn(col).width = 11;
    return plan;
  }

  // 寫進工作表：每一列填公式，例如 ='P003'!N5/'P003'!K5（機台 CT.）、='P003'!M5/'P003'!K5（目檢 CT.）。
  // 找不到的料號留空白、標淺紅色。
  function addCtColumns(main, ref, o) {
    var opts = Object.assign({}, o);
    if (!opts.rows) opts.rows = sectionDataRows(main, o.dataStartRow - 1);
    if (!opts.index) opts.index = buildIndex(ref, o.refKeyCol);
    var machine = addLookupColumn(main, ref, ctOptions(opts, 'machine'));
    var labor = addLookupColumn(main, ref, ctOptions(opts, 'labor', machine.targetCol + 1));
    var rate = addRateColumn(main, o, labor.targetCol, o.keyCol);
    // 外觀：機台 CT.、目檢 CT.、目檢-檢驗率三行當成一塊（表頭一框、每一段一框），字一律黑色；
    // CT. 顯示到小數三位，檢驗率顯示成 %
    var headerEnd = o.dataStartRow - 1;
    var fmts = {};
    fmts[machine.targetCol] = CT_FORMAT;
    fmts[labor.targetCol] = CT_FORMAT;
    fmts[rate.targetCol] = RATE_FORMAT;
    var cols = [machine.targetCol, labor.targetCol, rate.targetCol];
    styleBlock(main, {
      headerEnd: headerEnd, left: Math.min.apply(null, cols), right: Math.max.apply(null, cols),
      sections: findSections(main, { headerEnd: headerEnd, keyCol: 1 }), font: sheetFont(main, headerEnd), numFmtByCol: fmts
    });
    return { machine: machine, labor: labor, rate: rate, rows: opts.rows };
  }

  /* ---------- 需求工時：各月數量 × CT. ---------- */

  // 儲存格的值當成數字用；公式結果是錯誤（#DIV/0! 等）時回傳 { error }
  function numberOrError(cell) {
    var raw = cellRawValue(cell);
    if (raw && typeof raw === 'object' && !(raw instanceof Date)) {
      var res = raw.result !== undefined ? raw.result : raw;
      if (res && typeof res === 'object' && res.error) return { error: res.error };
      if (raw.error) return { error: raw.error };
    }
    var n = toNumber(raw);
    if (n !== null && isNaN(n)) return { error: '#VALUE!' };
    // blank：這一格真的什麼都沒有（公式不算空白，就算它還沒有計算結果）
    var empty = raw === null || raw === undefined || raw === '';
    return { value: n === null ? 0 : n, blank: empty };
  }

  // 找出已經存在的區塊（表頭文字是 title 的那一格）：回傳 { col, width }，沒有就回傳 null
  function findBlock(ws, headerEnd, title) {
    var col = findHeaderColumn(ws, headerEnd, title);
    if (!col) return null;
    var width = 1;
    listMerges(ws).forEach(function (m) {
      if (m.left === col && m.top <= headerEnd && m.right > m.left) {
        var cell = ws.findCell(m.top, m.left);
        if (cell && lookupKey(cell.value) === lookupKey(title)) width = m.right - m.left + 1;
      }
    });
    return { col: col, width: width };
  }

  /*
   * 先算出需求工時，不動工作表（給畫面預覽、統計用）。
   *
   * o.dataStartRow  資料從第幾列開始
   * o.keyCol        料號在哪一行（只拿來顯示）
   * o.monthCols     各月數量在哪幾行（陣列，第一個是當月）
   * o.ctCol         CT. 在哪一行
   * o.title         大表頭文字，例如 '需求工時'
   * o.headers       各月的小表頭文字（使用者自己輸入的月份）
   *
   * 公式：各月 =數量*CT.（當月也一樣，不加 WIP）
   *
   * 回傳：
   *   startCol    要填入的第一行（原本最後一行之後空一行；已經有同名大表頭就用原本的位置）
   *   isNew       是不是新增
   *   headers     各月的小表頭文字
   *   rows        [{ row, label, ct, noCt, values: [數字或 { error }], formulas: [...] }]
   *   noCt        沒有 CT. 的列數（CT. 是空白）
   *   errors      算不出來的列數（CT. 或數量是錯誤值、不是數字）
   */
  function planDemand(main, o) {
    var headerEnd = o.dataStartRow - 1;
    var ext = extent(main);
    var old = findBlock(main, headerEnd, o.title);
    var startCol = old ? old.col : (o.startCol || ext.colCount + 2);   // o.startCol：新增時指定從哪一行開始
    var lastRow = ext.lastRow;

    var headers = o.monthCols.map(function (c, i) {
      return o.headers && o.headers[i] !== undefined ? normalizeText(o.headers[i]) : '';
    });

    var ctL = numToCol(o.ctCol);
    // o.rateCol（可省略）：再乘上這一行（目檢-檢驗率），例如 =E6*M6*N6
    var rateL = o.rateCol ? numToCol(o.rateCol) : '';
    var rows = [];
    var noCt = 0, errors = 0;
    // 要算的列：有給 o.rows（Cutting、Dipping 兩段的資料列）就只算那些列
    var rowList = o.rows;
    if (!rowList) {
      rowList = [];
      for (var rr = o.dataStartRow; rr <= lastRow; rr++) rowList.push(rr);
    }
    for (var ri = 0; ri < rowList.length; ri++) {
      var r = rowList[ri];
      var row = main.findRow(r);
      if (!row) continue;
      var keyCell = row.findCell(o.keyCol);
      var label = keyCell ? displayText(keyCell) : '';
      if (normalizeText(label) === '') continue;
      var ct = numberOrError(row.findCell(o.ctCol));
      var rate = o.rateCol ? numberOrError(row.findCell(o.rateCol)) : { value: 1 };
      var item = {
        row: r, label: label, ct: ct.error ? ct.error : ct.value, noCt: !!ct.blank,
        rate: rate.error ? rate.error : rate.value, noRate: !!(o.rateCol && rate.blank), values: [], formulas: []
      };
      var times = ctL + r + (rateL ? '*' + rateL + r : '');
      var factor = ct.value * rate.value;
      var bad = false;
      o.monthCols.forEach(function (c, i) {
        var qty = numberOrError(row.findCell(c));
        var qL = numToCol(c);
        var err = ct.error || rate.error || qty.error;
        item.formulas.push(qL + r + '*' + times);
        item.values.push(err ? { error: err } : qty.value * factor);
        if (err) bad = true;
      });
      if (item.noCt) noCt++;
      if (bad) errors++;
      rows.push(item);
    }

    return {
      startCol: startCol,
      isNew: !old,
      oldWidth: old ? old.width : 0,
      headers: headers,
      rows: rows,
      noCt: noCt,
      errors: errors
    };
  }

  var INTEGER_FORMAT = '#,##0';

  function centered(style) {
    return Object.assign({}, style, {
      alignment: Object.assign({}, style && style.alignment, { horizontal: 'center', vertical: 'middle' })
    });
  }

  /*
   * 把需求工時寫進工作表。參數和 planDemand 一樣，回傳的也是同一份結果。
   *
   * 表頭：大表頭（o.title）放在第 2 列、跨所有月份置中；
   *       各月的小表頭放在第 3 列到表頭最後一列，每一行各自上下合併、置中。
   * 資料：每格都是公式，點儲存格看得到是哪幾格相乘。沒有 CT. 的列一樣寫公式（結果是 0），
   *       並標上和 CT. 一樣的淺紅色底，方便辨認。
   */
  function addDemandColumns(main, o) {
    var plan = planDemand(main, o);
    var headerEnd = o.dataStartRow - 1;
    var n = o.monthCols.length;
    var start = plan.startCol;
    var end = start + n - 1;
    var c, r;

    // 重做時，先把原本那一塊清掉（表頭的合併、內容、樣式）
    if (!plan.isNew) {
      var oldEnd = start + Math.max(plan.oldWidth, n) - 1;
      listMerges(main).forEach(function (m) {
        if (m.left >= start && m.right <= oldEnd) main.unMergeCells(m.top, m.left, m.bottom, m.right);
      });
      eachStoredRow(main, function (row) {
        for (var cc = start; cc <= oldEnd; cc++) {
          var cell = row.findCell(cc);
          if (cell) { cell.value = null; cell.style = {}; }
        }
      });
    }

    // 表頭的位置
    var titleRow = headerEnd >= 3 ? 2 : (headerEnd === 2 ? 1 : 0);
    var subTop = headerEnd >= 3 ? 3 : headerEnd;       // 小表頭從第幾列開始（到 headerEnd）

    // 小表頭：樣式照抄該月數量那一行的表頭
    if (headerEnd >= 1) {
      for (var i = 0; i < n; i++) {
        c = start + i;
        for (r = subTop; r <= headerEnd; r++) {
          var srcHead = main.findCell(r, o.monthCols[i]);
          main.getCell(r, c).style = centered(srcHead ? srcHead.style : {});
        }
        main.getCell(subTop, c).value = plan.headers[i];
        if (headerEnd > subTop) main.mergeCellsWithoutStyle(subTop, c, headerEnd, c);
      }
    }
    // 大表頭：跨所有月份置中
    if (titleRow) {
      for (c = start; c <= end; c++) {
        var srcTitle = main.findCell(titleRow, o.monthCols[c - start]);
        main.getCell(titleRow, c).style = centered(srcTitle ? srcTitle.style : {});
      }
      main.getCell(titleRow, start).value = o.title;
      if (end > start) main.mergeCellsWithoutStyle(titleRow, start, titleRow, end);
    }

    // 資料
    plan.rows.forEach(function (item) {
      for (var i = 0; i < n; i++) {
        var cell = main.getCell(item.row, start + i);
        var base = main.findCell(item.row, o.monthCols[i]);
        var style = Object.assign({}, base ? base.style : {});
        style.numFmt = INTEGER_FORMAT;                 // 只顯示整數（儲存格裡的數值仍保留小數）
        if (item.noCt) style.fill = MISSING_FILL;
        cell.style = style;
        cell.value = { formula: item.formulas[i], result: item.values[i] };
      }
    });

    for (c = start; c <= end; c++) main.getColumn(c).width = 12;

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  /* ---------- Dipping_Milling：機台需求工時、目檢需求工時兩塊 ---------- */

  var DEMAND_TITLES = { machine: '機台需求工時', labor: '目檢需求工時' };

  /*
   * 兩塊需求工時：
   *   機台需求工時   各月數量 × 機台 CT.（當月也一樣，不加 WIP）
   *   目檢需求工時   各月數量 × 目檢 CT. × 目檢-檢驗率
   * 位置：機台需求工時從「目檢-檢驗率之後空一行」開始（原本最後一行之後空一行），
   *       目檢需求工時從「機台需求工時最後一個月之後空一行」開始；已經有同名大表頭就覆蓋原本的位置。
   * 只算 Cutting、Dipping 兩段的資料列（o.rows，沒給就自己找）。
   *
   * o: { dataStartRow, keyCol, monthCols, machineCtCol, laborCtCol, laborRateCol, headers, rows }
   * 回傳 { machine: planDemand 的結果, labor: planDemand 的結果, rows }
   */
  function demandOptions(o, which, startCol) {
    return {
      dataStartRow: o.dataStartRow, keyCol: o.keyCol, monthCols: o.monthCols,
      ctCol: which === 'machine' ? o.machineCtCol : o.laborCtCol,
      rateCol: which === 'labor' ? o.laborRateCol : undefined,      // 目檢需求工時再乘目檢-檢驗率
      title: DEMAND_TITLES[which], headers: o.headers, rows: o.rows, startCol: startCol
    };
  }
  function planDemandBlocks(main, o) {
    var opts = Object.assign({}, o);
    if (!opts.rows) opts.rows = sectionDataRows(main, o.dataStartRow - 1);
    var machine = planDemand(main, demandOptions(opts, 'machine'));
    var labor = planDemand(main, demandOptions(opts, 'labor', machine.startCol + o.monthCols.length + 1));
    return { machine: machine, labor: labor, rows: opts.rows };
  }
  function addDemandBlocks(main, o) {
    var opts = Object.assign({}, o);
    if (!opts.rows) opts.rows = sectionDataRows(main, o.dataStartRow - 1);
    var machine = addDemandColumns(main, demandOptions(opts, 'machine'));
    var labor = addDemandColumns(main, demandOptions(opts, 'labor', machine.startCol + o.monthCols.length + 1));
    // 外觀：每一塊的表頭（大表頭 + 月份）一框、每一段一框，字一律黑色，只顯示整數
    var headerEnd = o.dataStartRow - 1;
    var sections = findSections(main, { headerEnd: headerEnd, keyCol: 1 });
    var font = sheetFont(main, headerEnd);
    var n = o.monthCols.length;
    [machine, labor].forEach(function (p) {
      styleBlock(main, {
        headerEnd: headerEnd, left: p.startCol, right: p.startCol + n - 1,
        sections: sections, font: font, numFmt: INTEGER_FORMAT
      });
      p.totals = addBlockTotals(main, p, sections, n, font);
    });
    return { machine: machine, labor: labor, rows: opts.rows };
  }

  /*
   * 需求工時那一塊，每一段（Cutting、Dipping）最後一列下面空一列加一列 Total（和第 2 步的 Total 同一列）：
   *   「Total」放在這一塊左邊那一行（原本空著的那一行），各月 =SUM(P6:P8)
   *   外觀和第 2 步一樣：淺綠底，外框粗線、裡面細線
   * 回傳 [{ key, row, values: [各月合計] }]
   */
  function addBlockTotals(main, plan, sections, n, font) {
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    var labelCol = plan.startCol - 1;
    var right = plan.startCol + n - 1;
    var out = [];
    sections.forEach(function (s) {
      if (!s.count) return;
      var r = s.lastRow + 2;
      var values = [];
      for (var c = labelCol; c <= right; c++) {
        var cell = main.getCell(r, c);
        var st = {
          font: Object.assign({}, font), fill: TOTAL_FILL,
          border: { top: THICK, bottom: THICK, left: c === labelCol ? THICK : THIN, right: c === right ? THICK : THIN }
        };
        if (c === labelCol) {
          st.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.style = st;
          cell.value = 'Total';
          continue;
        }
        var i = c - plan.startCol;
        var sum = 0, err = null;
        plan.rows.forEach(function (it) {
          if (it.row < s.firstRow || it.row > s.lastRow) return;
          var v = it.values[i];
          if (v && typeof v === 'object') { if (!err) err = v.error; } else if (typeof v === 'number') sum += v;
        });
        st.numFmt = INTEGER_FORMAT;
        st.alignment = { vertical: 'middle' };
        cell.style = st;
        var L = numToCol(c);
        cell.value = { formula: 'SUM(' + L + s.firstRow + ':' + L + s.lastRow + ')', result: err ? { error: err } : sum };
        values.push(err ? { error: err } : sum);
      }
      out.push({ key: s.key, row: r, values: values });
    });
    if (main.getColumn(labelCol).width === undefined || main.getColumn(labelCol).width < 9) main.getColumn(labelCol).width = 9;
    return out;
  }

  /* ---------- MPS 匯總：各月預計總產量、前三大、其他 ---------- */

  var TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB5E48C' } };   // 淺綠色（#B5E48C）：Total 那一列的底色

  /*
   * 先算出 MPS 各月的匯總，不動工作表（給畫面預覽用）。
   *
   * o.dataStartRow  資料從第幾列開始
   * o.keyCol        料號在哪一行（用來判斷資料到第幾列、以及顯示前三大是誰）
   * o.wipCol        當月 WIP 在哪一行
   * o.monthCols     各月預計產量在哪幾行（陣列，第一個是當月）
   * o.topN          前幾大（預設 3）
   *
   * 當月（第一個月）的數量一律先加上 WIP：總產量、前三大、標示都是用「WIP + 當月產量」來算。
   *
   * 回傳：
   *   found       有沒有資料
   *   firstRow / lastRow   資料的第一列、最後一列
   *   totalRow / topRow / restRow   三個匯總列的列號（最後一列資料之後空一列開始）
   *   labelCol    文字標籤放在哪一行（第一個月左邊那一行）
   *   months      [{ col, total, top, rest, error, topRows: [{ row, label, value }] }]
   *               topRows 只列數值大於 0 的
   */
  function planMpsTotals(main, o) {
    var lastRow = lastKeyRow(main, o.keyCol, o.dataStartRow);
    if (!lastRow) return { found: false };
    var topN = o.topN || 3;
    var months = o.monthCols.map(function (c, i) {
      var items = [];
      var error = null;
      for (var r = o.dataStartRow; r <= lastRow; r++) {
        var row = main.findRow(r);
        var qty = numberOrError(row ? row.findCell(c) : null);
        var wip = i === 0 ? numberOrError(row ? row.findCell(o.wipCol) : null) : { value: 0 };
        if (qty.error || wip.error) { if (!error) error = qty.error || wip.error; continue; }
        var keyCell = row ? row.findCell(o.keyCol) : null;
        items.push({ row: r, label: keyCell ? displayText(keyCell) : '', value: qty.value + wip.value });
      }
      var total = 0;
      items.forEach(function (it) { total += it.value; });
      var topRows = items.filter(function (it) { return it.value > 0; })
        .sort(function (a, b) { return b.value - a.value || a.row - b.row; }).slice(0, topN);
      var top = 0;
      topRows.forEach(function (it) { top += it.value; });
      return { col: c, total: total, top: top, rest: total - top, error: error, topRows: topRows };
    });
    var firstMonth = Math.min.apply(null, o.monthCols);
    return {
      found: true,
      firstRow: o.dataStartRow,
      lastRow: lastRow,
      totalRow: lastRow + 2,
      topRow: lastRow + 3,
      restRow: lastRow + 4,
      labelCol: firstMonth > 1 ? firstMonth - 1 : 0,
      months: months
    };
  }

  /*
   * 把 MPS 匯總寫進工作表（最後一列資料之後空一列開始，共三列，對齊各月的行），每格都是公式。
   * 以 WIP 在 D 行、當月在 E 行、資料在第 5 到 74 列為例：
   *
   *                 當月（要加 WIP）                              其他月份
   *   Total       =SUM(D5:D74)+SUM(E5:E74)                       =SUM(F5:F74)
   *   Top 3       =SUMPRODUCT(LARGE(D5:D74+E5:E74,{1,2,3}))      =LARGE(F5:F74,1)+LARGE(F5:F74,2)+LARGE(F5:F74,3)
   *   Non-Top 3   =E76-E77                                       =F76-F77
   *
   * 外觀：文字標籤放在第一個月左邊那一行、置中；標籤加上各月圍成一塊，外框粗線、裡面細格線。
   *       Total 那一列整列是淺綠色底（#B5E48C）。
   *
   * 各月前三大的儲存格用「條件式格式」標上淺橘色，數值是 0 的不算：
   *   當月     =AND($D5+E5>0, SUMPRODUCT(--($D$5:$D$74+E$5:E$74>$D5+E5))<3)   比「WIP + 當月」大的不到 3 個
   *   其他月   =AND(F5>0, F5>=LARGE(F$5:F$74,3))
   *
   * o.labels = { total, top, rest }  三列的文字標籤
   * 其餘參數和 planMpsTotals 一樣，回傳的也是同一份結果。
   */
  function addMpsTotals(main, o) {
    var plan = planMpsTotals(main, o);
    if (!plan.found) return plan;
    var first = plan.firstRow, last = plan.lastRow;
    var topN = Math.min(o.topN || 3, last - first + 1);
    var rows = [plan.totalRow, plan.topRow, plan.restRow];
    var texts = [o.labels.total, o.labels.top, o.labels.rest];
    var leftCol = plan.labelCol || Math.min.apply(null, o.monthCols);
    var rightCol = Math.max.apply(null, o.monthCols);
    var inMonth = {};
    o.monthCols.forEach(function (c) { inMonth[c] = true; });

    // 重做時，先把這三列在這一塊裡原本的內容清掉
    rows.forEach(function (r) {
      var row = main.findRow(r);
      if (!row) return;
      eachStoredCell(row, function (cell, c) {
        if (c < leftCol || c > rightCol) return;
        cell.value = null; cell.style = {};
      });
    });

    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    function boxBorder(r, c) {
      return {
        top: r === rows[0] ? THICK : THIN,
        bottom: r === rows[rows.length - 1] ? THICK : THIN,
        left: c === leftCol ? THICK : THIN,
        right: c === rightCol ? THICK : THIN
      };
    }
    var sample = main.findCell(first, o.monthCols[0]);
    var font = sample && sample.style && sample.style.font ? Object.assign({}, sample.style.font) : undefined;
    if (font) delete font.bold;

    var wipL = numToCol(o.wipCol);
    var ks = [];
    for (var k = 1; k <= topN; k++) ks.push(k);

    plan.months.forEach(function (m, i) {
      var L = numToCol(m.col);
      var range = L + first + ':' + L + last;
      var wipRange = wipL + first + ':' + wipL + last;
      var formulas;
      if (i === 0) {
        formulas = [
          'SUM(' + wipRange + ')+SUM(' + range + ')',
          'SUMPRODUCT(LARGE(' + wipRange + '+' + range + ',{' + ks.join(',') + '}))',
          L + plan.totalRow + '-' + L + plan.topRow
        ];
      } else {
        formulas = [
          'SUM(' + range + ')',
          ks.map(function (n) { return 'LARGE(' + range + ',' + n + ')'; }).join('+'),
          L + plan.totalRow + '-' + L + plan.topRow
        ];
      }
      var results = m.error ? [{ error: m.error }, { error: m.error }, { error: m.error }] : [m.total, m.top, m.rest];
      rows.forEach(function (r, j) {
        var cell = main.getCell(r, m.col);
        var style = { numFmt: INTEGER_FORMAT, border: boxBorder(r, m.col), alignment: { vertical: 'middle' } };
        if (font) style.font = font;
        if (j === 0) style.fill = TOTAL_FILL;
        cell.style = style;
        cell.value = { formula: formulas[j], result: results[j] };
      });
    });

    if (plan.labelCol) {
      rows.forEach(function (r, j) {
        var cell = main.getCell(r, plan.labelCol);
        var style = { border: boxBorder(r, plan.labelCol), alignment: { horizontal: 'center', vertical: 'middle' } };
        if (font) style.font = font;
        if (j === 0) style.fill = TOTAL_FILL;
        cell.style = style;
        cell.value = texts[j];
      });
      var labelColumn = main.getColumn(plan.labelCol);
      if (!labelColumn.width || labelColumn.width < 12) labelColumn.width = 12;
    }

    // 前三大標淺橘色：先拿掉之前加在這幾行的規則，再一個月加一條
    main.conditionalFormattings = (main.conditionalFormattings || []).filter(function (cf) {
      var m = /^([A-Z]+)\d+:([A-Z]+)\d+$/.exec(String(cf.ref || ''));
      return !(m && m[1] === m[2] && inMonth[colToNum(m[1])]);
    });
    var fillStyle = { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: TOP_FILL_ARGB } } };
    plan.months.forEach(function (m, i) {
      var L = numToCol(m.col);
      var here = L + first;
      var formula;
      if (i === 0) {
        var both = '$' + wipL + first + '+' + here;
        formula = 'AND(' + both + '>0,SUMPRODUCT(--($' + wipL + '$' + first + ':$' + wipL + '$' + last + '+' +
          L + '$' + first + ':' + L + '$' + last + '>' + both + '))<' + topN + ')';
      } else {
        formula = 'AND(' + here + '>0,' + here + '>=LARGE(' + L + '$' + first + ':' + L + '$' + last + ',' + topN + '))';
      }
      main.addConditionalFormatting({
        ref: L + first + ':' + L + last,
        rules: [{ type: 'expression', priority: nextPriority(main), formulae: [formula], style: fillStyle }]
      });
    });

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  /* ---------- 分段的 MPS（Cutting、Dipping）：找出各段、每段一列 Total ---------- */

  // 第 1 步寫的分段標題，例如「Cutting（Plastic Lens-Cutting）」「Dipping（Plastic Lens-Dipping）」
  var SECTION_TITLE_RE = /^\s*(Cutting|Dipping)\s*[（(]/i;

  /*
   * 在第 1 步匯出的 MPS 工作表裡找出各段。
   * 標題在 keyCol（料號那一行，通常是第 1 行）；標題下面一列起、料號有值的列是這一段的資料，
   * 遇到第一個料號空白的列就結束（所以第 2 步加的 Total 列、空白列都不會算進去）。
   *
   * o: { headerEnd, keyCol }
   * 回傳 [{ key: 'cutting' | 'dipping', title, titleRow, firstRow, lastRow, count }]
   *      count 是 0 時 firstRow > lastRow
   */
  function findSections(ws, o) {
    var keyCol = o.keyCol || 1;
    var last = extent(ws).lastRow;
    var out = [];
    var cur = null;
    for (var r = (o.headerEnd || 0) + 1; r <= last; r++) {
      var row = ws.findRow(r);
      var v = row ? plainValue(cellRawValue(row.findCell(keyCol))) : null;
      var t = v === null || v === undefined ? '' : String(v instanceof Date ? 'x' : v).trim();
      var m = SECTION_TITLE_RE.exec(t);
      if (m) {
        cur = { key: m[1].toLowerCase(), title: t, titleRow: r, firstRow: r + 1, lastRow: r, count: 0, open: true };
        out.push(cur);
        continue;
      }
      if (!cur || !cur.open) continue;
      if (t === '') { cur.open = false; continue; }
      cur.lastRow = r;
      cur.count++;
    }
    return out.map(function (s) {
      return { key: s.key, title: s.title, titleRow: s.titleRow, firstRow: s.firstRow, lastRow: s.lastRow, count: s.count };
    });
  }

  /*
   * 每一段算出各月的 Total，不動工作表（給畫面預覽用）。
   *
   * o.headerEnd   表頭最後一列
   * o.keyCol      料號在哪一行（分段標題也在這一行）
   * o.monthCols   各月產量在哪幾行（第一個是當月；每個月都只加總該月產量，當月不加 WIP）
   *
   * 回傳：
   *   found      有沒有找到分段
   *   labelCol   「Total」放在哪一行（第一個月左邊那一行）
   *   sections   [{ key, title, titleRow, firstRow, lastRow, count, totalRow, months: [{ col, total, error }] }]
   *              totalRow = 這一段最後一列 + 2（空一列）；沒有資料的段 totalRow 是 0（不加 Total）
   */
  function planSectionTotals(main, o) {
    var sections = findSections(main, { headerEnd: o.headerEnd, keyCol: o.keyCol });
    if (!sections.length) return { found: false, sections: [] };
    var firstMonth = Math.min.apply(null, o.monthCols);
    sections.forEach(function (s) {
      s.totalRow = s.count ? s.lastRow + 2 : 0;
      s.months = o.monthCols.map(function (c, i) {
        var total = 0, error = null;
        for (var r = s.firstRow; r <= s.lastRow; r++) {
          var row = main.findRow(r);
          var qty = numberOrError(row ? row.findCell(c) : null);
          if (qty.error) { if (!error) error = qty.error; continue; }
          total += qty.value;
        }
        return { col: c, total: total, error: error };
      });
    });
    return { found: true, labelCol: firstMonth > 1 ? firstMonth - 1 : 0, sections: sections };
  }

  /*
   * 每一段的最後一列下面空一列，加一列 Total（對齊各月的行），每格都是公式。
   * 以當月在 E 行、Cutting 在第 6 到 8 列為例（Total 在第 10 列）：
   *   D10  Total
   *   E10  =SUM(E6:E8)      當月（不加 WIP）
   *   F10  =SUM(F6:F8)      其他月份
   *
   * 外觀：整列淺綠色底（#B5E48C）；「Total」置中；標籤加上各月圍成一塊，外框粗線、裡面細格線。
   * o.label  文字標籤（預設 Total）；其餘參數和 planSectionTotals 一樣，回傳的也是同一份結果。
   */
  function addSectionTotals(main, o) {
    var plan = planSectionTotals(main, o);
    if (!plan.found) return plan;
    var leftCol = plan.labelCol || Math.min.apply(null, o.monthCols);
    var rightCol = Math.max.apply(null, o.monthCols);
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };

    plan.sections.forEach(function (s) {
      if (!s.totalRow) return;
      var r = s.totalRow;
      // 重做時，先把這一列在這一塊裡原本的內容清掉
      var old = main.findRow(r);
      if (old) eachStoredCell(old, function (cell, c) {
        if (c < leftCol || c > rightCol) return;
        cell.value = null; cell.style = {};
      });
      var font = sheetFont(main, o.headerEnd);
      var style = function (c, extra) {
        var st = Object.assign({
          fill: TOTAL_FILL,
          border: { top: THICK, bottom: THICK, left: c === leftCol ? THICK : THIN, right: c === rightCol ? THICK : THIN }
        }, extra);
        if (font) st.font = font;
        return st;
      };
      s.months.forEach(function (m, i) {
        var L = numToCol(m.col);
        var range = L + s.firstRow + ':' + L + s.lastRow;
        var formula = 'SUM(' + range + ')';
        var cell = main.getCell(r, m.col);
        cell.style = style(m.col, { numFmt: INTEGER_FORMAT, alignment: { vertical: 'middle' } });
        cell.value = { formula: formula, result: m.error ? { error: m.error } : m.total };
      });
      // 「Total」標籤；標籤和月份之間如果有沒選到的行，也一起畫進框裡
      for (var c = leftCol; c <= rightCol; c++) {
        if (o.monthCols.indexOf(c) >= 0) continue;
        var cell = main.getCell(r, c);
        cell.style = style(c, c === plan.labelCol ? { alignment: { horizontal: 'center', vertical: 'middle' } } : {});
        cell.value = c === plan.labelCol ? (o.label || 'Total') : null;
      }
    });
    if (plan.labelCol) {
      var labelColumn = main.getColumn(plan.labelCol);
      if (!labelColumn.width || labelColumn.width < 12) labelColumn.width = 12;
    }

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    return plan;
  }

  /* ---------- 需求工時匯總：每月總需求工時、前三大、其他 ---------- */

  var TOP_FILL_ARGB = 'FFFFD8A8';     // 淺橘色

  // 新的條件式格式規則要用的優先順序：接在這張工作表現有規則的後面
  function nextPriority(ws) {
    var max = 0;
    (ws.conditionalFormattings || []).forEach(function (cf) {
      (cf.rules || []).forEach(function (rule) { if (rule.priority > max) max = rule.priority; });
    });
    return max + 1;
  }
  var HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA2D2FF' } };   // 淺藍色（#A2D2FF）：Production、RMA 這兩列的底色
  var LABEL_WIDTH = 26;               // 文字標籤那一行至少要多寬，才放得下「Production 總需求工時」

  // 料號那一行最後一列有內容的列號（匯總列在它下面，不算資料）
  function lastKeyRow(ws, keyCol, dataStartRow) {
    var last = 0;
    eachStoredRow(ws, function (row, r) {
      if (r < dataStartRow) return;
      var cell = row.findCell(keyCol);
      if (cell && normalizeText(displayText(cell)) !== '') last = Math.max(last, r);
    });
    return last;
  }

  /*
   * 先算出匯總結果，不動工作表（給畫面預覽用）。
   *
   * o.dataStartRow  資料從第幾列開始
   * o.keyCol        料號在哪一行（用來判斷資料到第幾列、以及顯示前三大是誰）
   * o.title         需求工時那一塊的大表頭文字，用來找出它在哪幾行
   * o.topN          前幾大（預設 3）
   *
   * 回傳：
   *   found       有沒有找到需求工時那一塊
   *   startCol / width   需求工時從哪一行開始、共幾行（幾個月）
   *   firstRow / lastRow 資料的第一列、最後一列
   *   totalRow / topRow / restRow  三個匯總列的列號（最後一列資料之後空一列開始）
   *   labelCol    匯總列的文字標籤放在哪一行（需求工時左邊那一行）
   *   months      [{ col, header, total, top, rest, error, topRows: [{ row, label, value }] }]
   *               topRows 只列數值大於 0 的
   */
  function planSummary(main, o) {
    var headerEnd = o.dataStartRow - 1;
    var block = findBlock(main, headerEnd, o.title);
    if (!block) return { found: false };
    var lastRow = lastKeyRow(main, o.keyCol, o.dataStartRow);
    if (!lastRow) return { found: true, startCol: block.col, width: block.width, months: [], firstRow: 0, lastRow: 0 };
    var topN = o.topN || 3;
    var subRow = headerEnd >= 3 ? 3 : headerEnd;

    var months = [];
    for (var c = block.col; c < block.col + block.width; c++) {
      var items = [];
      var error = null;
      for (var r = o.dataStartRow; r <= lastRow; r++) {
        var row = main.findRow(r);
        var cell = row ? row.findCell(c) : null;
        var v = numberOrError(cell);
        if (v.error) { if (!error) error = v.error; continue; }
        var keyCell = row ? row.findCell(o.keyCol) : null;
        items.push({ row: r, label: keyCell ? displayText(keyCell) : '', value: v.value });
      }
      var total = 0;
      items.forEach(function (it) { total += it.value; });
      // 前幾大：數值是 0（或負數）的不算，免得一堆 0 也被當成前三大
      var sorted = items.filter(function (it) { return it.value > 0; })
        .sort(function (a, b) { return b.value - a.value || a.row - b.row; });
      var topRows = sorted.slice(0, topN);
      var top = 0;
      topRows.forEach(function (it) { top += it.value; });
      var headCell = subRow >= 1 ? main.findCell(subRow, c) : null;
      months.push({
        col: c, header: headCell ? displayText(headCell) : '',
        total: total, top: top, rest: total - top, error: error, topRows: topRows
      });
    }

    return {
      found: true,
      startCol: block.col,
      width: block.width,
      firstRow: o.dataStartRow,
      lastRow: lastRow,
      totalRow: lastRow + 2,
      topRow: lastRow + 3,
      restRow: lastRow + 4,
      labelCol: block.col > 1 ? block.col - 1 : 0,
      months: months
    };
  }

  /*
   * 把匯總寫進工作表（最後一列資料之後空一列開始，共三列），每格都是公式：
   *   Production 總需求工時   =SUM(N5:N74)
   *   Top 3 總需求工時        =LARGE(N5:N74,1)+LARGE(N5:N74,2)+LARGE(N5:N74,3)
   *   Others 總需求工時       =N76-N77          （Production − Top 3）
   *
   * 外觀：文字標籤放在需求工時左邊那一行、置中；標籤加上各月共七行、三列圍成一塊，
   *       外框粗線、裡面細格線。Production 那一列整列是淺藍色底（#A2D2FF）。
   *
   * 每個月前三大的儲存格用「條件式格式」標上淺橘色：數值改了，標示會跟著變。
   * 規則是「大於 0，而且不小於該月第三大的數值」，所以 0 不會被標示。
   *
   * o.labels = { total, top, rest }  三列的文字標籤
   * 其餘參數和 planSummary 一樣，回傳的也是同一份結果。
   */
  function addSummaryRows(main, o) {
    var plan = planSummary(main, o);
    if (!plan.found || !plan.lastRow) return plan;
    var dataRows = plan.lastRow - plan.firstRow + 1;
    var topN = Math.min(o.topN || 3, dataRows);
    var endCol = plan.startCol + plan.width - 1;
    var leftCol = plan.labelCol || plan.startCol;       // 整塊最左邊那一行
    var rows = [plan.totalRow, plan.topRow, plan.restRow];
    var texts = [o.labels.total, o.labels.top, o.labels.rest];

    // 重做時，先把這三列原本的內容清掉。只清這一塊的行，同一列左邊可能有 MPS 的匯總。
    rows.forEach(function (r) {
      var row = main.findRow(r);
      if (!row) return;
      eachStoredCell(row, function (cell, c) {
        if (c < leftCol || c > endCol) return;
        cell.value = null; cell.style = {};
      });
    });

    // 框線：最外圍粗線，裡面細線
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    function boxBorder(r, c) {
      return {
        top: r === rows[0] ? THICK : THIN,
        bottom: r === rows[rows.length - 1] ? THICK : THIN,
        left: c === leftCol ? THICK : THIN,
        right: c === endCol ? THICK : THIN
      };
    }
    // 字型跟資料列一樣（不加粗）
    var sample = main.findCell(plan.lastRow, plan.startCol);
    var font = sample && sample.style && sample.style.font ? Object.assign({}, sample.style.font) : undefined;
    if (font) delete font.bold;

    plan.months.forEach(function (m) {
      var L = numToCol(m.col);
      var range = L + plan.firstRow + ':' + L + plan.lastRow;
      var large = [];
      for (var k = 1; k <= topN; k++) large.push('LARGE(' + range + ',' + k + ')');
      var formulas = [
        'SUM(' + range + ')',
        large.join('+'),
        L + plan.totalRow + '-' + L + plan.topRow
      ];
      var results = m.error ? [{ error: m.error }, { error: m.error }, { error: m.error }] : [m.total, m.top, m.rest];
      rows.forEach(function (r, i) {
        var cell = main.getCell(r, m.col);
        var style = { numFmt: INTEGER_FORMAT, border: boxBorder(r, m.col), alignment: { vertical: 'middle' } };
        if (font) style.font = font;
        if (i === 0) style.fill = HEAD_FILL;                   // Production 那一列
        cell.style = style;
        cell.value = { formula: formulas[i], result: results[i] };
      });
    });

    // 文字標籤：需求工時左邊那一行，置中；行寬不夠就加寬到放得下
    if (plan.labelCol) {
      rows.forEach(function (r, i) {
        var cell = main.getCell(r, plan.labelCol);
        var style = { border: boxBorder(r, plan.labelCol), alignment: { horizontal: 'center', vertical: 'middle' } };
        if (font) style.font = font;
        if (i === 0) style.fill = HEAD_FILL;
        cell.style = style;
        cell.value = texts[i];
      });
      var labelColumn = main.getColumn(plan.labelCol);
      if (!labelColumn.width || labelColumn.width < LABEL_WIDTH) labelColumn.width = LABEL_WIDTH;
    }

    // 前三大標淺橘色：先拿掉之前加在這一塊的規則，再加一條涵蓋所有月份的規則。
    // 公式寫的是左上角那一格（例如 N5），Excel 會自動套用到範圍裡的每一格：
    //   =AND(N5>0, N5>=LARGE(N$5:N$74,3))    大於 0，而且不小於這個月第三大的數值
    var blockRef = function (ref) {
      var m = /^([A-Z]+)\d+:([A-Z]+)\d+$/.exec(String(ref || ''));
      return !!m && colToNum(m[1]) >= plan.startCol && colToNum(m[2]) <= endCol;
    };
    main.conditionalFormattings = (main.conditionalFormattings || []).filter(function (cf) {
      return !blockRef(cf.ref);
    });
    var firstL = numToCol(plan.startCol), lastL = numToCol(endCol);
    main.addConditionalFormatting({
      ref: firstL + plan.firstRow + ':' + lastL + plan.lastRow,
      rules: [{
        type: 'expression', priority: nextPriority(main),
        formulae: ['AND(' + firstL + plan.firstRow + '>0,' + firstL + plan.firstRow + '>=LARGE(' +
          firstL + '$' + plan.firstRow + ':' + firstL + '$' + plan.lastRow + ',' + topN + '))'],
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: TOP_FILL_ARGB } } }
      }]
    });

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    return plan;
  }

  /* ---------- RMA 需求工時：使用者自己輸入的項目 ---------- */

  /*
   * 先找出 RMA 區塊要放在哪裡，不動工作表。
   *
   * 參數和 planSummary 一樣，另外：
   *   o.summaryMark   匯總第一列標籤裡一定會有的字（用來確認第 6 步做過了），例如 '總需求工時'
   *
   * 回傳：
   *   found       找不找得到放的位置；找不到時 reason = 'demand'（沒有需求工時）、
   *               'rows'（沒有資料）、'summary'（還沒做匯總）
   *   startRow    RMA 區塊的第一列（匯總最後一列之後空一列）
   *   labelCol    項目名稱放在哪一行（需求工時左邊那一行）
   *   ctCol       CT. 放在哪一行（項目名稱左邊那一行）
   *   startCol / width   各月從哪一行開始、共幾個月
   *   headers     各月的表頭文字
   */
  function planRma(main, o) {
    var sp = planSummary(main, o);
    if (!sp.found) return { found: false, reason: 'demand' };
    if (!sp.lastRow) return { found: false, reason: 'rows' };
    var mark = sp.labelCol ? main.findCell(sp.totalRow, sp.labelCol) : null;
    if (!mark || normalizeText(displayText(mark)).indexOf(o.summaryMark) < 0) {
      return { found: false, reason: 'summary' };
    }
    return {
      found: true,
      startRow: sp.restRow + 2,
      restRow: sp.restRow,
      labelCol: sp.labelCol,
      ctCol: sp.labelCol > 1 ? sp.labelCol - 1 : 0,
      startCol: sp.startCol,
      width: sp.width,
      headers: sp.months.map(function (m) { return m.header; })
    };
  }

  /*
   * 把 RMA 區塊寫進工作表（匯總下面空一列）：
   *
   *   第一列   RMA總需求工時    =SUM(N81:N83)        底下各項目的合計
   *   項目列   項目名稱         =120*L81             這個月的重工數量 × 左邊那一格的 CT.
   *
   * 項目名稱左邊那一格放 CT.（數字），第一列的同一個位置寫「CT.」當作標題。
   * 數量是使用者在畫面上輸入的，直接寫在公式裡；之後在 Excel 改 CT. 那一格，工時會跟著變。
   * 外觀和匯總那一塊一樣：外框粗線、裡面細格線；第一列的標籤置中，項目名稱靠左。
   * 第一列（RMA總需求工時）整列是淺藍色底（#A2D2FF），左邊的「CT.」標題不上色。
   *
   * o.items  = [{ name, ct, qty: [各月數量] }]
   * o.labels = { title: 'RMA總需求工時', ct: 'CT.' }
   * 其餘參數和 planRma 一樣，回傳的也是同一份結果（多了 firstItemRow、lastRow）。
   */
  function addRmaRows(main, o) {
    var plan = planRma(main, o);
    if (!plan.found) return plan;
    var items = o.items || [];
    var top = plan.startRow;
    var first = top + 1;
    var last = top + items.length;
    var endCol = plan.startCol + plan.width - 1;

    // 重做時，把匯總下面原本的內容清掉（這幾行從這一列以下都是 RMA）
    var clearFrom = plan.ctCol || plan.labelCol;
    eachStoredRow(main, function (row, r) {
      if (r < top) return;
      eachStoredCell(row, function (cell, c) {
        if (c < clearFrom || c > endCol) return;
        cell.value = null; cell.style = {};
      });
    });

    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    function boxBorder(r, c) {
      return {
        top: r === top ? THICK : THIN,
        bottom: r === last ? THICK : THIN,
        left: c === plan.labelCol ? THICK : THIN,
        right: c === endCol ? THICK : THIN
      };
    }
    var sample = main.findCell(o.dataStartRow, plan.startCol);
    var font = sample && sample.style && sample.style.font ? Object.assign({}, sample.style.font) : undefined;
    if (font) delete font.bold;
    function styled(extra) {
      var st = Object.assign({}, extra);
      if (font) st.font = font;
      return st;
    }

    // 第一列：標籤、CT. 標題、各月合計
    var titleCell = main.getCell(top, plan.labelCol);
    titleCell.style = styled({ border: boxBorder(top, plan.labelCol), alignment: { horizontal: 'center', vertical: 'middle' }, fill: HEAD_FILL });
    titleCell.value = o.labels.title;
    if (plan.ctCol) {
      var ctHead = main.getCell(top, plan.ctCol);
      ctHead.style = styled({ alignment: { horizontal: 'center', vertical: 'middle' } });
      ctHead.value = o.labels.ct;
    }
    var ctL = plan.ctCol ? numToCol(plan.ctCol) : '';
    for (var c = plan.startCol; c <= endCol; c++) {
      var i = c - plan.startCol;
      var L = numToCol(c);
      var total = 0;
      items.forEach(function (it) { total += (it.qty[i] || 0) * it.ct; });
      var sumCell = main.getCell(top, c);
      sumCell.style = styled({ numFmt: INTEGER_FORMAT, border: boxBorder(top, c), alignment: { vertical: 'middle' }, fill: HEAD_FILL });
      sumCell.value = items.length
        ? { formula: 'SUM(' + L + first + ':' + L + last + ')', result: total }
        : 0;
    }

    // 項目列
    items.forEach(function (it, k) {
      var r = first + k;
      var nameCell = main.getCell(r, plan.labelCol);
      nameCell.style = styled({ border: boxBorder(r, plan.labelCol), alignment: { horizontal: 'left', vertical: 'middle' } });
      nameCell.value = it.name;
      if (plan.ctCol) {
        var ctCell = main.getCell(r, plan.ctCol);
        ctCell.style = styled({ numFmt: 'General', alignment: { vertical: 'middle' } });
        ctCell.value = it.ct;
      }
      for (var c = plan.startCol; c <= endCol; c++) {
        var qty = it.qty[c - plan.startCol] || 0;
        var cell = main.getCell(r, c);
        cell.style = styled({ numFmt: INTEGER_FORMAT, border: boxBorder(r, c), alignment: { vertical: 'middle' } });
        cell.value = plan.ctCol
          ? { formula: String(qty) + '*' + ctL + r, result: qty * it.ct }
          : qty * it.ct;
      }
    });

    var labelColumn = main.getColumn(plan.labelCol);
    if (!labelColumn.width || labelColumn.width < LABEL_WIDTH) labelColumn.width = LABEL_WIDTH;

    var wb = main.workbook;
    if (wb) wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

    plan.firstItemRow = first;
    plan.lastRow = last;
    return plan;
  }

  /* ---------- 人力報表：各項目在哪一列、當月在哪一行 ---------- */

  // 之後總表要抓的六個項目。key 是程式裡用的名字，label 是畫面上顯示的文字。
  var MANPOWER_ITEMS = [
    { key: 'online', label: 'Online' },
    { key: 'onlineProd', label: 'Online（生產）' },
    { key: 'onlineRework', label: 'Online（重工）' },
    { key: 'offDirect', label: 'Off line（直接）' },
    { key: 'offIndirect', label: 'Off line（間接）' },
    { key: 'ttl', label: 'TTL（On + Off直接）' }
  ];

  // 比對列標題用：全形轉半形、轉小寫、去掉空白和連字號（Off line、Off-line、OFFLINE 都一樣）
  function labelKey(text) {
    return normalizeText(text).toLowerCase().replace(/[\s\-_]/g, '');
  }

  /*
   * 列出人力報表每一列的文字，給使用者選「這個項目在第幾列」。
   * 回傳 [{ row, text, key }]，只列有內容的列：
   *   text  這一列的文字（數字不算；合併儲存格用左上角的字），最多取前三段
   *   key   比對用（labelKey 處理過的 text）
   */
  function manpowerRows(ws, ext) {
    var out = [];
    for (var r = 1; r <= ext.lastRow; r++) {
      var row = ws.findRow(r);
      if (!row) continue;
      var parts = [];
      var any = false;
      for (var c = 1; c <= ext.colCount; c++) {
        var cell = row.findCell(c);
        if (!cell) continue;
        var v = plainValue(cellRawValue(cell));
        if (v === null || v === undefined || v === '') continue;
        any = true;
        if (typeof v !== 'string') continue;
        var t = normalizeText(v).replace(/\s+/g, ' ');
        var asNumber = toNumber(t);
        if (t === '' || (asNumber !== null && !isNaN(asNumber))) continue;   // 文字型態的數字不算標題
        if (parts.indexOf(t) < 0 && parts.length < 3) parts.push(t);
      }
      if (!any) continue;
      var text = parts.join(' ');
      out.push({ row: r, text: text, key: labelKey(text) });
    }
    return out;
  }

  // 依列標題猜六個項目各在哪一列。回傳 { online: 5, onlineProd: 6, … }，猜不到的是 0。
  function guessManpowerRows(rows) {
    function has(k, word) { return k.indexOf(word) >= 0; }
    function isTotal(k) { return has(k, 'ttl') || has(k, 'total') || has(k, '+'); }
    var tests = {
      online: [
        function (k) { return k === 'online'; },
        function (k) { return has(k, 'online') && !has(k, '生產') && !has(k, '重工') && !has(k, 'off') && !isTotal(k); }
      ],
      onlineProd: [function (k) { return has(k, 'online') && has(k, '生產') && !has(k, 'off') && !isTotal(k); }],
      onlineRework: [function (k) { return has(k, 'online') && has(k, '重工') && !has(k, 'off') && !isTotal(k); }],
      offDirect: [function (k) { return has(k, 'offline') && has(k, '直接') && !has(k, '間接') && !isTotal(k); }],
      offIndirect: [function (k) { return has(k, 'offline') && has(k, '間接') && !isTotal(k); }],
      ttl: [
        function (k) { return (has(k, 'ttl') || has(k, 'total')) && has(k, 'on') && has(k, 'off'); },
        function (k) { return has(k, 'ttl'); }
      ]
    };
    var found = {};
    MANPOWER_ITEMS.forEach(function (item) {
      found[item.key] = 0;
      tests[item.key].some(function (test) {
        var hit = rows.filter(function (x) { return x.key !== '' && test(x.key); })[0];
        if (hit) found[item.key] = hit.row;
        return !!hit;
      });
    });
    return found;
  }

  // 第 r 列裡是數字的行（由左到右）
  function numericColumns(ws, r, colCount) {
    var cols = [];
    var row = ws.findRow(r);
    if (!row) return cols;
    for (var c = 1; c <= colCount; c++) {
      var n = toNumber(cellRawValue(row.findCell(c)));
      if (n !== null && !isNaN(n)) cols.push(c);
    }
    return cols;
  }

  // 猜當月在哪一行：把選好的各列裡是數字的行合在一起，剛好是 months 個連續的行時取第一個；
  // 其他情況（例如連過去的月份也列出來）不猜，回傳 0，讓使用者自己選。
  function guessFirstMonthCol(ws, rows, colCount, months) {
    var seen = {};
    rows.forEach(function (r) {
      numericColumns(ws, r, colCount).forEach(function (c) { seen[c] = true; });
    });
    var cols = Object.keys(seen).map(Number).sort(function (x, y) { return x - y; });
    if (cols.length !== months) return 0;
    return cols[cols.length - 1] - cols[0] === months - 1 ? cols[0] : 0;
  }

  // 每一行的表頭文字：從 aboveRow 的上一列往上找，第一個有內容的格子。回傳 labels[行號]
  function labelsAbove(ws, aboveRow, colCount) {
    var labels = [];
    for (var c = 1; c <= colCount; c++) {
      labels[c] = '';
      for (var r = aboveRow - 1; r >= 1; r--) {
        var row = ws.findRow(r);
        var cell = row ? row.findCell(c) : null;
        var t = cell ? normalizeText(displayText(cell)).replace(/\s+/g, ' ') : '';
        if (t !== '') { labels[c] = t; break; }
      }
    }
    return labels;
  }

  /*
   * 依使用者選的列和行，把之後總表要抓的數值讀出來給畫面核對。
   *   rows      { online: 5, … }
   *   firstCol  當月在哪一行
   *   months    幾個月
   * 回傳 { cols: [行號…], values: { online: [{ text, number, blank, bad }…], … }, blank, bad }
   */
  function readManpower(ws, rows, firstCol, months) {
    var out = { cols: [], values: {}, blank: 0, bad: 0 };
    for (var i = 0; i < months; i++) out.cols.push(firstCol + i);
    MANPOWER_ITEMS.forEach(function (item) {
      var r = rows[item.key];
      var row = r ? ws.findRow(r) : null;
      out.values[item.key] = out.cols.map(function (c) {
        var cell = row ? row.findCell(c) : null;
        var n = toNumber(cellRawValue(cell));
        var info = { text: cell ? displayText(cell) : '', number: n, blank: n === null, bad: n !== null && isNaN(n) };
        if (info.blank) out.blank++;
        if (info.bad) out.bad++;
        return info;
      });
    });
    return out;
  }

  /* ---------- 總表：先建立表格的樣式和框線，內容之後再填 ---------- */

  var REPORT_SHEET = '總表';
  var REPORT = {
    firstCol: 2,           // B：項目名稱
    unitCol: 3,            // C：單位、子項目
    monthCol: 4,           // D：第一個月
    months: 6,             // D 到 I
    noteCol: 10,           // J：備註
    noteRow: 2,            // 第 2 列：法定加班上限
    headRow: 3,            // 第 3 列：月份表頭
    lastRow: 24
  };
  var REPORT_FONT = '微軟正黑體';
  var OVERTIME_LIMIT = 46;     // 法定加班上限（小時/人/月）
  var WEEKDAY_OT_HOURS = 2;    // 平日每天加班 2 小時
  var DAY_HOURS = 8;           // 一天工作 8 小時
  var YO_LIMIT_DEFAULT = 78;   // 第二張表：YO 加班上限（小時/人/月），畫面上可以改
  var REPORT_SECOND_OFFSET = 26;   // 第二張表在第一張下面空三列（第一張是第 2–24 列，第二張從第 28 列開始）
  // 第二張表直接參照第一張表的列（其他列是第二張表自己算）
  var REPORT_SECOND_COPY = ['section', 'hoursTop3', 'hoursNonTop3', 'hoursRework', 'weekdays', 'workMinutes',
    'overtimeMinutes', 'effTop3', 'effNonTop3', 'attendance', 'needOffDirect', 'needOffIndirect',
    'haveOnline', 'haveOffDirect', 'haveOffIndirect'];
  // 需求人力 Off line 的人員種類（之後要增減種類改這裡）
  var REPORT_OFFLINE = {
    offDirect: ['文書', '物料', '材製損', '包裝', '領班'],
    offIndirect: ['物料', '領班']
  };
  function isNum(v) { return typeof v === 'number' && !isNaN(v); }
  var REPORT_COLOR = {
    red: 'FFFF0000',       // 法定加班上限
    head: 'FF366092',      // 表頭深藍
    section: 'FFDCE6F1',   // MPS (Production) 那一列
    subtotal: 'FFEBF1DE',  // 小計
    loading: 'FFFFCC99',   // Loading Rate
    navy: 'FF000080',      // 人力相關的字
    white: 'FFFFFFFF'
  };
  // 每一列：key 給之後填內容用；label 在 B，unit 在 C；merge = B、C 合併
  var REPORT_ROWS = [
    { row: 4, key: 'section', label: 'MPS (Production)', merge: true, section: true },
    { row: 5, key: 'hoursTop3', label: '需求工時(Top 3)', unit: 'min/月' },
    { row: 6, key: 'hoursNonTop3', label: '需求工時(Non-Top 3)', unit: 'min/月' },
    { row: 7, key: 'hoursRework', label: '需求工時（重工）', unit: 'min/月' },
    { row: 8, key: 'weekdays', label: '平日工作天數', unit: '天/月' },
    { row: 9, key: 'holidays', label: '假日工作天數', unit: '天/月' },
    { row: 10, key: 'workMinutes', label: '工作時間', unit: '分/天' },
    { row: 11, key: 'overtimeMinutes', label: '加班時間', unit: '分/天' },
    { row: 12, key: 'effTop3', label: '綜合效率%（Top-3 機種）', merge: true },
    { row: 13, key: 'effNonTop3', label: '綜合效率%（Non-Top 3 機種）', merge: true },
    { row: 14, key: 'attendance', label: '出勤率%', merge: true },
    { row: 15, key: 'needOnline', group: '需求人力', unit: 'On line', navy: true },
    { row: 16, key: 'needOffDirect', unit: 'Off line_直接', navy: true },
    { row: 17, key: 'needOffIndirect', unit: 'Off line_間接', navy: true },
    { row: 18, key: 'needSubtotal', unit: '小計', navy: true, subtotal: true },
    { row: 19, key: 'haveOnline', group: '現有人力\n(早+夜)', unit: 'On line', navy: true },
    { row: 20, key: 'haveOffDirect', unit: 'Off line_直接', navy: true },
    { row: 21, key: 'haveOffIndirect', unit: 'Off line_間接', navy: true },
    { row: 22, key: 'haveSubtotal', unit: '小計(on+off)', navy: true, subtotal: true },
    { row: 23, key: 'gap', label: '人力需求差異(現有-需求)', merge: true, navy: true },
    { row: 24, key: 'loading', label: 'Loading Rate%(含加班)', merge: true, navy: true, loading: true }
  ];

  /*
   * 總表各列要從 MPS 工作表哪一列抓數值：用列的文字標籤找。
   *   Total               第 2 步匯總（當月不含 WIP）
   *   Top 3 總需求工時     第 6 步
   *   Others 總需求工時    第 6 步
   *   RMA總需求工時        第 7 步
   */
  var REPORT_LINKS = [
    { key: 'section', label: 'Total', step: 2 },
    { key: 'hoursTop3', label: 'Top 3 總需求工時', step: 6 },
    { key: 'hoursNonTop3', label: 'Others 總需求工時', step: 6 },
    { key: 'hoursRework', label: 'RMA總需求工時', step: 7 }
  ];

  function labelMatchKey(v) {
    return normalizeText(v).replace(/\s+/g, ' ').toLowerCase();
  }

  /*
   * 在 MPS 工作表找出總表各列的來源。
   *   o.dataStartRow  資料從第幾列開始（標籤只在這一列以下找，避免找到表頭的字）
   *   o.months        幾個月（6）
   * 回傳 { section: { found, row, labelCol, refs: [{ sheet, col, row, value }] }, … }
   * 數值是標籤右邊的格子，由左往右取 months 個（空白略過，碰到文字就停）。
   * 同一個標籤出現很多次時，取最下面那一個（匯總都在資料的下面）。
   */
  function findReportLinks(main, o) {
    var ext = extent(main);
    var out = {};
    REPORT_LINKS.forEach(function (link) {
      var want = labelMatchKey(link.label);
      var hit = null;
      eachStoredRow(main, function (row, r) {
        if (r < (o.dataStartRow || 1)) return;
        eachStoredCell(row, function (cell, c) {
          if (cell.isMerged && cell.master !== cell) return;
          var v = plainValue(cellRawValue(cell));
          if (typeof v === 'string' && labelMatchKey(v) === want) hit = { row: r, col: c };
        });
      });
      if (!hit) { out[link.key] = { found: false, label: link.label, step: link.step, refs: [] }; return; }
      var refs = [];
      var row = main.findRow(hit.row);
      for (var c = hit.col + 1; c <= ext.colCount && refs.length < o.months; c++) {
        var cell = row.findCell(c);
        var raw = cell ? cellRawValue(cell) : null;
        var v = plainValue(raw);
        if (v === null || v === undefined || v === '') {
          var isFormula = raw && typeof raw === 'object' && (raw.formula !== undefined || raw.sharedFormula !== undefined);
          if (!isFormula) { if (refs.length) break; else continue; }
        }
        if (typeof v === 'string' && isNaN(toNumber(v))) break;
        var n = toNumber(v);
        refs.push({ sheet: main.name, col: c, row: hit.row, value: n === null || isNaN(n) ? null : n });
      }
      out[link.key] = { found: true, label: link.label, step: link.step, row: hit.row, labelCol: hit.col, refs: refs };
    });
    return out;
  }

  function solid(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb: argb } }; }

  /*
   * 在活頁簿裡建立「總表」工作表（已經有的話先刪掉重建）。
   * o.monthRefs  各月表頭要參照的儲存格 [{ sheet: 'MPS', col: 14, row: 3, text: 'Oct' }]；
   *              有給的話表頭填公式（例如 ='MPS'!N3），沒給的那幾格留空白
   * o.links      各列要抓的儲存格 { section: [{ sheet, col, row, value }], hoursTop3: […], … }（findReportLinks 的 refs）；
   *              有給的話那一列的 D 到 I 填公式（例如 ='MPS'!E76）
   * o.inputs     使用者在畫面上輸入的數值：
   *                weekdays     [6 個月的平日工作天數]（null = 沒填）
   *                workMinutes  [6 個月的工作時間（分/天）]
   *                effTop3 / effNonTop3 / attendance   [6 個月的 %]（例如 80 代表 80%，存成 0.8）
   *                otHours      每日加班幾小時
   *                otRate       加班出勤人數 %（例如 80 代表 80%）
   *                offDirect / offIndirect   各種人員的人數，順序同 REPORT_OFFLINE（null = 沒填）
   * o.yoLimit    第二張表（第 28 列起）的 YO 加班上限，預設 78；第二張表的假日工作天數 =(78-D34*2)/8，
   *              其他輸入和抓來的數值參照第一張表（例如 D34 =D8）
   *              假日工作天數一律填公式 =(46-D8*2)/8；加班時間填公式 =2*60*80%（兩個數字都有才填），備註寫輸入的數值
   * 回傳 { worksheet, rows: { hoursTop3: 5, … }, monthCol, months }
   */
  function addReportSheet(wb, o) {
    o = o || {};
    wb.worksheets
      .filter(function (w) { return sameSheetName(w.name, REPORT_SHEET); })
      .forEach(function (w) { wb.removeWorksheet(w.id); });
    var ws = wb.addWorksheet(REPORT_SHEET);
    var R = REPORT;
    var C = REPORT_COLOR;
    var lastMonth = R.monthCol + R.months - 1;
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    var LEFT = { horizontal: 'left', vertical: 'middle' };
    var CENTER = { horizontal: 'center', vertical: 'middle' };
    var yoLimit = isNum(o.yoLimit) ? o.yoLimit : YO_LIMIT_DEFAULT;
    var OFF2 = REPORT_SECOND_OFFSET;                       // 第二張表往下移幾列
    var bottom = R.lastRow + OFF2;

    // 整張總表都用微軟正黑體（charset 136 = 繁體中文）
    function font(extra) { return Object.assign({ name: REPORT_FONT, size: 12, charset: 136 }, extra); }

    // 行寬；表格外的空白格子也設成微軟正黑體，之後在旁邊打字字型才一致
    ws.getColumn(1).width = 6;
    ws.getColumn(R.firstCol).width = 24;
    ws.getColumn(R.unitCol).width = 17;    // 要放得下「Off line_直接」「小計(on+off)」，不然字會跑到 D 行
    for (var mc = R.monthCol; mc <= lastMonth; mc++) ws.getColumn(mc).width = 10;
    ws.getColumn(R.noteCol).width = 56;
    for (var fc = 1; fc <= R.noteCol + 1; fc++) ws.getColumn(fc).font = font();
    for (var fr = 1; fr <= bottom + 1; fr++) {
      for (fc = 1; fc <= R.noteCol + 1; fc++) ws.getCell(fr, fc).font = font();
    }

    /* ----- 畫一張表的外框、底色、標籤（off = 往下移幾列），回傳各列的列號 ----- */
    function drawFrame(off, title) {
      var noteRow = R.noteRow + off, headRow = R.headRow + off, lastRow = R.lastRow + off;
      ws.getRow(noteRow).height = 24;
      ws.getRow(headRow).height = 21;
      for (var hr = headRow + 1; hr <= lastRow; hr++) ws.getRow(hr).height = 18;

      // 外框粗線、裡面細線；Loading Rate 那一列上面也是粗線
      function border(r, c) {
        return {
          top: r === headRow || r === lastRow ? THICK : THIN,
          bottom: r === lastRow || r === lastRow - 1 ? THICK : THIN,
          left: c === R.firstCol ? THICK : THIN,
          right: c === R.noteCol ? THICK : THIN
        };
      }

      // 加班上限（紅底白字，B、C 合併）
      ws.getCell(noteRow, R.firstCol).value = title;
      ws.mergeCells(noteRow, R.firstCol, noteRow, R.unitCol);
      ws.getCell(noteRow, R.firstCol).style = {
        font: font({ bold: true, color: { argb: C.white } }), fill: solid(C.red), alignment: LEFT
      };
      ws.getCell(noteRow, R.unitCol).style = { font: font(), fill: solid(C.red) };

      // 表頭（深藍底白字；B、C 合併，畫一條斜線）
      for (var c = R.firstCol; c <= R.noteCol; c++) {
        var b = border(headRow, c);
        if (c <= R.unitCol) b.diagonal = { up: false, down: true, style: 'medium', color: { argb: 'FF000000' } };
        ws.getCell(headRow, c).style = {
          font: font({ size: 14, bold: true, color: { argb: C.white } }),
          fill: solid(C.head), alignment: CENTER, border: b
        };
      }
      ws.mergeCellsWithoutStyle(headRow, R.firstCol, headRow, R.unitCol);
      ws.getCell(headRow, R.noteCol).value = '備註';

      // 各列
      var rows = { noteRow: noteRow, headRow: headRow };
      var groupTop = 0;
      REPORT_ROWS.forEach(function (d, k) {
        var r = d.row + off;
        rows[d.key] = r;
        var fill = d.section ? solid(C.section) : d.loading ? solid(C.loading) : null;
        var color = d.navy ? { argb: C.navy } : undefined;
        for (var c = R.firstCol; c <= R.noteCol; c++) {
          var f = fill;
          if (d.subtotal && c >= R.unitCol) f = solid(C.subtotal);
          var st = { border: border(r, c), alignment: c <= R.unitCol ? LEFT : CENTER, font: font({ color: color }) };
          if (f) st.fill = f;
          if (c <= R.unitCol) st.font = font({ bold: true, color: color, size: d.section ? 14 : 12 });
          ws.getCell(r, c).style = st;
        }
        if (d.label) ws.getCell(r, R.firstCol).value = d.label;
        if (d.unit) ws.getCell(r, R.unitCol).value = d.unit;
        if (d.merge) ws.mergeCellsWithoutStyle(r, R.firstCol, r, R.unitCol);

        // 需求人力、現有人力：B 行往下合併到這一組的最後一列
        if (d.group) {
          groupTop = r;
          var g = ws.getCell(r, R.firstCol);
          g.value = d.group;
          g.style = Object.assign({}, g.style, {
            font: font({ color: color }),
            alignment: { horizontal: 'left', vertical: 'middle', wrapText: true }
          });
        }
        var nextRow = REPORT_ROWS[k + 1];
        if (groupTop && (!nextRow || nextRow.group || !nextRow.unit || nextRow.label)) {
          if (r > groupTop) ws.mergeCellsWithoutStyle(groupTop, R.firstCol, r, R.firstCol);
          groupTop = 0;
        }
      });
      return rows;
    }

    function cellOf(rows, key, m) { return ws.getCell(rows[key], R.monthCol + m); }
    function addr(rows, key, m) { return numToCol(R.monthCol + m) + rows[key]; }
    function num(rows, key, m) {
      var v = plainValue(cellRawValue(cellOf(rows, key, m)));
      return typeof v === 'number' ? v : (v === null || v === undefined ? 0 : NaN);
    }
    function setFormula(cell, formula, result, fmt) {
      cell.value = isNum(result) && isFinite(result) ? { formula: formula, result: result } : { formula: formula };
      cell.numFmt = fmt;
    }
    function setNote(rows, key, text) {
      var c = ws.getCell(rows[key], R.noteCol);
      c.value = text;
      c.alignment = { horizontal: 'left', vertical: 'middle' };
    }

    /* ----- 由其他列算出來的列（兩張表都一樣，只有加班上限不同） ----- */
    function fillComputed(rows, limit) {
      for (var m = 0; m < R.months; m++) {
        var a = function (key) { return addr(rows, key, m); };
        var n = function (key) { return num(rows, key, m); };

        // 假日工作天數 =（加班上限 − 平日工作天數 × 每日加班 2 小時）÷ 每天 8 小時
        setFormula(cellOf(rows, 'holidays', m),
          '(' + limit + '-' + a('weekdays') + '*' + WEEKDAY_OT_HOURS + ')/' + DAY_HOURS,
          (limit - n('weekdays') * WEEKDAY_OT_HOURS) / DAY_HOURS, '0.00');

        // 需求人力 On line：
        //   Top 3     需求工時 ÷（平日+假日天數）÷（工作時間+加班時間）÷ Top-3 效率 ÷ 出勤率
        //   Non-Top 3 需求工時 ÷（平日+假日天數）÷（工作時間+加班時間）÷ Non-Top 3 效率 ÷ 出勤率
        //   重工      需求工時 ÷ 平日天數 ÷ 工作時間 ÷ Non-Top 3 效率 ÷ 出勤率
        var days = '(' + a('weekdays') + '+' + a('holidays') + ')';
        var mins = '(' + a('workMinutes') + '+' + a('overtimeMinutes') + ')';
        var f = a('hoursTop3') + '/' + days + '/' + mins + '/' + a('effTop3') + '/' + a('attendance') + '+' +
          a('hoursNonTop3') + '/' + days + '/' + mins + '/' + a('effNonTop3') + '/' + a('attendance') + '+' +
          a('hoursRework') + '/' + a('weekdays') + '/' + a('workMinutes') + '/' + a('effNonTop3') + '/' + a('attendance');
        var d = n('weekdays') + n('holidays');
        var t = n('workMinutes') + n('overtimeMinutes');
        var res = n('hoursTop3') / d / t / n('effTop3') / n('attendance') +
          n('hoursNonTop3') / d / t / n('effNonTop3') / n('attendance') +
          n('hoursRework') / n('weekdays') / n('workMinutes') / n('effNonTop3') / n('attendance');
        setFormula(cellOf(rows, 'needOnline', m), f, res, '#,##0.0');

        // 需求人力小計 = On line + Off line_直接（不含 Off line_間接）
        setFormula(cellOf(rows, 'needSubtotal', m), a('needOnline') + '+' + a('needOffDirect'),
          n('needOnline') + n('needOffDirect'), '#,##0.0');
        // 現有人力小計 = On line + Off line_直接（不含 Off line_間接）
        setFormula(cellOf(rows, 'haveSubtotal', m), a('haveOnline') + '+' + a('haveOffDirect'),
          n('haveOnline') + n('haveOffDirect'), 'General');
        // 人力需求差異 = 現有人力小計 − 需求人力小計；Loading Rate = 現有人力小計 ÷ 需求人力小計
        var hv = n('haveSubtotal'), nv = n('needSubtotal');
        setFormula(cellOf(rows, 'gap', m), a('haveSubtotal') + '-' + a('needSubtotal'), hv - nv, '#,##0.0');
        setFormula(cellOf(rows, 'loading', m), a('haveSubtotal') + '/' + a('needSubtotal'), nv ? hv / nv : NaN, '0%');
      }
      setNote(rows, 'needSubtotal', '不含Off line_間接');
      setNote(rows, 'haveSubtotal', '不含Off line_間接');
    }

    /* ===== 第一張表：法定加班上限 46 小時 ===== */
    var rows = drawFrame(0, '法定加班上限' + OVERTIME_LIMIT + 'hr/人');

    // 月份表頭：參照 MPS 工作表需求工時的月份
    var refs = o.monthRefs || [];
    for (var i = 0; i < R.months; i++) {
      var ref = refs[i];
      if (ref && ref.sheet && ref.col && ref.row) {
        ws.getCell(rows.headRow, R.monthCol + i).value = {
          formula: sheetRef(ref.sheet) + numToCol(ref.col) + ref.row,
          result: ref.text || ''
        };
      }
    }

    // 從其他工作表抓的列（MPS、人力）
    var links = o.links || {};
    Object.keys(links).forEach(function (key) {
      if (!rows[key]) return;
      (links[key] || []).slice(0, R.months).forEach(function (ref, i) {
        var f = { formula: sheetRef(ref.sheet) + numToCol(ref.col) + ref.row };
        if (ref.value !== null && ref.value !== undefined) f.result = ref.value;
        var cell = cellOf(rows, key, i);
        cell.value = f;
        cell.numFmt = ref.numFmt || INTEGER_FORMAT;
      });
    });

    // 使用者輸入的列
    var inp = o.inputs || {};
    for (var m = 0; m < R.months; m++) {
      var wd = inp.weekdays ? inp.weekdays[m] : null;
      if (isNum(wd)) {
        cellOf(rows, 'weekdays', m).value = wd;
        cellOf(rows, 'weekdays', m).numFmt = 'General';
      }
      var wm = inp.workMinutes ? inp.workMinutes[m] : null;
      if (isNum(wm)) {
        cellOf(rows, 'workMinutes', m).value = wm;
        cellOf(rows, 'workMinutes', m).numFmt = INTEGER_FORMAT;
      }
      // 綜合效率、出勤率：輸入 80 存成 0.8，格式顯示成 80%
      ['effTop3', 'effNonTop3', 'attendance'].forEach(function (key) {
        var v = inp[key] ? inp[key][m] : null;
        if (!isNum(v)) return;
        var pc = cellOf(rows, key, m);
        pc.value = Math.round(v * 1e6) / 1e8;
        pc.numFmt = Math.round(v) === v ? '0%' : '0.0%';
      });
      // 加班時間（分/天）= 每日加班小時 × 60 × 加班出勤人數 %
      if (isNum(inp.otHours) && isNum(inp.otRate)) {
        setFormula(cellOf(rows, 'overtimeMinutes', m), inp.otHours + '*60*' + inp.otRate + '%',
          inp.otHours * 60 * inp.otRate / 100, 'General');
      }
    }
    if (isNum(inp.otHours) || isNum(inp.otRate)) {
      setNote(rows, 'overtimeMinutes', '每日加班 ' + (isNum(inp.otHours) ? inp.otHours : '?') +
        ' 小時 × 60 × 加班出勤人數 ' + (isNum(inp.otRate) ? inp.otRate : '?') + '%');
    }

    // 需求人力 Off line_直接、Off line_間接：各種人員的人數相加（每個月一樣），備註寫各種人數
    [['needOffDirect', 'offDirect'], ['needOffIndirect', 'offIndirect']].forEach(function (pair) {
      var key = pair[0], cats = REPORT_OFFLINE[pair[1]], counts = inp[pair[1]] || [];
      if (!counts.some(isNum)) return;
      var nums = cats.map(function (c, i) { return isNum(counts[i]) ? counts[i] : 0; });
      var total = nums.reduce(function (x, y) { return x + y; }, 0);
      for (var m = 0; m < R.months; m++) setFormula(cellOf(rows, key, m), nums.join('+'), total, 'General');
      setNote(rows, key, cats.map(function (c, i) { return c + ' ' + nums[i]; }).join('、'));
    });

    fillComputed(rows, OVERTIME_LIMIT);

    /* ===== 第二張表：YO 加班上限（預設 78 小時，可在畫面上改） =====
     * 和第一張表一樣，只有假日工作天數改用 YO 的加班上限；
     * 其他輸入、抓來的數值都參照第一張表（例如 =D8），上面改了下面會跟著變。 */
    var rows2 = drawFrame(OFF2, 'YO加班上限' + yoLimit + 'hr/人');
    for (i = 0; i < R.months; i++) {
      var head1 = ws.getCell(rows.headRow, R.monthCol + i);
      if (head1.value === null || head1.value === undefined) continue;
      ws.getCell(rows2.headRow, R.monthCol + i).value = {
        formula: numToCol(R.monthCol + i) + rows.headRow,
        result: plainValue(cellRawValue(head1)) || ''
      };
    }
    REPORT_SECOND_COPY.forEach(function (key) {
      for (var m = 0; m < R.months; m++) {
        var src = cellOf(rows, key, m);
        var v = plainValue(cellRawValue(src));
        if (v === null || v === undefined || v === '') continue;    // 上面空白，下面也留空白
        setFormula(cellOf(rows2, key, m), addr(rows, key, m), typeof v === 'number' ? v : NaN, src.numFmt || 'General');
      }
      var note = ws.getCell(rows[key], R.noteCol).value;
      if (note) setNote(rows2, key, note);
    });
    fillComputed(rows2, yoLimit);

    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    return { worksheet: ws, rows: rows, rows2: rows2, monthCol: R.monthCol, months: R.months, yoLimit: yoLimit };
  }

  /* ---------- RMA人力需求：先建立表格的樣式和框線，內容之後再填 ---------- */

  var RMA_SHEET = 'RMA人力需求';
  var RMA_LAYOUT = {
    firstCol: 2, unitCol: 3, monthCol: 4, months: 6, noteCol: 10,
    noteRow: 3,            // 第 3 列：法定加班上限（左）、Cut-off（右）
    headRow: 4,            // 第 4 列：月份表頭
    lastRow: 13
  };
  var RMA_ROWS = [
    { row: 5, key: 'section', label: 'MPS (PCS)', merge: true, section: true },
    { row: 6, key: 'hours', label: '需求工時', unit: 'min/月' },
    { row: 7, key: 'weekdays', label: '平日工作天數', unit: '天/月' },
    { row: 8, key: 'holidays', label: '假日工作天數', unit: '天/月' },
    { row: 9, key: 'workMinutes', label: '工作時間', unit: '分/天' },
    { row: 10, key: 'overtimeMinutes', label: '加班時間', unit: '分/天' },
    { row: 11, key: 'efficiency', label: '綜合效率%', merge: true },
    { row: 12, key: 'attendance', label: '出勤率%', merge: true },
    { row: 13, key: 'needOnline', label: '需求人力', unit: 'On line', navy: true, subtotal: true }
  ];

  /*
   * 在活頁簿裡建立「RMA人力需求」工作表（已經有的話先刪掉重建）。
   * o.cutoff     右上角的 Cut-off 文字，例如 'Cut-off: 2026/10/13Ver:A'
   * o.monthRefs  各月表頭要參照的儲存格 [{ sheet, col, row, text }]（參照總表的月份）
   * o.links      各列要參照的儲存格 { hours: [{ sheet, col, row, value, numFmt }], weekdays: […], … }
   * o.qty        MPS (PCS) 那一列：各月 RMA 數量 [[120, 300], [80, 0], …]，填公式 =120+300
   * 回傳 { worksheet, rows: { hours: 6, … }, monthCol, months }
   */
  function addRmaManpowerSheet(wb, o) {
    o = o || {};
    wb.worksheets
      .filter(function (w) { return sameSheetName(w.name, RMA_SHEET); })
      .forEach(function (w) { wb.removeWorksheet(w.id); });
    var ws = wb.addWorksheet(RMA_SHEET);
    var R = RMA_LAYOUT;
    var C = REPORT_COLOR;
    var THIN = { style: 'thin' }, THICK = { style: 'medium' };
    var LEFT = { horizontal: 'left', vertical: 'middle' };
    var CENTER = { horizontal: 'center', vertical: 'middle' };
    function font(extra) { return Object.assign({ name: REPORT_FONT, size: 12, charset: 136 }, extra); }

    // 行寬、列高；表格外的空白格子也設成微軟正黑體
    ws.getColumn(1).width = 3.5;
    ws.getColumn(R.firstCol).width = 15;
    ws.getColumn(R.unitCol).width = 11;
    for (var mc = R.monthCol; mc < R.monthCol + R.months; mc++) ws.getColumn(mc).width = 11;
    ws.getColumn(R.noteCol).width = 30;
    for (var fc = 1; fc <= R.noteCol + 1; fc++) ws.getColumn(fc).font = font();
    for (var fr = 1; fr <= R.lastRow + 1; fr++) {
      for (fc = 1; fc <= R.noteCol + 1; fc++) ws.getCell(fr, fc).font = font();
      ws.getRow(fr).height = fr === R.noteRow ? 18 : 21;
    }

    function border(r, c) {
      return {
        top: r === R.headRow ? THICK : THIN,
        bottom: r === R.lastRow ? THICK : THIN,
        left: c === R.firstCol ? THICK : THIN,
        right: c === R.noteCol ? THICK : THIN
      };
    }

    // 第 3 列：法定加班上限（紅底白字置中，B、C 合併）；右邊 J 欄是 Cut-off
    ws.getCell(R.noteRow, R.firstCol).value = '法定加班上限' + OVERTIME_LIMIT + 'hr/人';
    ws.mergeCells(R.noteRow, R.firstCol, R.noteRow, R.unitCol);
    ws.getCell(R.noteRow, R.firstCol).style = {
      font: font({ bold: true, color: { argb: C.white } }), fill: solid(C.red), alignment: CENTER
    };
    ws.getCell(R.noteRow, R.unitCol).style = { font: font(), fill: solid(C.red) };
    var cut = ws.getCell(R.noteRow, R.noteCol);
    cut.value = o.cutoff || '';
    cut.style = { font: font({ bold: true, size: 11 }), alignment: { horizontal: 'left', vertical: 'bottom' } };

    // 第 4 列：表頭（深藍底白字；B、C 合併畫斜線；備註靠左）
    for (var c = R.firstCol; c <= R.noteCol; c++) {
      var b = border(R.headRow, c);
      if (c <= R.unitCol) b.diagonal = { up: false, down: true, style: 'medium', color: { argb: 'FF000000' } };
      ws.getCell(R.headRow, c).style = {
        font: font({ size: 14, bold: true, color: { argb: C.white } }),
        fill: solid(C.head), alignment: c === R.noteCol ? LEFT : CENTER, border: b
      };
    }
    ws.mergeCellsWithoutStyle(R.headRow, R.firstCol, R.headRow, R.unitCol);
    (o.monthRefs || []).slice(0, R.months).forEach(function (ref, i) {
      if (!ref || !ref.sheet) return;
      ws.getCell(R.headRow, R.monthCol + i).value = {
        formula: sheetRef(ref.sheet) + numToCol(ref.col) + ref.row, result: ref.text || ''
      };
    });
    ws.getCell(R.headRow, R.noteCol).value = '備註';

    // 第 5 到 13 列
    var rows = { noteRow: R.noteRow, headRow: R.headRow };
    RMA_ROWS.forEach(function (d) {
      rows[d.key] = d.row;
      var fill = d.section ? solid(C.section) : d.subtotal ? solid(C.subtotal) : null;
      var color = d.navy ? { argb: C.navy } : undefined;
      for (var c = R.firstCol; c <= R.noteCol; c++) {
        var st = { border: border(d.row, c), alignment: c <= R.unitCol || c === R.noteCol ? LEFT : CENTER, font: font({ color: color }) };
        if (fill) st.fill = fill;
        if (c <= R.unitCol) st.font = font({ bold: !(d.navy && c === R.firstCol), color: color, size: d.section ? 14 : 12 });
        ws.getCell(d.row, c).style = st;
      }
      if (d.label) ws.getCell(d.row, R.firstCol).value = d.label;
      if (d.unit) ws.getCell(d.row, R.unitCol).value = d.unit;
      if (d.merge) ws.mergeCellsWithoutStyle(d.row, R.firstCol, d.row, R.unitCol);
    });

    // MPS (PCS)：各月 RMA 項目的數量相加
    (o.qty || []).slice(0, R.months).forEach(function (list, m) {
      if (!list || !list.length) return;
      var total = list.reduce(function (x, y) { return x + y; }, 0);
      var cell = ws.getCell(rows.section, R.monthCol + m);
      cell.value = { formula: list.join('+'), result: total };
      cell.numFmt = INTEGER_FORMAT;
    });
    // 參照其他工作表的列（MPS 的 RMA總需求工時、總表的輸入）
    var links = o.links || {};
    Object.keys(links).forEach(function (key) {
      if (!rows[key]) return;
      (links[key] || []).slice(0, R.months).forEach(function (ref, m) {
        if (!ref) return;
        var f = { formula: sheetRef(ref.sheet) + numToCol(ref.col) + ref.row };
        if (isNum(ref.value)) f.result = ref.value;
        var cell = ws.getCell(rows[key], R.monthCol + m);
        cell.value = f;
        cell.numFmt = ref.numFmt || 'General';
      });
    });

    // 需求人力 On line = 需求工時 ÷（平日 + 假日工作天數）÷（工作時間 + 加班時間）÷ 綜合效率% ÷ 出勤率%
    // 顯示到小數點第二位
    function rnum(key, m) {
      var v = plainValue(cellRawValue(ws.getCell(rows[key], R.monthCol + m)));
      return typeof v === 'number' ? v : (v === null || v === undefined || v === '' ? 0 : NaN);
    }
    for (var m = 0; m < R.months; m++) {
      var L = numToCol(R.monthCol + m);
      var a = function (key) { return L + rows[key]; };
      var f = a('hours') + '/(' + a('weekdays') + '+' + a('holidays') + ')/(' + a('workMinutes') + '+' +
        a('overtimeMinutes') + ')/' + a('efficiency') + '/' + a('attendance');
      var res = rnum('hours', m) / (rnum('weekdays', m) + rnum('holidays', m)) /
        (rnum('workMinutes', m) + rnum('overtimeMinutes', m)) / rnum('efficiency', m) / rnum('attendance', m);
      var on = ws.getCell(rows.needOnline, R.monthCol + m);
      on.value = isFinite(res) ? { formula: f, result: res } : { formula: f };
      on.numFmt = '#,##0.00';
    }

    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
    return { worksheet: ws, rows: rows, monthCol: R.monthCol, months: R.months };
  }

  /*
   * 從 MPS 工作表第 7 步的 RMA 區塊讀出各項目各月的數量。
   * 項目那幾列的公式是 =120*L81（數量 × CT.），數量就是 * 前面的數字。
   *   link  findReportLinks(...).hoursRework（RMA總需求工時那一列的位置和各月的行）
   * 回傳 { items: ['鏡片重工', …], qty: [[120, 300], …（每個月一組）], bad: 讀不懂的格子數 }
   */
  function readRmaQuantities(main, link) {
    var out = { items: [], qty: [], bad: 0 };
    if (!link || !link.found) return out;
    var cols = link.refs.map(function (r) { return r.col; });
    cols.forEach(function () { out.qty.push([]); });
    for (var r = link.row + 1; ; r++) {
      var row = main.findRow(r);
      var label = row ? row.findCell(link.labelCol) : null;
      var name = label ? normalizeText(displayText(label)) : '';
      if (!name) break;
      out.items.push(name);
      cols.forEach(function (c, i) {
        var v = cellRawValue(row.findCell(c));
        var f = v && typeof v === 'object' && v.formula !== undefined ? String(v.formula) : null;
        var m = f ? /^\s*(-?\d+(?:\.\d+)?)\s*\*/.exec(f) : null;
        if (m) out.qty[i].push(Number(m[1]));
        else out.bad++;
      });
    }
    return out;
  }

  function sameSheetName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }

  /* ---------- 總表（Dipping_Milling 版） ---------- */

  /*
   * 版面：B 到 K 行、第 2 到 34 列。
   *   B、C：項目（沒有子項目時 B:C 合併）；D：單位；E 到 J：六個月；K：Notes
   *   第 2 列表頭：B:C 合併畫斜線，深藍底白字
   * 各列的 key 給之後填公式用。fill 是整列（B 到 K）的底色。
   * 目前只建立格式、文字、框線；各格的數值之後再一項一項補上。
   */
  var DM_REPORT = {
    sheet: '總表',
    labelCol: 2, subCol: 3, unitCol: 4, monthCol: 5, months: 6, noteCol: 11,
    headRow: 2, firstRow: 3, lastRow: 34
  };
  var DM_FILL = {
    head: 'FF3D5D8F',      // 表頭深藍
    mps: 'FFDAEEF3',       // MPS 淺藍
    milling: 'FFEBF1DE',   // Milling 站 淺綠
    dipping: 'FFE4DFEC',   // Dipping 站 淺紫
    labor: 'FFDDD9C4',     // 目檢 淺褐
    total: 'FFFCD5B4'      // 總需求人力、總現有人力、Loading Rate 淺橘
  };
  // b：B 行（或 B:C 合併）的文字；c：C 行的子項目；u：單位
  var DM_ROWS = [
    { r: 3, key: 'workDays', b: '工作天數', u: 'day' },
    { r: 4, key: 'workMinutes', b: '工作時間', u: 'min/班' },
    { r: 5, key: 'machineMinutes', b: '機台每日可動工時', u: 'min/天' },
    { r: 6, key: 'shifts', b: '工作班別數', u: '班' },
    { r: 7, key: 'attendance', b: '出勤率', u: '%' },
    { r: 8, key: 'efficiency', b: '作業效率', u: '%' },
    { r: 9, key: 'mpsMilling', c: 'Milling', fill: 'mps' },
    { r: 10, key: 'mpsDipping', c: 'Dipping', fill: 'mps' },
    { r: 11, key: 'millingHours', c: '機台需求工時', u: 'min', fill: 'milling' },
    { r: 12, key: 'millingRatio', c: '人機比', u: '台/人', fill: 'milling' },
    { r: 13, key: 'millingMachines', c: '開機數', u: '台/人', fill: 'milling' },
    { r: 14, key: 'millingTenders', c: '顧機人數', u: '人/班', fill: 'milling' },
    { r: 15, key: 'millingOtDays', b: '加班天數 - Milling', u: '天', fill: 'milling' },
    { r: 16, key: 'millingNeed', b: '需求顧機人數 - Milling', u: '人/天', fill: 'milling' },
    { r: 17, key: 'dippingHours', c: '機台需求工時', u: 'min', fill: 'dipping' },
    { r: 18, key: 'dippingRatio', c: '人機比', u: '台/人', fill: 'dipping' },
    { r: 19, key: 'dippingMachines', c: '開機數', u: '台/人', fill: 'dipping' },
    { r: 20, key: 'dippingTenders', c: '顧機人數', u: '人/班', fill: 'dipping' },
    { r: 21, key: 'dippingOtDays', b: '加班天數 - Dipping', u: '天', fill: 'dipping' },
    { r: 22, key: 'dippingNeed', b: '需求顧機人數 - dipping', u: '人/天', fill: 'dipping' },
    { r: 23, key: 'laborMinMilling', c: '需求工時 - Milling', fill: 'labor' },
    { r: 24, key: 'laborMinDipping', c: '需求工時 - Dipping', fill: 'labor' },
    { r: 25, key: 'laborPplMilling', c: '需求人力 - Milling', fill: 'labor' },
    { r: 26, key: 'laborPplDipping', c: '需求人力 - Dipping', fill: 'labor' },
    { r: 27, key: 'laborOtDays', b: '加班天數 - 目檢', u: '天', fill: 'labor' },
    { r: 28, key: 'laborNeed', b: '需求人數 - 目檢', u: '人/天', fill: 'labor' },
    { r: 29, key: 'needTotal', b: '總需求人力', u: '人/天', fill: 'total' },
    { r: 30, key: 'haveMilling', c: 'Milling' },
    { r: 31, key: 'haveDipping', c: 'Dipping' },
    { r: 32, key: 'haveTotal', b: '總現有人力', u: '人/天', fill: 'total', thickBottom: true },
    { r: 33, key: 'gap', b: '人力需求差異', u: '人/天' },
    { r: 34, key: 'loading', b: 'Loading Rate', u: '%', fill: 'total' }
  ];
  // B 行上下合併的群組標籤（align: 'center' 置中；'left' 靠左、垂直置中；'top' 靠左上）
  var DM_GROUPS = [
    { top: 9, bottom: 10, text: 'MPS', align: 'center' },
    { top: 11, bottom: 14, text: 'Milling站', align: 'center' },
    { top: 17, bottom: 20, text: 'Dipping站', align: 'center' },
    { top: 23, bottom: 26, text: '目檢站', align: 'center' },
    { top: 30, bottom: 31, text: '現有人力', align: 'center' }
  ];
  // D 行（單位）上下合併
  var DM_UNIT_MERGES = [
    { top: 9, bottom: 10, text: 'PCS' },
    { top: 23, bottom: 24, text: 'min/月' },
    { top: 25, bottom: 26, text: '人/班' },
    { top: 30, bottom: 31, text: '人/天' }
  ];

  /*
   * MPS 工作表某一塊需求工時（例如「機台需求工時」）在每一段的各月合計要填的公式：
   *   有第 5 步加的 Total 列（該段最後一列之後空一列、這一塊左邊那一行寫 Total）：直接參照，例如 ='MPS'!P10
   *   沒有：直接加總該段，例如 =SUM('MPS'!P6:P8)
   * o: { headerEnd, title, months }
   * 回傳 { found, cutting: [{ formula, result, fromTotal }…] 或 null, dipping: 同上 }
   */
  function dmBlockTotalFormulas(main, o) {
    var out = { found: false, cutting: null, dipping: null };
    var block = findBlock(main, o.headerEnd, o.title);
    if (!block) return out;
    out.found = true;
    var ref = sheetRef(main.name);
    findSections(main, { headerEnd: o.headerEnd, keyCol: 1 }).forEach(function (s) {
      if (out[s.key] || !s.count) return;
      var totalRow = s.lastRow + 2;
      var lab = block.col > 1 ? main.findCell(totalRow, block.col - 1) : null;
      var hasTotal = !!lab && lookupKey(lab.value) === lookupKey('Total');
      var list = [];
      for (var i = 0; i < o.months; i++) {
        var c = block.col + i;
        var L = numToCol(c);
        var sum = 0, err = null;
        for (var r = s.firstRow; r <= s.lastRow; r++) {
          var v = numberOrError(main.findCell(r, c));
          if (v.error) { if (!err) err = v.error; } else sum += v.value;
        }
        list.push({
          formula: hasTotal ? ref + L + totalRow : 'SUM(' + ref + L + s.firstRow + ':' + L + s.lastRow + ')',
          result: err ? { error: err } : sum, fromTotal: hasTotal
        });
      }
      out[s.key] = list;
    });
    return out;
  }

  /*
   * 總表 MPS 那兩列（Milling = 第 1 步的 Cutting、Dipping）要填的公式：該段各月的 MPS 合計。
   *   有第 2 步的 Total 列（該段最後一列之後空一列、標籤是 Total）：直接參照它，例如 ='MPS'!E10
   *   沒有：直接加總該段，例如 =SUM('MPS'!E6:E8)（當月也不加 WIP）
   * o: { headerEnd, monthCols }
   * 回傳 { found, cutting: [{ formula, result, fromTotal }…6] 或 null, dipping: 同上, sections }
   */
  function dmMpsFormulas(main, o) {
    var plan = planSectionTotals(main, { headerEnd: o.headerEnd, keyCol: 1, monthCols: o.monthCols });
    var out = { found: plan.found, cutting: null, dipping: null, sections: plan.sections };
    if (!plan.found) return out;
    var ref = sheetRef(main.name);
    plan.sections.forEach(function (s) {
      if (out[s.key] || !s.count) return;
      var hasTotal = false;
      if (plan.labelCol) {
        var lab = main.findCell(s.totalRow, plan.labelCol);
        hasTotal = !!lab && lookupKey(lab.value) === lookupKey('Total');
      }
      out[s.key] = s.months.map(function (m, i) {
        var L = numToCol(m.col);
        var formula;
        if (hasTotal) formula = ref + L + s.totalRow;
        else {
          var range = L + s.firstRow + ':' + L + s.lastRow;
          formula = 'SUM(' + ref + range + ')';
        }
        return { formula: formula, result: m.error ? { error: m.error } : m.total, fromTotal: hasTotal };
      });
    });
    return out;
  }

  /*
   * 在活頁簿裡建立 Dipping_Milling 版的總表（已經有同名工作表就換掉）。
   * o.monthRefs  [{ sheet, col, row, text }] 月份表頭要參照的儲存格（MPS 的機台需求工時月份）；沒有就留空白
   * o.mps        dmMpsFormulas() 的結果：MPS Milling、Dipping 兩列各月的公式
   * o.millingHours  dmBlockTotalFormulas() 的 cutting：Milling站 機台需求工時各月的公式
   * o.millingRatio  { machines, people } Milling站 人機比的機台數、人員數
   * o.dippingHours、o.dippingRatio   Dipping站，同上
   * o.laborHours  dmBlockTotalFormulas()（目檢需求工時）的結果：目檢站 需求工時 - Milling、- Dipping 兩列
   * o.inputs     使用者在畫面上輸入的值（沒填的是 null，那一格留空白）：
   *   workDays      [6 個月] 工作天數
   *   workHours、restHours   工作時數、休息時數（小時，六個月共用）→ 工作時間 =(工作時數-休息時數)*60（分鐘）
   *   machineHours  [6 個月] 機台每日可動工時（小時）→ =小時*60（分鐘）
   *   shifts        [6 個月] 工作班別數
   *   attendance、efficiency  [6 個月] 出勤率、作業效率（小數，0.95 = 95%）
   * 回傳 { worksheet, rows: { key: 列號 }, monthCol, months }
   */
  function addDmReportSheet(wb, o) {
    var R = DM_REPORT;
    wb.worksheets.filter(function (w) { return sameSheetName(w.name, R.sheet); })
      .forEach(function (w) { wb.removeWorksheet(w.id); });
    var ws = wb.addWorksheet(R.sheet, { views: [{ state: 'normal', showGridLines: true }] });
    var lastMonth = R.monthCol + R.months - 1;
    var THIN = { style: 'thin', color: { argb: 'FF000000' } };
    var THICK = { style: 'medium', color: { argb: 'FF000000' } };
    function font(extra) { return Object.assign({ name: REPORT_FONT, size: 11 }, extra || {}); }
    var LEFT = { horizontal: 'left', vertical: 'middle' };
    var CENTER = { horizontal: 'center', vertical: 'middle' };
    var TOP_LEFT = { horizontal: 'left', vertical: 'top' };
    var thickBottomRows = {};
    DM_ROWS.forEach(function (x) { if (x.thickBottom) thickBottomRows[x.r] = true; });

    function border(r, c) {
      return {
        top: r === R.headRow ? THICK : THIN,
        bottom: r === R.headRow || r === R.lastRow || thickBottomRows[r] ? THICK : THIN,
        left: c === R.labelCol ? THICK : THIN,
        right: c === R.noteCol ? THICK : THIN
      };
    }

    // 行寬、列高
    var widths = { 1: 3, 2: 24, 3: 20, 4: 10, 11: 40 };
    for (var c = 1; c <= R.noteCol; c++) ws.getColumn(c).width = widths[c] || 13;
    ws.getRow(1).height = 10;
    ws.getRow(R.headRow).height = 24;
    for (var r = R.firstRow; r <= R.lastRow; r++) ws.getRow(r).height = 21;

    // 表頭
    for (c = R.labelCol; c <= R.noteCol; c++) {
      var b = border(R.headRow, c);
      if (c <= R.subCol) b.diagonal = { up: false, down: true, style: 'medium', color: { argb: 'FF000000' } };
      ws.getCell(R.headRow, c).style = {
        font: font({ size: 12, bold: true, color: { argb: 'FFFFFFFF' } }),
        fill: solid(DM_FILL.head), alignment: CENTER, border: b
      };
    }
    ws.mergeCellsWithoutStyle(R.headRow, R.labelCol, R.headRow, R.subCol);
    ws.getCell(R.headRow, R.unitCol).value = 'Unit';
    ws.getCell(R.headRow, R.noteCol).value = 'Notes';
    for (var m = 0; m < R.months; m++) {
      var ref = o && o.monthRefs ? o.monthRefs[m] : null;
      var hc = ws.getCell(R.headRow, R.monthCol + m);
      hc.value = ref ? { formula: sheetRef(ref.sheet) + numToCol(ref.col) + ref.row, result: ref.text || '' } : null;
    }

    // 內容：每一格先給底色、框線、字型
    var rows = {};
    DM_ROWS.forEach(function (x) {
      rows[x.key] = x.r;
      var fill = x.fill ? solid(DM_FILL[x.fill]) : undefined;
      for (var cc = R.labelCol; cc <= R.noteCol; cc++) {
        var isLabel = cc <= R.subCol;
        var st = {
          font: font(isLabel ? { bold: true } : {}),
          alignment: cc === R.unitCol ? CENTER : (isLabel || cc === R.noteCol ? LEFT : { horizontal: 'right', vertical: 'middle' }),
          border: border(x.r, cc)
        };
        if (fill) st.fill = fill;
        ws.getCell(x.r, cc).style = st;
      }
      if (x.b) {
        ws.getCell(x.r, R.labelCol).value = x.b;
        ws.mergeCellsWithoutStyle(x.r, R.labelCol, x.r, R.subCol);
      }
      if (x.c) ws.getCell(x.r, R.subCol).value = x.c;
      if (x.u) ws.getCell(x.r, R.unitCol).value = x.u;
    });
    DM_GROUPS.forEach(function (g) {
      var cell = ws.getCell(g.top, R.labelCol);
      cell.value = g.text;
      cell.style = Object.assign({}, cell.style, { alignment: g.align === 'top' ? TOP_LEFT : g.align === 'left' ? LEFT : CENTER });
      ws.mergeCellsWithoutStyle(g.top, R.labelCol, g.bottom, R.labelCol);
    });
    DM_UNIT_MERGES.forEach(function (u) {
      ws.getCell(u.top, R.unitCol).value = u.text;
      ws.mergeCellsWithoutStyle(u.top, R.unitCol, u.bottom, R.unitCol);
    });

    // MPS：Milling、Dipping 各月的合計（公式，由 dmMpsFormulas() 算好傳進來）
    var mps = (o && o.mps) || {};
    [['cutting', 'mpsMilling'], ['dipping', 'mpsDipping']].forEach(function (p) {
      var list = mps[p[0]];
      if (!list) return;
      for (var mm = 0; mm < R.months && mm < list.length; mm++) {
        if (!list[mm]) continue;
        var cell = ws.getCell(rows[p[1]], R.monthCol + mm);
        cell.value = { formula: list[mm].formula, result: list[mm].result };
        cell.style = Object.assign({}, cell.style, { numFmt: '#,##0' });
      }
    });

    // 使用者輸入的值
    var inp = (o && o.inputs) || {};
    function isN(v) { return typeof v === 'number' && !isNaN(v); }
    function fmtNum(v) { return String(Math.round(v * 1e6) / 1e6); }
    function put(key, mm, value, numFmt) {
      var cell = ws.getCell(rows[key], R.monthCol + mm);
      cell.value = value;
      if (numFmt) cell.style = Object.assign({}, cell.style, { numFmt: numFmt });
    }
    for (m = 0; m < R.months; m++) {
      var wd = inp.workDays && inp.workDays[m];
      if (isN(wd)) put('workDays', m, wd, 'General');
      if (isN(inp.workHours) && isN(inp.restHours)) {
        put('workMinutes', m, {
          formula: '(' + fmtNum(inp.workHours) + '-' + fmtNum(inp.restHours) + ')*60',
          result: (inp.workHours - inp.restHours) * 60
        }, '#,##0');
      }
      var mh = inp.machineHours && inp.machineHours[m];
      if (isN(mh)) put('machineMinutes', m, { formula: fmtNum(mh) + '*60', result: mh * 60 }, '#,##0');
      var sh = inp.shifts && inp.shifts[m];
      if (isN(sh)) put('shifts', m, sh, 'General');
      var at = inp.attendance && inp.attendance[m];
      if (isN(at)) put('attendance', m, at, '0.0%');
      var ef = inp.efficiency && inp.efficiency[m];
      if (isN(ef)) put('efficiency', m, ef, '0.0%');
    }
    if (isN(inp.workHours) && isN(inp.restHours)) {
      ws.getCell(rows.workMinutes, R.noteCol).value =
        '工作 ' + fmtNum(inp.workHours) + ' 小時 − 休息 ' + fmtNum(inp.restHours) + ' 小時，× 60';
    }
    if (inp.machineHours && inp.machineHours.some(isN)) {
      ws.getCell(rows.machineMinutes, R.noteCol).value = '輸入的小時 × 60';
    }

    // 目檢站 需求工時：參照 MPS 工作表該段「目檢需求工時」的 Total 列（由 dmBlockTotalFormulas() 算好傳進來）
    var lh = (o && o.laborHours) || {};
    [['cutting', 'laborMinMilling'], ['dipping', 'laborMinDipping']].forEach(function (p) {
      var list = lh[p[0]];
      if (!list) return;
      for (var mm = 0; mm < R.months && mm < list.length; mm++) {
        if (list[mm]) put(p[1], mm, { formula: list[mm].formula, result: list[mm].result }, '#,##0');
      }
    });

    // Milling站、Dipping站（算法相同，各自的列）
    //   機台需求工時：參照 MPS 工作表該段「機台需求工時」的 Total 列（由 dmBlockTotalFormulas() 算好傳進來）
    //   人機比：=機台數/人員數（六個月共用），Notes 寫「1 人顧 4 機」
    //   開機數：=ROUNDUP(機台需求工時/((工作天數+該站的加班天數)*機台每日可動工時),1)
    //   顧機人數：=ROUNDUP(開機數/人機比,1)
    //   需求顧機人數：=ROUNDUP(顧機人數*工作班別數/出勤率/作業效率,0)
    function addr(key, mm) { return numToCol(R.monthCol + mm) + rows[key]; }
    [
      { p: 'milling', hours: o && o.millingHours, ratio: o && o.millingRatio },
      { p: 'dipping', hours: o && o.dippingHours, ratio: o && o.dippingRatio }
    ].forEach(function (st) {
      var k = function (name) { return st.p + name; };     // 例如 millingHours、dippingRatio
      var ratio = st.ratio || {};
      var hasRatio = isN(ratio.machines) && isN(ratio.people) && ratio.people > 0;
      for (var mm = 0; mm < R.months; mm++) {
        if (st.hours && st.hours[mm]) put(k('Hours'), mm, { formula: st.hours[mm].formula, result: st.hours[mm].result }, '#,##0');
        if (hasRatio) {
          put(k('Ratio'), mm, {
            formula: fmtNum(ratio.machines) + '/' + fmtNum(ratio.people), result: ratio.machines / ratio.people
          }, '0.0#');
        }
        put(k('Machines'), mm, {
          formula: 'ROUNDUP(' + addr(k('Hours'), mm) + '/((' + addr('workDays', mm) + '+' + addr(k('OtDays'), mm) + ')*' +
            addr('machineMinutes', mm) + '),1)'
        }, '0.0');
        put(k('Tenders'), mm, {
          formula: 'ROUNDUP(' + addr(k('Machines'), mm) + '/' + addr(k('Ratio'), mm) + ',1)'
        }, '0.0');
        // 需求顧機人數 =ROUNDUP(顧機人數*工作班別數/出勤率/作業效率,0)
        put(k('Need'), mm, {
          formula: 'ROUNDUP(' + addr(k('Tenders'), mm) + '*' + addr('shifts', mm) + '/' + addr('attendance', mm) + '/' +
            addr('efficiency', mm) + ',0)'
        }, '0');
      }
      if (hasRatio) {
        ws.getCell(rows[k('Ratio')], R.noteCol).value = fmtNum(ratio.people) + ' 人顧 ' + fmtNum(ratio.machines) + ' 機';
      }
    });

    // 目檢站 需求人力（小數一位無條件進位）
    //   =ROUNDUP(需求工時/工作時間/工作班別數/(工作天數+加班天數 - 目檢)/出勤率/作業效率,1)
    // 需求人數 - 目檢
    //   =ROUNDUP(需求人力 - Milling*工作班別數+需求人力 - Dipping*工作班別數,1)
    // 總需求人力 =需求顧機人數 - Milling+需求顧機人數 - dipping+需求人數 - 目檢
    // 總現有人力 =現有人力 Milling+Dipping
    // 人力需求差異 =總現有人力-總需求人力
    for (m = 0; m < R.months; m++) {
      [['laborMinMilling', 'laborPplMilling'], ['laborMinDipping', 'laborPplDipping']].forEach(function (p) {
        put(p[1], m, {
          formula: 'ROUNDUP(' + addr(p[0], m) + '/' + addr('workMinutes', m) + '/' + addr('shifts', m) + '/(' +
            addr('workDays', m) + '+' + addr('laborOtDays', m) + ')/' + addr('attendance', m) + '/' + addr('efficiency', m) + ',1)'
        }, '0.0');
      });
      put('laborNeed', m, {
        formula: 'ROUNDUP(' + addr('laborPplMilling', m) + '*' + addr('shifts', m) + '+' +
          addr('laborPplDipping', m) + '*' + addr('shifts', m) + ',1)'
      }, '0.0');
      put('needTotal', m, { formula: addr('millingNeed', m) + '+' + addr('dippingNeed', m) + '+' + addr('laborNeed', m) }, '0.0');
      put('haveTotal', m, { formula: addr('haveMilling', m) + '+' + addr('haveDipping', m) }, '0.0');
      put('gap', m, { formula: addr('haveTotal', m) + '-' + addr('needTotal', m) }, '0.0');
    }

    // 列印：橫向、一頁寬
    Object.assign(ws.pageSetup, { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 });

    var wbAll = ws.workbook;
    if (wbAll) wbAll.calcProperties = Object.assign({}, wbAll.calcProperties, { fullCalcOnLoad: true });
    return { worksheet: ws, rows: rows, monthCol: R.monthCol, months: R.months, lastMonthCol: lastMonth };
  }

  /*
   * 匯出前的修正：關掉 Excel 的「公式不一致」綠色小三角。
   * 每一列的公式參照不同的列（例如 ='P003'!C2/'P003'!B2、='P003'!C5/'P003'!B5），Excel 會以為公式打錯而標綠色三角。
   * 做法：在每張工作表的 XML 加上 <ignoredErrors>（忽略公式不一致、數字存成文字等提示；#DIV/0! 這類錯誤照樣顯示）。
   * 任何一步不順利就原封不動匯出。
   */
  var IGNORED_ERRORS = '<ignoredErrors><ignoredError sqref="A1:XFD1048576" formula="1" formulaRange="1" ' +
    'numberStoredAsText="1" unlockedFormula="1"/></ignoredErrors>';
  // <ignoredErrors> 在工作表 XML 裡要放在這些元素之前（Excel 檢查元素順序）
  var IGNORED_BEFORE = ['<smartTags', '<drawing', '<legacyDrawing', '<legacyDrawingHF', '<drawingHF', '<picture',
    '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst', '</worksheet>'];

  function addIgnoredErrors(buffer) {
    try {
      var entries = readZip(buffer);
      if (!entries) return buffer;
      var changed = false;
      entries.forEach(function (e) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(e.name) || (e.method !== 8 && e.method !== 0)) return;
        var bytes = e.method === 8 ? inflateRaw(e.data, e.usize) : e.data;
        if (crc32(bytes) !== e.crc) return;
        var text = new TextDecoder('utf-8').decode(bytes);
        if (text.indexOf('<ignoredErrors') >= 0) return;
        var at = -1;
        for (var i = 0; i < IGNORED_BEFORE.length && at < 0; i++) {
          var j = text.indexOf(IGNORED_BEFORE[i]);
          if (j >= 0) at = j;
        }
        if (at < 0) return;
        var plain = new TextEncoder().encode(text.slice(0, at) + IGNORED_ERRORS + text.slice(at));
        e.data = plain;                                           // 改過的檔案不壓縮，直接存
        e.method = 0;
        e.crc = crc32(plain);
        e.usize = plain.length;
        changed = true;
      });
      return changed ? writeZip(entries) : buffer;
    } catch (err) {
      return buffer;
    }
  }

  function toBlob(workbook) {
    return workbook.xlsx.writeBuffer().then(function (buf) {
      var ab = buf instanceof ArrayBuffer ? buf : buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
      return new Blob([addIgnoredErrors(ab)], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });
  }

  var ExcelIO = {
    VERSION: 24,          // 程式版本，顯示在畫面右上角；每次改版加 1（index.html 的 ?v= 也要改）
    colToNum: colToNum,
    numToCol: numToCol,
    normalizeText: normalizeText,
    parseColumns: parseColumns,
    formatColumns: formatColumns,
    plainValue: plainValue,
    matchKey: matchKey,
    displayText: displayText,
    load: load,
    fixFontFlags: fixFontFlags,
    inflateRaw: inflateRaw,
    listMerges: listMerges,
    inspect: inspect,
    extent: extent,
    filter: filter,
    SECTION_GAP: SECTION_GAP,
    processKey: processKey,
    detectProcessBlocks: detectProcessBlocks,
    copySheet: copySheet,
    guessHeaderRow: guessHeaderRow,
    columnLabels: columnLabels,
    lookupKey: lookupKey,
    buildIndex: buildIndex,
    planLookup: planLookup,
    lookupFormula: lookupFormula,
    addLookupColumn: addLookupColumn,
    CT_HEADERS: CT_HEADERS,
    sectionDataRows: sectionDataRows,
    planCtColumns: planCtColumns,
    addCtColumns: addCtColumns,
    DEMAND_TITLES: DEMAND_TITLES,
    planDemandBlocks: planDemandBlocks,
    addDemandBlocks: addDemandBlocks,
    findHeaderColumn: findHeaderColumn,
    planDemand: planDemand,
    addDemandColumns: addDemandColumns,
    planMpsTotals: planMpsTotals,
    addMpsTotals: addMpsTotals,
    findSections: findSections,
    planSectionTotals: planSectionTotals,
    addSectionTotals: addSectionTotals,
    planSummary: planSummary,
    addSummaryRows: addSummaryRows,
    planRma: planRma,
    addRmaRows: addRmaRows,
    MANPOWER_ITEMS: MANPOWER_ITEMS,
    manpowerRows: manpowerRows,
    guessManpowerRows: guessManpowerRows,
    guessFirstMonthCol: guessFirstMonthCol,
    labelsAbove: labelsAbove,
    readManpower: readManpower,
    REPORT_SHEET: REPORT_SHEET,
    REPORT: REPORT,
    REPORT_ROWS: REPORT_ROWS,
    addReportSheet: addReportSheet,
    REPORT_LINKS: REPORT_LINKS,
    REPORT_OFFLINE: REPORT_OFFLINE,
    YO_LIMIT_DEFAULT: YO_LIMIT_DEFAULT,
    RMA_SHEET: RMA_SHEET,
    RMA_LAYOUT: RMA_LAYOUT,
    RMA_ROWS: RMA_ROWS,
    addRmaManpowerSheet: addRmaManpowerSheet,
    readRmaQuantities: readRmaQuantities,
    REPORT_SECOND_OFFSET: REPORT_SECOND_OFFSET,
    findReportLinks: findReportLinks,
    toNumber: toNumber,
    DM_REPORT: DM_REPORT,
    DM_ROWS: DM_ROWS,
    addDmReportSheet: addDmReportSheet,
    dmMpsFormulas: dmMpsFormulas,
    dmBlockTotalFormulas: dmBlockTotalFormulas,
    addIgnoredErrors: addIgnoredErrors,
    toBlob: toBlob
  };

  global.ExcelIO = ExcelIO;
  if (typeof module !== 'undefined' && module.exports) module.exports = ExcelIO;
})(typeof window !== 'undefined' ? window : globalThis);
