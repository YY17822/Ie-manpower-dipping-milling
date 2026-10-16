/*
 * sw.js — 離線快取。
 * 自己的檔案：有網路時一定向伺服器確認有沒有新版，並更新快取；沒網路時用快取。
 *             （所以改了程式重新整理就會生效）
 * ExcelJS（網路版）：抓過一次就一直用快取。
 * 每次新增檔案時，把檔名加進 SHELL，並把 VERSION 加 1。
 * 快取名稱開頭是專案代號，新版啟用時只刪「自己專案、但不是目前版本」的舊快取，
 * 同一個網址底下其他專案的快取不會被刪。
 */
importScripts('js/config.js');                  // 專案代號（PROJECT_ID）寫在這個檔案
var CONFIG = self.APP_CONFIG;
var VERSION = 9;                                  // 每次改版加 1（和 excel-io.js 的 VERSION、index.html 的 ?v= 一樣）
var CACHE = CONFIG.PROJECT_ID + '-v' + VERSION;    // 例如 ie-manpower-dipping-milling-v1
var EXCELJS_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
var SHELL = [
  './',
  'index.html',
  'manifest.webmanifest',
  'js/config.js',
  'css/style.css',
  'js/excel-io.js',
  'js/app.js',
  'js/ui.js',
  'js/steps/step1-mps-filter.js',
  'js/steps/step2-mps-totals.js',
  'js/steps/step3-p003-import.js',
  'js/steps/step4-ct-lookup.js',
  'js/steps/step5-demand-hours.js',
  'js/steps/step6-manpower.js',
  'js/steps/step7-report.js',
  'js/steps/step8-rma-manpower.js',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/icon-maskable-512.png'
];

// 只認「這個專案代號-v數字」的快取（別的專案、或代號只是開頭相同的專案都不算）
var OWN_CACHE = new RegExp('^' + CONFIG.PROJECT_ID.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-v\\d+$');
var LEGACY_CACHE = CONFIG.LEGACY_CACHE_PATTERN ? new RegExp(CONFIG.LEGACY_CACHE_PATTERN) : null;
function isOldOwnCache(name) {
  if (name === CACHE) return false;
  return OWN_CACHE.test(name) || !!(LEGACY_CACHE && LEGACY_CACHE.test(name));
}

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE).then(function (cache) {
      return cache.addAll(SHELL).then(function () {
        // ExcelJS：資料夾裡有就用資料夾的，沒有就存一份網路版；兩個都失敗也不影響安裝
        return cache.add('vendor/exceljs.min.js').catch(function () {
          return cache.add(new Request(EXCELJS_CDN, { mode: 'cors' })).catch(function () {});
        });
      });
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(isOldOwnCache).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);

  if (req.url === EXCELJS_CDN) {
    event.respondWith(
      caches.match(req).then(function (hit) {
        return hit || fetch(req).then(function (res) {
          if (res.ok) { var copy = res.clone(); caches.open(CACHE).then(function (c) { c.put(req, copy); }); }
          return res;
        });
      })
    );
    return;
  }

  if (url.origin !== self.location.origin) return;

  // cache: 'no-cache' = 不直接用瀏覽器自己留的舊檔，先問伺服器有沒有更新
  var fresh = req.mode === 'navigate' ? fetch(req) : fetch(req.url, { cache: 'no-cache' });
  event.respondWith(
    fresh.then(function (res) {
      if (res.ok) { var copy = res.clone(); caches.open(CACHE).then(function (c) { c.put(req, copy); }); }
      return res;
    }).catch(function () {
      return caches.match(req, { ignoreSearch: true }).then(function (hit) {
        if (hit) return hit;
        if (req.mode === 'navigate') return caches.match('index.html');
        return Response.error();
      });
    })
  );
});
