/*
 * config.js — 這個專案的身分設定，網頁和 sw.js（離線快取）都讀這一個檔案。
 *
 * 複製這個專案做成新專案時，先改這裡的 PROJECT_ID（每個專案都要不同），
 * 瀏覽器裡的快取和記住的設定才會和原本的專案分開，不會互相讀到、蓋掉或刪掉。
 *   快取名稱       PROJECT_ID + '-v' + 版本號，例如 ie-manpower-dipping-milling-v1（版本號在 sw.js）
 *   記住的設定     PROJECT_ID + '.' + 名稱，例如 ie-manpower-dipping-milling.step1
 *
 * LEGACY_*：這個專案以前用的舊名稱，新版第一次打開時會把舊設定搬過來、把舊快取刪掉。
 *           新專案用不到，複製之後請改成 null。
 */
(function (g) {
  'use strict';
  g.APP_CONFIG = {
    PROJECT_ID: 'ie-manpower-dipping-milling',
    LEGACY_STORE_PREFIX: null,      // 新專案，沒有舊的設定名稱
    LEGACY_CACHE_PATTERN: null      // 新專案，沒有舊的快取名稱
  };
})(typeof self !== 'undefined' ? self : this);
