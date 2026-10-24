/*
 * config.js — 這個專案的身分設定，網頁和 sw.js（離線快取）都讀這一個檔案。
 *
 * 複製這個專案做成新專案時，先改這裡的 PROJECT_ID（每個專案都要不同），
 * 瀏覽器裡的快取和記住的設定才會和原本的專案分開，不會互相讀到、蓋掉或刪掉。
 *   快取名稱       PROJECT_ID + '-v' + 版本號，例如 ie-manpower-dipping-milling-v23（版本號在 sw.js）
 *   記住的設定     PROJECT_ID + '.' + 名稱，例如 ie-manpower-dipping-milling.step1
 *
 * SUPABASE_*：共用設定的資料庫，複製成新專案時不用改（同一個 Supabase、同一個 schema），只要 PROJECT_ID 不同就不會混在一起。
 *   名稱照自架 Supabase 的 .env 取，兩邊找得到同一個名字：
 *     SUPABASE_PUBLIC_URL        Supabase 網址                            ＝ .env 的 SUPABASE_PUBLIC_URL
 *     SUPABASE_PUBLISHABLE_KEY   公開鑰匙（舊式名稱：ANON_KEY）           ＝ .env 的 SUPABASE_PUBLISHABLE_KEY
 *     SUPABASE_SCHEMA            這個程式用的 schema                      ＝ .env 的 PGRST_DB_SCHEMAS 清單裡的其中一個
 *   換資料庫（雲端 → 本機自架 → 公司主機）時只改前兩個，程式不用動：
 *     Supabase 雲端   https://xxxx.supabase.co        Settings → API Keys 的 Publishable key
 *     Docker 自架     http://<主機 IP 或名稱>:8000    .env 的 SUPABASE_PUBLISHABLE_KEY（舊版 .env 沒有這行時用 ANON_KEY）
 *   只能放 publishable key；SUPABASE_SECRET_KEY（舊式名稱：SERVICE_ROLE_KEY）絕對不能放進來。
 *
 * LEGACY_*：這個專案以前用的舊名稱，新版第一次打開時會把舊設定搬過來、把舊快取刪掉。
 *           這個專案一開始就用 ie-manpower-dipping-milling 當開頭，沒有舊名稱，所以是 null。
 */
(function (g) {
  'use strict';
  g.APP_CONFIG = {
    PROJECT_ID: 'ie-manpower-dipping-milling',
    LEGACY_STORE_PREFIX: null,      // 沒有舊的設定名稱
    LEGACY_CACHE_PATTERN: null,     // 沒有舊的快取名稱

    // 共用設定的資料庫（Supabase）。8 個人力需求 PWA 共用同一個 schema，用 PROJECT_ID 分開。
    // SUPABASE_PUBLIC_URL 留空字串 = 不用 Supabase，設定只存在這台電腦。
    // 本機 Docker 自架（~/docker-lab/supabase-project/.env）：
    SUPABASE_PUBLIC_URL: 'http://localhost:8000',
    SUPABASE_PUBLISHABLE_KEY: 'sb_publishable_gzgqcIVPviJVCFZvCfrlXw_5EPFugBU',
    SUPABASE_SCHEMA: 'ie_manpower'                 // .env 的 PGRST_DB_SCHEMAS 要包含它
  };
})(typeof self !== 'undefined' ? self : this);
