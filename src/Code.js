/**
 * ==================================================
 * Enterprise News Aggregator for GAS
 * Version: 3.0.0 (Pro Edition)
 * ==================================================
 *
 *  - Google News RSS をキーワード × 複数リージョンでクロール
 *  - 直近URLをシート＋CacheService(MD5キー) で重複排除
 *  - DBシートに保存し、指定時間にまとめてHTMLメール配信
 *  - Configシートで DeliveryHours / Region / Language / UserName /
 *    MailTo / MailCc / MailBcc を制御
 *  - ログシートに INFO/WARN/ERROR を蓄積（cleanupLogs でメンテ）
 *
 *  想定シート構成:
 *   - シート名: Config, DB, Logs
 *
 *  Config シート:
 *   - A1: "SearchKeywords"
 *   - A2〜: キーワード（例: 法改正, 生成AI, Google Apps Script, ...）
 *   - C1: "SettingKey" / D1: "SettingValue"
 *   - C2: DeliveryHours,  D2: "7, 12, 19"
 *   - C3: Region,         D3: "JP, US, GB"
 *   - C4: Language,       D4: "ja"
 *   - C5: UserName,       D5: "ユーザー"
 *   - C6: MailTo,         D6: "you@example.com, other@example.com"
 *   - C7: MailCc,         D7: "cc@example.com"
 *   - C8: MailBcc,        D8: "bcc@example.com"
 */

/**
 * @typedef {Object} AppSettings
 * @property {string} [DeliveryHours]  Comma separated hours (e.g. "7,12,19")
 * @property {string} [Region]         Comma separated country codes (e.g. "JP,US,GB")
 * @property {string} [Language]       Language code (e.g. "ja")
 * @property {string} [UserName]       Recipient display name
 * @property {string} [MailTo]         To addresses (comma separated)
 * @property {string} [MailCc]         Cc addresses (comma separated)
 * @property {string} [MailBcc]        Bcc addresses (comma separated)
 */

/**
 * @typedef {Object} ArticleRow
 * @property {number} rowIndex
 * @property {Date}   timestamp
 * @property {string} keyword
 * @property {string} title
 * @property {string} link
 * @property {string} date
 * @property {string} source
 * @property {boolean|string} isSent
 */

/**
 * @typedef {Object} ArticlePayload
 * @property {string} keyword
 * @property {string} title
 * @property {string} link
 * @property {string} date
 * @property {string} source
 */

// ==================================================
//  Global Config
// ==================================================

const NEWS_AGENT = Object.freeze({
  NAME: 'News Agent Pro',
  VERSION: '3.0.0'
});

const CONFIG = Object.freeze({
  SHEET_NAMES: Object.freeze({
    CONFIG: 'Config',
    DB: 'DB',
    LOGS: 'Logs'
  }),
  CACHE: Object.freeze({
    TTL_SECONDS: 21600, // 6h
    PREFIX: 'news_url_'
  }),
  RETRY: Object.freeze({
    MAX_ATTEMPTS: 3,
    BASE_DELAY_MS: 1000
  }),
  UI: Object.freeze({
    TOAST_SECONDS: 3
  }),
  LOCK: Object.freeze({
    ACQUIRE_TIMEOUT_MS: 30000
  }),
  LOG: Object.freeze({
    MAX_ROWS: 2000 // cleanupLogs() で使用
  }),
  DB_COLUMNS: Object.freeze({
    // 1-based index
    TIMESTAMP: 1,
    KEYWORD: 2,
    TITLE: 3,
    LINK: 4,
    DATE: 5,
    SOURCE: 6,
    SENT_FLAG: 7,
    FIRST_DATA_ROW: 2
  }),
  CONFIG_COLUMNS: Object.freeze({
    KEY_COL: 3,   // "SettingKey"
    VALUE_COL: 4, // "SettingValue"
    FIRST_DATA_ROW: 2
  })
});

// ==================================================
//  Utility
// ==================================================

/**
 * CacheService 用のキー生成ヘルパー
 * - URL 全体を key にすると 250文字制限に引っかかるため
 * - URL を MD5 ハッシュにして短い key にする
 * @param {string} url
 * @returns {string}
 */
function createCacheKeyFromUrl(url) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url);
  let hash = '';
  for (let i = 0; i < digest.length; i++) {
    const byte = digest[i] & 0xFF;
    hash += ('0' + byte.toString(16)).slice(-2);
  }
  return CONFIG.CACHE.PREFIX + hash;
}

/**
 * "1, 2, 3" -> [1, 2, 3]
 * @param {string|undefined} str
 * @param {function(number): boolean} [predicate]
 * @returns {number[]}
 */
function parseIntList(str, predicate) {
  if (!str) return [];
  const fn = typeof predicate === 'function' ? predicate : function () { return true; };
  return str
    .toString()
    .split(',')
    .map(function (s) { return parseInt(s.trim(), 10); })
    .filter(function (n) { return !isNaN(n) && fn(n); });
}

/**
 * "JP,US, GB" -> ["JP","US","GB"]
 * @param {string|undefined} str
 * @param {string[]} [fallback]
 * @returns {string[]}
 */
function parseRegionList(str, fallback) {
  const regions = (str || '')
    .toString()
    .split(',')
    .map(function (r) { return r.trim().toUpperCase(); })
    .filter(function (r) { return r.length > 0; });
  return regions.length > 0 ? regions : (fallback || ['JP']);
}

/**
 * 軽量な実行時間プロファイル
 * @template T
 * @param {string} label
 * @param {function():T} fn
 * @returns {T}
 */
function profile(label, fn) {
  const start = Date.now();
  try {
    return fn();
  } finally {
    const elapsed = Date.now() - start;
    console.log('[PROFILE]', label, '-', elapsed + 'ms');
  }
}

// ==================================================
//  UI & Entry Points
// ==================================================

/**
 * メニュー作成
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ News Agent')
    .addItem('📥 今すぐ収集 (Crawl)', 'manualCrawl')
    .addItem('📮 今すぐ配信 (Send Mail)', 'manualSend')
    .addSeparator()
    .addItem('🛠 接続テスト', 'testConnection')
    .addSeparator()
    .addItem('🧹 ログを整理 (cleanupLogs)', 'cleanupLogs')
    .addToUi();
}

/**
 * トリガー: 収集タスク
 */
function crawlTask() {
  executeSafely('Crawl Task', function (app) {
    app.crawl();
  });
}

/**
 * トリガー: 配信タスク
 */
function checkAndSendMailTask() {
  executeSafely('Mail Task', function (app) {
    app.checkAndSendMail();
  });
}

/**
 * 手動実行: 収集
 */
function manualCrawl() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '収集を開始します...',
    NEWS_AGENT.NAME,
    CONFIG.UI.TOAST_SECONDS
  );

  executeSafely('Manual Crawl', function (app) {
    const count = profile('crawl()', function () {
      return app.crawl();
    });
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '収集完了: ' + count + '件の新規記事',
      NEWS_AGENT.NAME,
      CONFIG.UI.TOAST_SECONDS
    );
  });
}

/**
 * 手動実行: 配信
 */
function manualSend() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '配信処理を開始します...',
    NEWS_AGENT.NAME,
    CONFIG.UI.TOAST_SECONDS
  );

  executeSafely('Manual Send', function (app) {
    const sentCount = profile('forceSendMail()', function () {
      return app.forceSendMail();
    });
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '配信完了: ' + sentCount + '件送信',
      NEWS_AGENT.NAME,
      CONFIG.UI.TOAST_SECONDS
    );
  });
}

/**
 * テスト用: 接続確認
 */
function testConnection() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'RSS接続テスト中...',
    NEWS_AGENT.NAME,
    CONFIG.UI.TOAST_SECONDS
  );

  try {
    const rss = new RssService();
    const result = rss.fetch('Google', 'JP', 'ja');
    const msg = result.length > 0
      ? '成功: ' + result.length + '件取得'
      : '失敗: 0件';
    SpreadsheetApp.getActiveSpreadsheet().toast(
      msg,
      '接続テスト結果',
      5
    );
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'エラー: ' + e.message,
      '接続テスト失敗',
      5
    );
  }
}

/**
 * テスト用: 即時実行デバッグ
 */
function testImmediateRun() {
  console.log('=== [TEST] テスト実行開始 ===');
  executeSafely('TestImmediateRun', function (app) {
    app.crawl();
    app.forceSendMail();
  });
  console.log('=== [TEST] テスト実行完了 ===');
}

/**
 * ログシートのクリーンアップ
 *  - 最新 MAX_ROWS 行だけ残して上を削除
 */
function cleanupLogs() {
  const logger = new LoggerService();
  logger.trimOldRows(CONFIG.LOG.MAX_ROWS);
}

/**
 * ロック & ログ付きセーフ実行ラッパー
 * @param {string} context
 * @param {function(NewsApp):void} taskFunction
 */
function executeSafely(context, taskFunction) {
  const lock = LockService.getScriptLock();
  const logger = new LoggerService();

  if (!lock.tryLock(CONFIG.LOCK.ACQUIRE_TIMEOUT_MS)) {
    console.warn('[LOCK-SKIP]', context);
    logger.warn(context, 'Process skipped (locked)');
    return;
  }

  try {
    console.log('[START]', context, 'v' + NEWS_AGENT.VERSION);
    const app = new NewsApp(logger);
    taskFunction(app);
    console.log('[END]', context);
  } catch (e) {
    console.error('[ERROR]', context, e);
    logger.error(context, e && e.stack ? e.stack : String(e));
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'エラー発生: ' + e.message,
        'Error - ' + context,
        5
      );
    } catch (uiError) {
      // バックグラウンド等で UI が取れない場合は無視
    }
  } finally {
    lock.releaseLock();
  }
}

// ==================================================
//  Application Core
// ==================================================

/**
 * アプリケーションサービス
 */
class NewsApp {
  /**
   * @param {LoggerService} [logger]
   */
  constructor(logger) {
    this.repo = new SheetRepository();
    this.logger = logger || new LoggerService();
    this.rss = new RssService();
    this.mailer = new MailService(this.logger);
    /** @type {AppSettings} */
    this.settings = this.repo.loadConfig();
  }

  /**
   * ニュースをクロールして DB に追加
   * @returns {number} 追加された記事数
   */
  crawl() {
    const keywords = this.repo.getKeywords();
    if (keywords.length === 0) {
      this.logger.info('Crawl', 'No keywords found. Skipping.');
      return 0;
    }

    const regions = parseRegionList(this.settings.Region, ['JP']);
    const lang = this.settings.Language || 'ja';

    let totalNewArticles = 0;
    const urlCache = new UrlCacheService(this.repo, this.logger);

    keywords.forEach(function (keyword) {
      regions.forEach(function (region) {
        try {
          const items = this.rss.fetch(keyword, region, lang);
          /** @type {ArticlePayload[]} */
          const newItems = [];

          items.forEach(function (item) {
            if (!urlCache.exists(item.link)) {
              item.keyword = keyword;
              newItems.push(item);
              urlCache.add(item.link);
            }
          });

          if (newItems.length > 0) {
            this.repo.saveArticles(newItems);
            totalNewArticles += newItems.length;
            this.logger.info(
              'Crawl',
              'Saved ' + newItems.length + ' articles for "' + keyword +
                '" (' + region + ')'
            );
          }

          // 呼び出し制限対策：キーワード×地域ごとに少し待つ
          Utilities.sleep(1500);
        } catch (e) {
          this.logger.error(
            'Crawl(' + keyword + '/' + region + ')',
            e && e.stack ? e.stack : String(e)
          );
        }
      }, this);
    }, this);

    if (totalNewArticles === 0) {
      this.logger.info('Crawl', 'No new articles found.');
    } else {
      this.logger.info('Crawl', 'Total ' + totalNewArticles + ' new articles saved.');
    }
    return totalNewArticles;
  }

  /**
   * 配信時間帯かどうかを確認し、該当すればメール送信
   */
  checkAndSendMail() {
    const currentHour = new Date().getHours();
    const deliveryHours = parseIntList(
      this.settings.DeliveryHours,
      function (h) { return h >= 0 && h <= 23; }
    );

    if (deliveryHours.length === 0) {
      // 設定が空の場合は [7] とみなす
      deliveryHours.push(7);
    }

    if (deliveryHours.indexOf(currentHour) === -1) {
      console.log(
        '[Mail-SKIP]',
        'Current hour:',
        currentHour,
        'Allowed:',
        deliveryHours.join(',')
      );
      return;
    }

    this._processMailSending();
  }

  /**
   * 配信時間のチェックなしで強制的に送信
   * @returns {number} 送信した件数
   */
  forceSendMail() {
    return this._processMailSending();
  }

  /**
   * 未送信記事をまとめて送信
   * @returns {number}
   * @private
   */
  _processMailSending() {
    const unsentArticles = this.repo.getUnsentArticles();
    if (unsentArticles.length === 0) {
      this.logger.info('Mail', 'No unsent articles. Skipping.');
      return 0;
    }

    try {
      this.mailer.sendDailyReport(unsentArticles, this.settings);
      this.repo.markAsSent(unsentArticles);
      this.logger.info('Mail', 'Sent ' + unsentArticles.length + ' articles.');
      return unsentArticles.length;
    } catch (e) {
      this.logger.error('Mail', 'Failed to send: ' + (e && e.stack ? e.stack : String(e)));
      throw e;
    }
  }
}

// ==================================================
//  Sheets / Repository
// ==================================================

/**
 * シートアクセスの責務をまとめたリポジトリ
 */
class SheetRepository {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.dbSheet = this._getSheetOrThrow(CONFIG.SHEET_NAMES.DB);
    this.configSheet = this._getSheetOrThrow(CONFIG.SHEET_NAMES.CONFIG);
  }

  /**
   * @param {string} name
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   * @private
   */
  _getSheetOrThrow(name) {
    const sheet = this.ss.getSheetByName(name);
    if (!sheet) {
      throw new Error('Sheet "' + name + '" not found.');
    }
    return sheet;
  }

  /**
   * Config シートから設定を読み取る
   * @returns {AppSettings}
   */
  loadConfig() {
    const colKey = CONFIG.CONFIG_COLUMNS.KEY_COL;
    const colVal = CONFIG.CONFIG_COLUMNS.VALUE_COL;
    const startRow = CONFIG.CONFIG_COLUMNS.FIRST_DATA_ROW;

    const lastRow = this.configSheet.getLastRow();
    if (lastRow < startRow) return /** @type {AppSettings} */ ({});

    const numRows = lastRow - startRow + 1;
    const range = this.configSheet.getRange(startRow, colKey, numRows, 2);
    const values = range.getValues();

    /** @type {Object.<string,string>} */
    const result = {};

    values.forEach(function (row) {
      const key = row[0];
      const val = row[1];
      if (key) {
        result[String(key)] = val != null ? String(val) : '';
      }
    });

    return /** @type {AppSettings} */ (result);
  }

  /**
   * 検索キーワード一覧を取得
   * @returns {string[]}
   */
  getKeywords() {
    const firstDataRow = 2;
    const lastRow = this.configSheet.getLastRow();
    if (lastRow < firstDataRow) return [];

    const numRows = lastRow - firstDataRow + 1;
    const values = this.configSheet.getRange(firstDataRow, 1, numRows, 1).getValues();

    return values
      .map(function (row) { return row[0]; })
      .map(function (v) { return v != null ? String(v).trim() : ''; })
      .filter(function (v) { return v.length > 0; });
  }

  /**
   * 直近 N 件の URL を Set として取得
   * @returns {Set<string>}
   */
  getAllUrls() {
    const lastRow = this.dbSheet.getLastRow();
    const firstDataRow = CONFIG.DB_COLUMNS.FIRST_DATA_ROW;

    if (lastRow < firstDataRow) return new Set();

    const limit = 2000;
    const startRow = Math.max(firstDataRow, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    const colLink = CONFIG.DB_COLUMNS.LINK;

    const values = this.dbSheet.getRange(startRow, colLink, numRows, 1).getValues();
    const urls = new Set();

    values.forEach(function (row) {
      const link = row[0];
      if (link) {
        urls.add(String(link));
      }
    });

    return urls;
  }

  /**
   * 記事を DB シートに保存
   * @param {ArticlePayload[]} articles
   */
  saveArticles(articles) {
    if (!articles || articles.length === 0) return;

    const now = new Date();
    const firstCol = CONFIG.DB_COLUMNS.TIMESTAMP;

    const rows = articles.map(function (a) {
      return [
        now,
        a.keyword,
        a.title,
        a.link,
        a.date,
        a.source,
        false
      ];
    });

    const startRow = this.dbSheet.getLastRow() + 1;
    this.dbSheet.getRange(startRow, firstCol, rows.length, rows[0].length).setValues(rows);
  }

  /**
   * 未送信のレコードを取得
   * @returns {ArticleRow[]}
   */
  getUnsentArticles() {
    const lastRow = this.dbSheet.getLastRow();
    const firstDataRow = CONFIG.DB_COLUMNS.FIRST_DATA_ROW;

    if (lastRow < firstDataRow) return [];

    const numRows = lastRow - firstDataRow + 1;
    const firstCol = CONFIG.DB_COLUMNS.TIMESTAMP;
    const numCols = CONFIG.DB_COLUMNS.SENT_FLAG;

    const data = this.dbSheet.getRange(firstDataRow, firstCol, numRows, numCols).getValues();

    /** @type {ArticleRow[]} */
    const rows = [];

    for (let i = 0; i < data.length; i++) {
      const rowIndex = firstDataRow + i;
      const row = data[i];

      const isSentCell = row[CONFIG.DB_COLUMNS.SENT_FLAG - 1];
      const isSent = (isSentCell === true || isSentCell === 'TRUE');

      if (!isSent) {
        rows.push({
          rowIndex: rowIndex,
          timestamp: row[0],
          keyword: row[1],
          title: row[2],
          link: row[3],
          date: row[4],
          source: row[5],
          isSent: isSentCell
        });
      }
    }

    return rows;
  }

  /**
   * 記事をまとめて「送信済み」に更新
   * @param {ArticleRow[]} articles
   */
  markAsSent(articles) {
    if (!articles || articles.length === 0) return;

    const sheet = this.dbSheet;
    const colSent = CONFIG.DB_COLUMNS.SENT_FLAG;

    const rows = articles.map(function (a) { return a.rowIndex; });
    const minRow = Math.min.apply(null, rows);
    const maxRow = Math.max.apply(null, rows);
    const numRows = maxRow - minRow + 1;

    const range = sheet.getRange(minRow, colSent, numRows, 1);
    const values = range.getValues(); // [[val], [val], ...]

    const rowSet = new Set(rows);

    for (let i = 0; i < values.length; i++) {
      const rowIndex = minRow + i;
      if (rowSet.has(rowIndex)) {
        values[i][0] = true;
      }
    }

    range.setValues(values);
  }
}

// ==================================================
//  Caching
// ==================================================

/**
 * URL 重複チェック用キャッシュ
 * - シート（直近2000件）＋ CacheService(MD5キー) の二段構え
 */
class UrlCacheService {
  /**
   * @param {SheetRepository} repo
   * @param {LoggerService} logger
   */
  constructor(repo, logger) {
    this.cache = CacheService.getScriptCache();
    this.repo = repo;
    this.logger = logger;
    this.memorySet = this.repo.getAllUrls();
  }

  /**
   * @param {string} url
   * @returns {boolean}
   */
  exists(url) {
    if (this.memorySet.has(url)) return true;

    try {
      const key = createCacheKeyFromUrl(url);
      const cached = this.cache.get(key);
      return cached !== null;
    } catch (e) {
      this.logger.warn('UrlCache', 'Cache get failed: ' + String(e));
      return false;
    }
  }

  /**
   * @param {string} url
   */
  add(url) {
    this.memorySet.add(url);
    try {
      const key = createCacheKeyFromUrl(url);
      this.cache.put(key, '1', CONFIG.CACHE.TTL_SECONDS);
    } catch (e) {
      this.logger.warn('UrlCache', 'Cache put failed: ' + String(e));
    }
  }
}

// ==================================================
//  Logging
// ==================================================

class LoggerService {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.LOGS);
    if (!this.sheet) {
      this.sheet = this.ss.insertSheet(CONFIG.SHEET_NAMES.LOGS);
      this.sheet.appendRow(['Timestamp', 'Level', 'Context', 'Message']);
      this.sheet.setColumnWidth(4, 400);
    }
  }

  /**
   * @param {string} ctx
   * @param {string} msg
   */
  info(ctx, msg) {
    this._log('INFO', ctx, msg);
  }

  /**
   * @param {string} ctx
   * @param {string} msg
   */
  warn(ctx, msg) {
    this._log('WARN', ctx, msg);
  }

  /**
   * @param {string} ctx
   * @param {string} msg
   */
  error(ctx, msg) {
    this._log('ERROR', ctx, msg);
  }

  /**
   * @param {'INFO'|'WARN'|'ERROR'} level
   * @param {string} ctx
   * @param {string} msg
   * @private
   */
  _log(level, ctx, msg) {
    const MAX_LEN = 2000;
    const text = msg != null ? String(msg) : '';
    const truncated = text.length > MAX_LEN
      ? text.substring(0, MAX_LEN) + '... (truncated)'
      : text;
    this.sheet.appendRow([new Date(), level, ctx, truncated]);
  }

  /**
   * ログシートの古い行を削除
   * @param {number} keepRows  ヘッダを除いて残す行数
   */
  trimOldRows(keepRows) {
    const lastRow = this.sheet.getLastRow();
    if (lastRow <= keepRows + 1) return; // ヘッダー含め keepRows+1 行以下なら何もしない

    const deleteCount = lastRow - keepRows;
    this.sheet.deleteRows(2, deleteCount - 1); // ヘッダー除いて削除
  }
}

// ==================================================
//  External Services (RSS / Mail)
// ==================================================

class RssService {
  /**
   * @param {string} keyword
   * @param {string} region
   * @param {string} lang
   * @returns {ArticlePayload[]}
   */
  fetch(keyword, region, lang) {
    const encodedKey = encodeURIComponent(keyword);
    const ceid = region + ':' + lang;
    const url =
      'https://news.google.com/rss/search?q=' +
      encodedKey +
      '&hl=' + lang +
      '&gl=' + region +
      '&ceid=' + ceid;

    return this._fetchWithRetry(url);
  }

  /**
   * @param {string} url
   * @returns {ArticlePayload[]}
   * @private
   */
  _fetchWithRetry(url) {
    let attempts = 0;
    while (attempts < CONFIG.RETRY.MAX_ATTEMPTS) {
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const code = response.getResponseCode();
        if (code === 200) {
          return this._parseXml(response.getContentText());
        }
        throw new Error('Status ' + code);
      } catch (e) {
        attempts++;
        if (attempts >= CONFIG.RETRY.MAX_ATTEMPTS) {
          throw e;
        }
        Utilities.sleep(CONFIG.RETRY.BASE_DELAY_MS * attempts);
      }
    }
    return [];
  }

  /**
   * RSS XML をパースして直近24時間分だけを返す
   * @param {string} xml
   * @returns {ArticlePayload[]}
   * @private
   */
  _parseXml(xml) {
    try {
      const document = XmlService.parse(xml);
      const root = document.getRootElement();
      if (!root) return [];

      const channel = root.getChild('channel');
      if (!channel) return [];

      const items = channel.getChildren('item') || [];
      const oneDayAgo = Date.now() - 24 * 60 * 60 * 1000;

      /** @type {ArticlePayload[]} */
      const results = [];

      items.forEach(function (item) {
        if (!item) return;

        const titleEl = item.getChild('title');
        const linkEl = item.getChild('link');
        const pubDateEl = item.getChild('pubDate');

        if (!titleEl || !linkEl || !pubDateEl) return;

        const pubDate = new Date(pubDateEl.getText());
        if (isNaN(pubDate.getTime()) || pubDate.getTime() < oneDayAgo) {
          return;
        }

        const sourceEl = item.getChild('source');

        results.push({
          title: titleEl.getText(),
          link: linkEl.getText(),
          date: Utilities.formatDate(pubDate, 'Asia/Tokyo', 'MM/dd HH:mm'),
          source: sourceEl ? sourceEl.getText() : 'Google News'
        });
      });

      return results;
    } catch (e) {
      console.error('XML Parse Error', e);
      return [];
    }
  }
}

class MailService {
  /**
   * @param {LoggerService} logger
   */
  constructor(logger) {
    this.logger = logger;
  }

  /**
   * @param {ArticleRow[]} articles
   * @param {AppSettings} settings
   */
  sendDailyReport(articles, settings) {
    const defaultTo = Session.getActiveUser().getEmail();
    const to = (settings.MailTo || defaultTo).toString();
    const cc = settings.MailCc ? settings.MailCc.toString() : '';
    const bcc = settings.MailBcc ? settings.MailBcc.toString() : '';

    const dateStr = Utilities.formatDate(
      new Date(),
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm'
    );
    const userName = settings.UserName || 'User';

    // 絵文字を削除して文字化けを防止
    const subject = '【News】Briefing for ' + userName + ' (' + dateStr + ')';

    const htmlBody = this._generateHtml(articles, userName, dateStr, settings);

    GmailApp.sendEmail(to, subject, 'HTML対応メーラーでご覧ください', {
      htmlBody: htmlBody,
      name: NEWS_AGENT.NAME,
      cc: cc,
      bcc: bcc
    });
  }

  /**
   * @param {ArticleRow[]} articles
   * @param {string} userName
   * @param {string} dateStr
   * @param {AppSettings} settings
   * @returns {string}
   * @private
   */
  _generateHtml(articles, userName, dateStr, settings) {
    const grouped = this._groupByKeyword(articles);
    let listHtml = '';

    Object.keys(grouped).forEach(function (keyword) {
      const items = grouped[keyword];
      listHtml += [
        '<div style="margin-top: 20px;">',
        '  <h3 style="background: #e8f0fe; color: #1967d2; padding: 8px 12px; border-radius: 6px; font-size: 14px; margin-bottom: 10px; display: inline-block;">',
        '    # ' + keyword,
        '  </h3>',
        items.map(function (a) {
          return [
            '  <div style="padding: 8px 0; border-bottom: 1px solid #f1f3f4;">',
            '    <a href="' + a.link + '" style="text-decoration: none; color: #202124; font-weight: 600; font-size: 15px; display: block; line-height: 1.4;">',
            '      ' + a.title,
            '    </a>',
            '    <div style="color: #5f6368; font-size: 12px; margin-top: 4px;">',
            '      <span style="color: #1a73e8;">' + a.source + '</span> &bull; ' + a.date,
            '    </div>',
            '  </div>'
          ].join('\n');
        }).join('\n'),
        '</div>'
      ].join('\n');
    });

    const regionLabel = settings.Region || 'N/A';
    const deliveryLabel = settings.DeliveryHours || 'N/A';

    return [
      '<!DOCTYPE html>',
      '<html>',
      '<body style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; background-color: #f6f6f6; padding: 20px; margin: 0;">',
      '  <div style="max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">',
      '    <div style="background: #1a73e8; padding: 20px; color: white;">',
      '      <h1 style="margin: 0; font-size: 20px;">News Briefing</h1>',
      '      <p style="margin: 5px 0 0; opacity: 0.9; font-size: 13px;">' +
        dateStr + ' | ' + articles.length + ' New Articles</p>',
      '    </div>',
      '    <div style="padding: 20px;">',
      '      <p style="color: #333; margin-top: 0;">Hi <strong>' + userName + '</strong>,<br>Here are the latest updates based on your interests.</p>',
      listHtml,
      '    </div>',
      '    <div style="background: #f8f9fa; padding: 15px 20px; text-align: center; font-size: 11px; color: #9aa0a6; border-top: 1px solid #eee;">',
      '      Region: ' + regionLabel + ' | Delivery: ' + deliveryLabel,
      '    </div>',
      '  </div>',
      '</body>',
      '</html>'
    ].join('\n');
  }

  /**
   * @param {ArticleRow[]} articles
   * @returns {Object.<string, ArticleRow[]>}
   * @private
   */
  _groupByKeyword(articles) {
    return articles.reduce(function (acc, curr) {
      const key = curr.keyword || 'Others';
      if (!acc[key]) acc[key] = [];
      acc[key].push(curr);
      return acc;
    }, /** @type {Object.<string, ArticleRow[]>} */ ({}));
  }
}
