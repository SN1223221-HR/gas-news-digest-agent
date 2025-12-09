/**
 * ==================================================
 * Enterprise News Aggregator for GAS
 * Version: 4.0.0 (Personalization Edition)
 * ==================================================
 */

// ==================================================
//  Global Config
// ==================================================

const NEWS_AGENT = Object.freeze({
  NAME: 'News Agent Pro',
  VERSION: '4.0.0'
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
    MAX_ROWS: 2000
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
    RATING: 8, // ★追加: 評価(1-5)
    FIRST_DATA_ROW: 2
  })
});

// ==================================================
//  UI Entry Points
// ==================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ News Agent')
    .addItem('📰 ニュースリーダー (Sidebar)', 'showNewsSidebar')
    .addItem('📊 トレンド・評価分析 (Analytics)', 'showAnalyticsDialog')
    .addSeparator()
    .addItem('📥 今すぐ収集 (Crawl)', 'manualCrawl')
    .addItem('📮 今すぐ配信 (Send Mail)', 'manualSend')
    .addToUi();
}

function showNewsSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('📰 News Feed Pro')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAnalyticsDialog() {
  const html = HtmlService.createTemplateFromFile('AnalyticsDialog')
    .evaluate()
    .setWidth(900)
    .setHeight(700)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Trend & Domain Analytics');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================================================
//  API for Sidebar (Rating)
// ==================================================

/**
 * 記事の評価を更新する
 * @param {string} link - ユニークキー代わり
 * @param {number} rating - 1 to 5
 */
function updateArticleRating(link, rating) {
  const repo = new SheetRepository();
  repo.updateRating(link, rating);
  return { success: true };
}

/**
 * サイドバー用データ取得 (評価付き)
 */
function getNewsFeedData() {
  const repo = new SheetRepository();
  const sheet = repo.dbSheet;
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return JSON.stringify([]);

  const limit = 100;
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;

  // 8列(Rating)まで取得
  const values = sheet.getRange(startRow, 1, numRows, 8).getValues();
  
  const feed = values.map(function(r) {
    return {
      timestamp: r[0],
      keyword: r[1],
      title: r[2],
      link: r[3],
      date: r[4],
      source: r[5],
      rating: r[7] // Column 8 is index 7
    };
  }).reverse();

  return JSON.stringify(feed);
}

// ==================================================
//  Domain Filtering Logic
// ==================================================

class DomainReputationService {
  constructor(repo) {
    this.repo = repo;
    // 過去の全データからドメインごとの平均スコアを計算（重い処理なのでキャッシュすべきだが今回は簡易実装）
    this.stats = this._calculateStats();
  }

  /**
   * ドメインが「ブロック対象（低評価）」か判定
   * 基準: 評価数が2件以上 かつ 平均評価が 2.0 未満
   */
  isBlocked(url) {
    const domain = this._extractDomain(url);
    if (!domain) return false;
    
    const stat = this.stats[domain];
    if (!stat) return false;

    // Strict Rule: 2回以上評価されていて、平均が2.0未満ならブロック
    if (stat.count >= 2 && (stat.sum / stat.count) < 2.0) {
      console.log(`[Blocked] Domain: ${domain}, Avg: ${(stat.sum / stat.count).toFixed(1)}`);
      return true;
    }
    return false;
  }

  _calculateStats() {
    const data = this.repo.getAllRatings(); // [[link, rating], ...]
    const stats = {}; // { "example.com": { sum: 10, count: 3 } }

    data.forEach(row => {
      const link = row[0];
      const rating = row[1];
      if (link && rating) {
        const domain = this._extractDomain(link);
        if (domain) {
          if (!stats[domain]) stats[domain] = { sum: 0, count: 0 };
          stats[domain].sum += Number(rating);
          stats[domain].count += 1;
        }
      }
    });
    return stats;
  }

  _extractDomain(url) {
    try {
      return new URL(url).hostname;
    } catch (e) {
      return null;
    }
  }
}

// ==================================================
//  Core Tasks
// ==================================================

function manualCrawl() {
  SpreadsheetApp.getActiveSpreadsheet().toast('収集を開始します...', NEWS_AGENT.NAME);
  executeSafely('Manual Crawl', app => {
    const count = app.crawl();
    SpreadsheetApp.getActiveSpreadsheet().toast(`収集完了: ${count}件 (低評価サイトは除外済)`, NEWS_AGENT.NAME);
  });
}

function manualSend() {
  executeSafely('Manual Send', app => app.forceSendMail());
}

function cleanupLogs() { const logger = new LoggerService(); logger.trimOldRows(CONFIG.LOG.MAX_ROWS); }
function testConnection() { /* 省略（既存のままでOK） */ }

function executeSafely(context, taskFunction) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK.ACQUIRE_TIMEOUT_MS)) return;
  try {
    const app = new NewsApp();
    taskFunction(app);
  } catch (e) {
    console.error(e);
  } finally {
    lock.releaseLock();
  }
}

// ==================================================
//  Application Classes
// ==================================================

class NewsApp {
  constructor() {
    this.repo = new SheetRepository();
    this.logger = new LoggerService();
    this.rss = new RssService();
    this.mailer = new MailService();
    this.reputation = new DomainReputationService(this.repo); // 評判システム
  }

  crawl() {
    const settings = this.repo.loadConfig();
    const keywords = this.repo.getKeywords();
    if (keywords.length === 0) return 0;

    const regions = parseRegionList(settings.Region, ['JP']);
    const lang = settings.Language || 'ja';
    let totalNewArticles = 0;
    const urlCache = new UrlCacheService(this.repo);

    keywords.forEach(keyword => {
      regions.forEach(region => {
        try {
          const items = this.rss.fetch(keyword, region, lang);
          const newItems = [];
          
          items.forEach(item => {
            // 重複チェック AND ドメイン評判チェック
            if (!urlCache.exists(item.link)) {
              if (!this.reputation.isBlocked(item.link)) {
                item.keyword = keyword;
                newItems.push(item);
                urlCache.add(item.link);
              }
            }
          });

          if (newItems.length > 0) {
            this.repo.saveArticles(newItems);
            totalNewArticles += newItems.length;
          }
          Utilities.sleep(1000);
        } catch (e) {
          this.logger.error('Crawl Error', e.message);
        }
      });
    });
    return totalNewArticles;
  }

  forceSendMail() {
    const settings = this.repo.loadConfig();
    const unsent = this.repo.getUnsentArticles();
    if (unsent.length === 0) return 0;
    this.mailer.sendDailyReport(unsent, settings);
    this.repo.markAsSent(unsent);
    return unsent.length;
  }
}

class SheetRepository {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.dbSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.DB) || this.ss.insertSheet(CONFIG.SHEET_NAMES.DB);
    this.configSheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIG);
  }

  loadConfig() {
    const data = this.configSheet.getDataRange().getValues();
    const result = {};
    for(let i=1; i<data.length; i++) {
      if(data[i][2]) result[String(data[i][2])] = data[i][3];
    }
    return result;
  }

  getKeywords() {
    const last = this.configSheet.getLastRow();
    return last < 2 ? [] : this.configSheet.getRange(2, 1, last-1, 1).getValues().flat().filter(String);
  }

  getAllUrls() {
    const last = this.dbSheet.getLastRow();
    if(last < 2) return new Set();
    const limit = 3000; // チェック範囲拡大
    const start = Math.max(2, last - limit + 1);
    const values = this.dbSheet.getRange(start, CONFIG.DB_COLUMNS.LINK, last-start+1, 1).getValues();
    const s = new Set();
    values.forEach(v => { if(v[0]) s.add(String(v[0])); });
    return s;
  }

  // ドメイン分析用にURLとRatingのペアを取得
  getAllRatings() {
    const last = this.dbSheet.getLastRow();
    if(last < 2) return [];
    // URLとRating列を取得
    return this.dbSheet.getRange(2, CONFIG.DB_COLUMNS.LINK, last-1, CONFIG.DB_COLUMNS.RATING - CONFIG.DB_COLUMNS.LINK + 1).getValues()
      .map(row => [row[0], row[row.length-1]]); // [Link, Rating]
  }

  saveArticles(articles) {
    if (!articles || articles.length === 0) return;
    const now = new Date();
    const rows = articles.map(a => [
      now, a.keyword, a.title, a.link, a.date, a.source, false, '' // Rating初期値は空
    ]);
    const start = this.dbSheet.getLastRow() + 1;
    this.dbSheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
  }

  updateRating(targetLink, rating) {
    const last = this.dbSheet.getLastRow();
    if(last < 2) return;
    // NOTE: パフォーマンスのため直近500件から検索（古すぎる記事への評価は無視）
    const limit = 500;
    const start = Math.max(2, last - limit + 1);
    const links = this.dbSheet.getRange(start, CONFIG.DB_COLUMNS.LINK, last-start+1, 1).getValues().flat();
    
    const index = links.indexOf(targetLink);
    if (index !== -1) {
      const rowIndex = start + index;
      this.dbSheet.getRange(rowIndex, CONFIG.DB_COLUMNS.RATING).setValue(rating);
    }
  }

  getUnsentArticles() {
    const last = this.dbSheet.getLastRow();
    if(last < 2) return [];
    const data = this.dbSheet.getRange(2, 1, last-1, CONFIG.DB_COLUMNS.SENT_FLAG).getValues();
    const rows = [];
    data.forEach((r, i) => {
      if(!r[6]) rows.push({ rowIndex: i + 2, title: r[2], link: r[3], source: r[5] });
    });
    return rows;
  }

  markAsSent(articles) {
    const sheet = this.dbSheet;
    articles.forEach(a => sheet.getRange(a.rowIndex, CONFIG.DB_COLUMNS.SENT_FLAG).setValue(true));
  }
}

// --- 以下、Logger, UrlCache, RssService, MailService は簡易版（変更なしだが動作に必要） ---
class LoggerService {
  constructor() { this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.LOGS) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.SHEET_NAMES.LOGS); }
  error(c, m) { this.sheet.appendRow([new Date(), 'ERROR', c, m]); }
  trimOldRows(k) { if(this.sheet.getLastRow() > k) this.sheet.deleteRows(2, this.sheet.getLastRow()-k); }
}
class UrlCacheService {
  constructor(r) { this.cache = CacheService.getScriptCache(); this.mem = r.getAllUrls(); }
  exists(u) { return this.mem.has(u) || this.cache.get(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, u).toString()) !== null; }
  add(u) { this.mem.add(u); this.cache.put(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, u).toString(), '1', 21600); }
}
class RssService {
  fetch(k, r, l) {
    try {
      const u = `https://news.google.com/rss/search?q=${encodeURIComponent(k)}&hl=${l}&gl=${r}&ceid=${r}:${l}`;
      const xml = UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText();
      const doc = XmlService.parse(xml);
      return doc.getRootElement().getChild('channel').getChildren('item').map(i => ({
        title: i.getChild('title').getText(),
        link: i.getChild('link').getText(),
        date: Utilities.formatDate(new Date(i.getChild('pubDate').getText()), 'Asia/Tokyo', 'MM/dd HH:mm'),
        source: i.getChild('source') ? i.getChild('source').getText() : 'News'
      }));
    } catch(e) { return []; }
  }
}
class MailService {
  sendDailyReport(a, s) { GmailApp.sendEmail(s.MailTo || Session.getActiveUser().getEmail(), 'News', `${a.length} articles`); }
}
function parseRegionList(s, f) { return (s||'').split(',').map(r=>r.trim().toUpperCase()).filter(r=>r) || f; }
