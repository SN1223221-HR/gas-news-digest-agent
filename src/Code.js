/**
 * ==================================================
 * Enterprise News Aggregator for GAS
 * Version: 2.1.0
 * ==================================================
 */

// グローバル定数
const CONFIG = {
  SHEET_NAMES: {
    CONFIG: 'Config',
    DB: 'DB',
    LOGS: 'Logs'
  },
  CACHE: {
    TTL: 21600, // 6時間
    PREFIX: 'news_url_'
  },
  RETRY: {
    MAX_ATTEMPTS: 3,
    DELAY_MS: 1000
  },
  UI: {
    TOAST_SECONDS: 3
  }
};

/* ==================================================
 * 1. UI & Triggers (Entry Points)
 * ================================================== */

/** メニュー作成 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ News Agent')
    .addItem('📥 今すぐ収集 (Crawl)', 'manualCrawl')
    .addItem('📮 今すぐ配信 (Send Mail)', 'manualSend')
    .addSeparator()
    .addItem('🛠 接続テスト', 'testConnection')
    .addToUi();
}

/** トリガー: 収集タスク */
function crawlTask() {
  executeSafely('Crawl Task', () => {
    const app = new NewsApp();
    app.crawl();
  });
}

/** トリガー: 配信タスク */
function checkAndSendMailTask() {
  executeSafely('Mail Task', () => {
    const app = new NewsApp();
    app.checkAndSendMail();
  });
}

/** 手動実行: 収集 */
function manualCrawl() {
  SpreadsheetApp.getActiveSpreadsheet().toast('収集を開始します...', 'News Agent', CONFIG.UI.TOAST_SECONDS);
  
  executeSafely('Manual Crawl', () => {
    const app = new NewsApp();
    const count = app.crawl();
    SpreadsheetApp.getActiveSpreadsheet().toast(`収集完了: ${count}件の新規記事`, 'News Agent', CONFIG.UI.TOAST_SECONDS);
  });
}

/** 手動実行: 配信 */
function manualSend() {
  SpreadsheetApp.getActiveSpreadsheet().toast('配信処理を開始します...', 'News Agent', CONFIG.UI.TOAST_SECONDS);
  
  executeSafely('Manual Send', () => {
    const app = new NewsApp();
    const sentCount = app.forceSendMail();
    SpreadsheetApp.getActiveSpreadsheet().toast(`配信完了: ${sentCount}件送信`, 'News Agent', CONFIG.UI.TOAST_SECONDS);
  });
}

/** テスト用: 接続確認 */
function testConnection() {
  SpreadsheetApp.getActiveSpreadsheet().toast('RSS接続テスト中...', 'News Agent', CONFIG.UI.TOAST_SECONDS);
  
  try {
    const rssService = new RssService();
    const result = rssService.fetch('Google', 'JP', 'ja');
    const msg = result.length > 0 ? `成功: ${result.length}件取得` : '失敗: 0件';
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '接続テスト結果', 5);
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`エラー: ${e.message}`, '接続テスト失敗', 5);
  }
}

/** テスト用: 即時実行デバッグ */
function testImmediateRun() {
  console.log("=== [TEST] テスト実行開始 ===");
  const app = new NewsApp();
  app.crawl();
  app.forceSendMail();
  console.log("=== [TEST] テスト実行完了 ===");
}

/**
 * 安全実行ラッパー
 */
function executeSafely(context, taskFunction) {
  const lock = LockService.getScriptLock();
  const logger = new LoggerService();
  
  if (lock.tryLock(30000)) {
    try {
      console.log(`[START] ${context}`);
      taskFunction();
      console.log(`[END] ${context}`);
    } catch (e) {
      console.error(`[ERROR] ${context}:`, e);
      logger.error(context, e.message);
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(`エラー発生: ${e.message}`, 'Error', 5);
      } catch (uiError) {} 
    } finally {
      lock.releaseLock();
    }
  } else {
    console.warn(`[SKIP] ${context} - Locked`);
    logger.warn(context, 'Process skipped (Locked)');
  }
}

/* ==================================================
 * 2. Application Logic
 * ================================================== */

class NewsApp {
  constructor() {
    this.repo = new SheetRepository();
    this.rss = new RssService();
    this.mailer = new MailService();
    this.logger = new LoggerService();
    this.settings = this.repo.loadConfig();
  }

  crawl() {
    const keywords = this.repo.getKeywords();
    if (keywords.length === 0) return 0;

    const region = this.settings.Region || 'JP';
    const lang = this.settings.Language || 'ja';
    let totalNewArticles = 0;
    const urlCache = new UrlCacheService(this.repo);

    keywords.forEach(keyword => {
      try {
        const items = this.rss.fetch(keyword, region, lang);
        const newItems = [];

        items.forEach(item => {
          if (!urlCache.exists(item.link)) {
            item.keyword = keyword;
            newItems.push(item);
            urlCache.add(item.link);
          }
        });

        if (newItems.length > 0) {
          this.repo.saveArticles(newItems);
          totalNewArticles += newItems.length;
        }
        Utilities.sleep(1500);
      } catch (e) {
        this.logger.error(`Crawl(${keyword})`, e.message);
      }
    });

    if (totalNewArticles > 0) {
      this.logger.info('Crawl', `Saved ${totalNewArticles} new articles.`);
    }
    return totalNewArticles;
  }

  checkAndSendMail() {
    const currentHour = new Date().getHours();
    const deliveryHours = this._parseHours(this.settings.DeliveryHours);

    if (!deliveryHours.includes(currentHour)) {
      console.log(`Skipping mail. Current: ${currentHour}`);
      return;
    }
    this._processMailSending();
  }

  forceSendMail() {
    return this._processMailSending();
  }

  _processMailSending() {
    const unsentArticles = this.repo.getUnsentArticles();
    if (unsentArticles.length === 0) return 0;

    try {
      this.mailer.sendDailyReport(unsentArticles, this.settings);
      this.repo.markAsSent(unsentArticles);
      this.logger.info('Mail', `Sent ${unsentArticles.length} articles.`);
      return unsentArticles.length;
    } catch (e) {
      this.logger.error('Mail', `Failed to send: ${e.message}`);
      throw e;
    }
  }

  _parseHours(str) {
    if (!str) return [7];
    return str.toString().split(',').map(h => parseInt(h.trim(), 10));
  }
}

/* ==================================================
 * 3. Infrastructure
 * ================================================== */

class SheetRepository {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.dbSheet = this._getSheetOrThrow(CONFIG.SHEET_NAMES.DB);
    this.configSheet = this._getSheetOrThrow(CONFIG.SHEET_NAMES.CONFIG);
  }

  _getSheetOrThrow(name) {
    const sheet = this.ss.getSheetByName(name);
    if (!sheet) throw new Error(`Sheet "${name}" not found.`);
    return sheet;
  }

  loadConfig() {
    const lastRow = this.configSheet.getLastRow();
    if (lastRow < 2) return {};
    const data = this.configSheet.getRange(2, 3, lastRow - 1, 2).getValues();
    return data.reduce((acc, [key, val]) => {
      if (key) acc[key] = val;
      return acc;
    }, {});
  }

  getKeywords() {
    const lastRow = this.configSheet.getLastRow();
    if (lastRow < 2) return [];
    return this.configSheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .flat().filter(k => k && String(k).trim() !== "");
  }

  getAllUrls() {
    const lastRow = this.dbSheet.getLastRow();
    if (lastRow < 2) return new Set();
    const limit = 2000;
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    const urls = this.dbSheet.getRange(startRow, 4, numRows, 1).getValues().flat();
    return new Set(urls);
  }

  saveArticles(articles) {
    const timestamp = new Date();
    const rows = articles.map(a => [
      timestamp, a.keyword, a.title, a.link, a.date, a.source, false
    ]);
    this.dbSheet.getRange(this.dbSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  getUnsentArticles() {
    const lastRow = this.dbSheet.getLastRow();
    if (lastRow < 2) return [];
    const data = this.dbSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    return data.map((row, i) => ({
      rowIndex: i + 2,
      timestamp: row[0],
      keyword: row[1],
      title: row[2],
      link: row[3],
      date: row[4],
      source: row[5],
      isSent: row[6]
    })).filter(a => a.isSent === false || a.isSent === 'FALSE');
  }

  markAsSent(articles) {
    if (articles.length === 0) return;
    const sheet = this.dbSheet;
    articles.forEach(a => {
      sheet.getRange(a.rowIndex, 7).setValue(true);
    });
  }
}

class UrlCacheService {
  constructor(repo) {
    this.cache = CacheService.getScriptCache();
    this.repo = repo;
    this.memorySet = this.repo.getAllUrls(); 
  }

  exists(url) {
    if (this.memorySet.has(url)) return true;
    const cached = this.cache.get(CONFIG.CACHE.PREFIX + url);
    return cached !== null;
  }

  add(url) {
    this.memorySet.add(url);
    try {
      this.cache.put(CONFIG.CACHE.PREFIX + url, "1", CONFIG.CACHE.TTL);
    } catch (e) {
      console.warn("Cache put failed", e);
    }
  }
}

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
  info(ctx, msg) { this._log('INFO', ctx, msg); }
  warn(ctx, msg) { this._log('WARN', ctx, msg); }
  error(ctx, msg) { this._log('ERROR', ctx, msg); }
  _log(level, ctx, msg) {
    this.sheet.appendRow([new Date(), level, ctx, msg]);
  }
}

class RssService {
  fetch(keyword, region, lang) {
    const encodedKey = encodeURIComponent(keyword);
    const ceid = `${region}:${lang}`;
    const url = `https://news.google.com/rss/search?q=${encodedKey}&hl=${lang}&gl=${region}&ceid=${ceid}`;
    return this._fetchWithRetry(url, keyword);
  }

  _fetchWithRetry(url, keyword) {
    let attempts = 0;
    while (attempts < CONFIG.RETRY.MAX_ATTEMPTS) {
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          return this._parseXml(response.getContentText());
        }
        throw new Error(`Status ${response.getResponseCode()}`);
      } catch (e) {
        attempts++;
        if (attempts >= CONFIG.RETRY.MAX_ATTEMPTS) throw e;
        Utilities.sleep(CONFIG.RETRY.DELAY_MS * attempts);
      }
    }
    return [];
  }

  _parseXml(xml) {
    try {
      const document = XmlService.parse(xml);
      const items = document.getRootElement().getChild("channel").getChildren("item");
      const oneDayAgo = new Date().getTime() - (24 * 60 * 60 * 1000);
      return items.map(item => {
        const pubDateStr = item.getChild("pubDate").getText();
        const pubDate = new Date(pubDateStr);
        if (pubDate.getTime() < oneDayAgo) return null;
        return {
          title: item.getChild("title").getText(),
          link: item.getChild("link").getText(),
          date: Utilities.formatDate(pubDate, "Asia/Tokyo", "MM/dd HH:mm"),
          source: item.getChild("source") ? item.getChild("source").getText() : "Google News"
        };
      }).filter(item => item !== null);
    } catch (e) {
      console.error("XML Parse Error", e);
      return [];
    }
  }
}

class MailService {
  sendDailyReport(articles, settings) {
    const recipient = Session.getActiveUser().getEmail();
    const dateStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    const userName = settings.UserName || "User";
    
    const subject = `【News】Briefing for ${userName} (${dateStr})`;

    const htmlBody = this._generateHtml(articles, userName, dateStr, settings);

    GmailApp.sendEmail(recipient, subject, "HTML対応メーラーでご覧ください", {
      htmlBody: htmlBody,
      name: 'News Agent Pro'
    });
  }

  _generateHtml(articles, userName, dateStr, settings) {
    const grouped = this._groupByKeyword(articles);
    let listHtml = '';

    Object.keys(grouped).forEach(keyword => {
      listHtml += `
        <div style="margin-top: 20px;">
          <h3 style="background: #e8f0fe; color: #1967d2; padding: 8px 12px; border-radius: 6px; font-size: 14px; margin-bottom: 10px; display: inline-block;">
            # ${keyword}
          </h3>
          ${grouped[keyword].map(a => `
            <div style="padding: 8px 0; border-bottom: 1px solid #f1f3f4;">
              <a href="${a.link}" style="text-decoration: none; color: #202124; font-weight: 600; font-size: 15px; display: block; line-height: 1.4;">
                ${a.title}
              </a>
              <div style="color: #5f6368; font-size: 12px; margin-top: 4px;">
                <span style="color: #1a73e8;">${a.source}</span> &bull; ${a.date}
              </div>
            </div>
          `).join('')}
        </div>
      `;
    });

    return `
      <!DOCTYPE html>
      <html>
      <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f6f6f6; padding: 20px; margin: 0;">
        <div style="max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
          <div style="background: #1a73e8; padding: 20px; color: white;">
            <h1 style="margin: 0; font-size: 20px;">News Briefing</h1>
            <p style="margin: 5px 0 0; opacity: 0.9; font-size: 13px;">${dateStr} | ${articles.length} New Articles</p>
          </div>
          <div style="padding: 20px;">
            <p style="color: #333; margin-top: 0;">Hi <strong>${userName}</strong>,<br>Here are the latest updates based on your interests.</p>
            ${listHtml}
          </div>
          <div style="background: #f8f9fa; padding: 15px 20px; text-align: center; font-size: 11px; color: #9aa0a6; border-top: 1px solid #eee;">
            Region: ${settings.Region} | Delivery: ${settings.DeliveryHours}
          </div>
        </div>
      </body>
      </html>
    `;
  }

  _groupByKeyword(articles) {
    return articles.reduce((acc, curr) => {
      (acc[curr.keyword] = acc[curr.keyword] || []).push(curr);
      return acc;
    }, {});
  }
}
