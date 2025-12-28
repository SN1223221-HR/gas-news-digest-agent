/**
 * Main Logic & Entry Points
 */
const NEWS_AGENT = Object.freeze({ NAME: 'News Agent Pro', VERSION: '8.1.0' });

const CONFIG = Object.freeze({
  SHEET_NAMES: Object.freeze({ CONFIG: 'Config', DB: 'DB', LOGS: 'Logs' }),
  LIMITS: Object.freeze({ ITEMS_PER_TOPIC_MAIL: 3 }),
  LOCK: Object.freeze({ ACQUIRE_TIMEOUT_MS: 30000 }),
  DB_COLUMNS: Object.freeze({
    TIMESTAMP: 1, KEYWORD: 2, TITLE: 3, LINK: 4, DATE: 5, SOURCE: 6,
    SENT_FLAG: 7, RATING: 8, COMMENT: 9, IS_READ: 10, CAMPAIGN_TAG: 11,
    SUMMARY: 12, FIRST_DATA_ROW: 2
  })
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚡ News Agent')
    .addItem('📥 今すぐ収集 (Crawl)', 'manualCrawl')
    .addItem('📮 今すぐ配信 (Send Mail)', 'manualSend')
    .addSeparator()
    .addItem('🌐 Webダッシュボードを開く', 'openDashboardUrl')
    .addToUi();
}

function manualCrawl() {
  executeSafely('Manual Crawl', app => {
    const count = app.crawl();
    SpreadsheetApp.getActiveSpreadsheet().toast(`収集完了: ${count}件`, NEWS_AGENT.NAME);
  });
}

function manualSend() {
  executeSafely('Manual Send', app => {
    const count = app.forceSendMail();
    SpreadsheetApp.getActiveSpreadsheet().toast(`配信完了: ${count}件`, NEWS_AGENT.NAME);
  });
}

function executeSafely(context, taskFunction) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK.ACQUIRE_TIMEOUT_MS)) return;
  try {
    const logger = new LoggerService();
    const app = new NewsApp(logger);
    taskFunction(app);
  } catch (e) { console.error(e); } finally { lock.releaseLock(); }
}

class NewsApp {
  constructor(logger) {
    this.repo = new SheetRepository();
    this.logger = logger || new LoggerService();
    this.rss = new RssService();
    this.mailer = new MailService();
  }

  crawl() {
    const settings = this.repo.loadConfig();
    const keywords = this.repo.getKeywords();
    const regions = (settings.Region || 'JP').split(',');
    const lang = settings.Language || 'ja';
    let newItems = [];
    const urlCache = new UrlCacheService(this.repo);

    keywords.forEach(kw => {
      regions.forEach(region => {
        try {
          const items = this.rss.fetch(kw.trim(), region.trim(), lang);
          items.forEach(item => {
            if (!urlCache.exists(item.link)) {
              item.keyword = kw;
              newItems.push(item);
              urlCache.add(item.link);
            }
          });
          Utilities.sleep(500);
        } catch(e) {}
      });
    });

    if (newItems.length > 0) this.repo.saveArticles(newItems);
    return newItems.length;
  }

  forceSendMail() {
    const unsent = this.repo.getUnsentArticles();
    if (unsent.length === 0) return 0;

    const grouped = {};
    unsent.forEach(art => {
      const kw = art.keyword || 'その他';
      if (!grouped[kw]) grouped[kw] = [];
      if (grouped[kw].length < CONFIG.LIMITS.ITEMS_PER_TOPIC_MAIL) {
        grouped[kw].push(art);
      }
    });

    const settings = this.repo.loadConfig();
    this.mailer.sendBriefing(grouped, settings);
    
    this.repo.markAsSent(unsent);
    return unsent.length;
  }
}
